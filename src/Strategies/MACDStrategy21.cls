VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACDStrategy21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' This is a variant of MACDStrategy7 which uses dummy orders to control the overall
' entries.
'
' First tried using StopStrategy6 for exits. It gave poor results but maybe the
' stop strategy could be improved.
'
' Next, tried using StopStrategy5 but with per-trade money management. No better!
'
' Next, tried using stoplimit orders for entries. Better than the basic MACDStrategy7!
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IStrategy

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum Colours
    NoColour = 0
    Green                   ' indicates confirmed strength
    Amber                   ' indicates potential strength
    Pink                    ' indicates potential weakness
    Red                     ' indicates confirmed weakness
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "MACDStrategy21"

Private Const MessagePrefix                         As String = ">>>>>>>>>>>>>>>> "

'@================================================================================
' Member variables
'@================================================================================

Private mContext As TradingContext
Private WithEvents mPositionManagerDummy As PositionManager
Attribute mPositionManagerDummy.VB_VarHelpID = -1
Private mParams As Parameters

Private mOrderContextBuyDummy As OrderContext
Private mOrderContextSellDummy As OrderContext
Private WithEvents mBracketOrderBuyDummy As BracketOrder
Attribute mBracketOrderBuyDummy.VB_VarHelpID = -1
Private WithEvents mBracketOrderSellDummy As BracketOrder
Attribute mBracketOrderSellDummy.VB_VarHelpID = -1

Private mOrderContextBuy As OrderContext
Private mOrderContextSell As OrderContext
Private WithEvents mBracketOrderBuy As BracketOrder
Attribute mBracketOrderBuy.VB_VarHelpID = -1
Private WithEvents mBracketOrderSell As BracketOrder
Attribute mBracketOrderSell.VB_VarHelpID = -1

Private mStopStrategyFactory As IPosnMgmtStrtgyFactory
Private mScalingStrategyFactory As IPosnMgmtStrtgyFactory

Private mMaxInitialStop As Double
Private mStopBreakoutThresholdTicks As Integer
Private mEntryBreakoutThresholdTicks As Integer
Private mEntryLimitOffsetTicks As Integer

Private mPrevColour As Colours
Private mCurrColour As Colours

' Normal timeframe studies
Private mBars As Bars
Private mTimeframe As Timeframe
Private mBoll As BollingerBands
Private mSwing As Swing
Private mMACD As MACD

Private mTickSize As Double
Private mTickValue As Double
Private mLastMinute As Date
Private mLastHalfHour As Date
Private mTimestamp As Date
Private mBarNumber As Long

Private mNoTradeThisBar  As Boolean

Private mRiskUnitPercent As Double
Private mMaxTradeSize As Long

Private mBackstopMAOffset As Integer

Private mSwingToBreakeven As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mPrevColour = NoColour
mCurrColour = NoColour

End Sub

Private Sub Class_Terminate()
Debug.Print "MACDStrategy21 terminated"
End Sub

'@================================================================================
' IStrategy Interface Members
'@================================================================================

Private Property Get IStrategy_DefaultParameters() As Parameters
Const ProcName As String = "IStrategy_DefaultParameters"
On Error GoTo Err

Dim lParams As New Parameters
lParams.SetParameterValue ParamBarLength, BarLength
lParams.SetParameterValue ParamBarUnit, BarUnit
lParams.SetParameterValue ParamIncludeBarsOutsideSession, IncludeBarsOutsideSession
lParams.SetParameterValue ParamBollingerPeriods, BollingerPeriods
lParams.SetParameterValue ParamBollingerMovingAverageType, BollingerMovingAverageType
lParams.SetParameterValue ParamMinimumSwingTicks, MinimumSwingTicks
lParams.SetParameterValue ParamEntryLimitOffsetTicks, EntryLimitOffsetTicks
lParams.SetParameterValue ParamMaxInitialStopTicks, MaxInitialStopTicks
lParams.SetParameterValue ParamEntryBreakoutThresholdTicks, EntryBreakoutThresholdTicks
lParams.SetParameterValue ParamSwingToBreakevenTicks, SwingToBreakevenTicks
lParams.SetParameterValue ParamRiskUnitPercent, RiskUnitPercent
lParams.SetParameterValue ParamMaxTradeSize, MaxTradeSize
lParams.SetParameterValue ParamBollingerCentreBandWidthTicks, BollCentreBandwidthTicks
lParams.SetParameterValue ParamBollingerEdgeBandWidthTicks, BollEdgeBandwidthTicks

Set IStrategy_DefaultParameters = lParams

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Sub IStrategy_Initialise( _
                ByVal pContext As InitialisationContext, _
                ByVal pParams As Parameters)
Const ProcName As String = "IStrategy_Initialise"
On Error GoTo Err

processParameters pParams

Dim lContract As IContract
Set lContract = pContext.Contract
mTickSize = lContract.TickSize
mTickValue = lContract.TickValue
mLastMinute = pContext.Session.CurrentSessionEndTime - (1 / 1440)
mLastHalfHour = pContext.Session.CurrentSessionEndTime - (30 / 1440)

mMaxInitialStop = mParams.GetParameterValue(ParamMaxInitialStopTicks) * mTickSize
mSwingToBreakeven = mParams.GetParameterValue(ParamSwingToBreakevenTicks) * mTickSize

'--------------------------------------------------------------------------------
' Studies for the normal timeframe
'--------------------------------------------------------------------------------
Set mTimeframe = pContext.AddTimeframe(GetTimePeriod(mParams.GetParameterValue(ParamBarLength), _
                                                TimePeriodUnitsFromString(mParams.GetParameterValue(ParamBarUnit))), _
                                500, _
                                mParams.GetParameterValue(ParamIncludeBarsOutsideSession), _
                                True)

ReDim linputvaluenames(0) As String

Dim params As Parameters

Set params = New Parameters
params.SetParameterValue "Minimum swing (ticks)", mParams.GetParameterValue(ParamMinimumSwingTicks)
params.SetParameterValue "Include implicit swing points", "Y"
linputvaluenames(0) = "$default"
Set mSwing = pContext.AddStudy("Swing", _
                            mTimeframe.BarStudy, _
                            linputvaluenames, _
                            mParams.GetParameterValue(ParamIncludeBarsOutsideSession), _
                            params, _
                            3, _
                            , _
                            True)

Set params = New Parameters
params.SetParameterValue "Periods", mParams.GetParameterValue(ParamBollingerPeriods)
params.SetParameterValue "Standard deviations", 2
params.SetParameterValue "Mov avg type", mParams.GetParameterValue(ParamBollingerMovingAverageType)
params.SetParameterValue "Centre band width", mParams.GetParameterValue(ParamBollingerCentreBandWidthTicks) * mTickSize
params.SetParameterValue "Edge band width", mParams.GetParameterValue(ParamBollingerEdgeBandWidthTicks) * mTickSize
linputvaluenames(0) = "$default"
Set mBoll = pContext.AddStudy("Bollinger Bands", _
                        mTimeframe.BarStudy, _
                        linputvaluenames, _
                        mParams.GetParameterValue(ParamIncludeBarsOutsideSession), _
                        params, _
                        3, _
                        , _
                        True)

Set params = New Parameters
params.SetParameterValue "Short Periods", 5
params.SetParameterValue "Mov avg type", "SMA"
params.SetParameterValue "Long Periods", 34
params.SetParameterValue "Smoothing Periods", 5
linputvaluenames(0) = "$default"
Set mMACD = pContext.AddStudy("MACD", _
                        mTimeframe.BarStudy, _
                        linputvaluenames, _
                        mParams.GetParameterValue(ParamIncludeBarsOutsideSession), _
                        params, _
                        10, _
                        , _
                        True)

'--------------------------------------------------------------------------------
' Other initialisation
'--------------------------------------------------------------------------------

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStrategy_Name() As String
IStrategy_Name = "MACDStrategy21"
End Property

Private Sub IStrategy_NotifyTick(ByVal pType As TickTypes)
Const ProcName As String = "IStrategy_NotifyTick"
On Error GoTo Err

Static sFirstBarProcessed As Boolean

If Not mContext.ReceivedFirstBidAndAsk Then Exit Sub
If pType <> TickTypes.TickTypeTrade Then Exit Sub

If mBars.TickVolume <> 1 Then
    If sFirstBarProcessed Then Exit Sub
    sFirstBarProcessed = True
End If

mTimestamp = mContext.TimeStamp

If mTimestamp >= mLastMinute Then
    mContext.LogTradeReason "Closing any outstanding position"
    mContext.PositionManager.ClosePositions
    Exit Sub
End If

mBarNumber = mBars.Bar.BarNumber
mNoTradeThisBar = False

checkTradingOpportunity

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

' occurs after historic price data has been loaded,
' but before the first tick is notified
Public Sub IStrategy_Start(ByVal pTradingContext As TradingContext, _
                ByVal pStopStrategyFactory As IPosnMgmtStrtgyFactory)
Const ProcName As String = "IStrategy_Start"
On Error GoTo Err

Set mStopStrategyFactory = pStopStrategyFactory

Set mContext = pTradingContext

Set mPositionManagerDummy = mContext.PositionManagerDummy

Set mOrderContextBuyDummy = mContext.PositionManagerDummy.OrderContexts.Add("BUY")
mOrderContextBuyDummy.PreventUnprotectedPositions = True

Set mOrderContextSellDummy = mContext.PositionManagerDummy.OrderContexts.Add("SELL")
mOrderContextSellDummy.PreventUnprotectedPositions = True

Set mOrderContextBuy = mContext.PositionManager.OrderContexts.Add("BUY")
mOrderContextBuy.PreventUnprotectedPositions = True

Set mOrderContextSell = mContext.PositionManager.OrderContexts.Add("SELL")
mOrderContextSell.PreventUnprotectedPositions = True

Set mBars = mContext.GetBars(mTimeframe)

mCurrColour = getColour(0)
mPrevColour = mCurrColour

Dim i As Integer
For i = 1 To 9
    mPrevColour = getColour(-i)
    If mPrevColour <> mCurrColour Then Exit For
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mBracketOrderBuyDummy Event Handlers
'@================================================================================

Private Sub mBracketOrderBuyDummy_Completed()
Const ProcName As String = "mBracketOrderBuyDummy_Completed"
On Error GoTo Err

checkTradingOpportunity

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mBracketOrderBuyDummy_SelfCancelled()
Const ProcName As String = "mBracketOrderBuyDummy_SelfCancelled"
On Error GoTo Err

checkTradingOpportunity

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mBracketOrderSellDummy Event Handlers
'@================================================================================

Private Sub mBracketOrderSellDummy_Completed()
Const ProcName As String = "mBracketOrderSellDummy_Completed"
On Error GoTo Err

checkTradingOpportunity

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mBracketOrderSellDummy_SelfCancelled()
Const ProcName As String = "mBracketOrderSellDummy_SelfCancelled"
On Error GoTo Err

checkTradingOpportunity

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mPositionManagerDummy Event Handlers
'@================================================================================

Private Sub mPositionManagerDummy_Change(ev As ChangeEventData)
Const ProcName As String = "mPositionManagerDummy_Change"
On Error GoTo Err

Dim l As PositionManagerChangeTypes
l = ev.ChangeType
Select Case l
Case PositionSizeChanged
    If mPositionManagerDummy.PositionSize = 0 And _
        mPositionManagerDummy.PendingPositionSize = 0 Then checkTradingOpportunity
Case ProviderReadinessChanged

Case PositionClosed

End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Private Function CancelPendingBuyOrders() As Boolean
Const ProcName As String = "CancelPendingBuyOrders"
On Error GoTo Err

If mOrderContextBuyDummy.PendingSize = 0 Then Exit Function

mContext.LogTradeReason "Previous signal cancelled - cancelling pending buy orders"
mBracketOrderBuyDummy.Cancel
If mOrderContextBuy.PendingSize <> 0 Then
    mBracketOrderBuy.Cancel
End If

CancelPendingBuyOrders = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function CancelPendingSellOrders() As Boolean
Const ProcName As String = "CancelPendingSellOrders"
On Error GoTo Err

If mOrderContextSellDummy.PendingSize = 0 Then Exit Function

mContext.LogTradeReason "Previous signal cancelled - cancelling pending sell orders"
mBracketOrderSellDummy.Cancel
If mOrderContextSell.PendingSize <> 0 Then
    mBracketOrderSell.Cancel
End If

CancelPendingSellOrders = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub InitialiseTickProcessing()
Const ProcName As String = "InitialiseTickProcessing"
On Error GoTo Err

Dim colour As Colours
colour = getColour(-1)
If colour <> mCurrColour Then
    mPrevColour = mCurrColour
    mCurrColour = colour
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IsBuySignal() As Boolean
Const ProcName As String = "IsBuySignal"
On Error GoTo Err

IsBuySignal = False

Dim lMessage As String
If Not IsPotentialBuy(lMessage) Then Exit Function

If IsConfirmedBuy(lMessage) Then IsBuySignal = True

If mOrderContextBuyDummy.PendingSize = 0 Then mContext.LogTradeReason lMessage

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IsConfirmedBuy(lMessage) As Boolean
Const ProcName As String = "IsConfirmedBuy"
On Error GoTo Err

Dim lSwingLowPoint As SwingPoint
lSwingLowPoint = mSwing.SwingLowPoint(0)

Dim lPotentialSwingLowPoint As SwingPoint
lPotentialSwingLowPoint = mSwing.PotentialSwingLowPoint

If mBarNumber - lSwingLowPoint.BarNumber <= 3 Then
    IsConfirmedBuy = True
ElseIf lPotentialSwingLowPoint.Value = MaxDouble Then
    lMessage = lMessage & getRejectMessage("more than 3 Bars since swing low at ") & FormatTimestamp(lSwingLowPoint.BarStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
    IsConfirmedBuy = False
ElseIf mBarNumber - lPotentialSwingLowPoint.BarNumber > 3 Then
    lMessage = lMessage & getRejectMessage("more than 3 Bars since potential swing low at ") & FormatTimestamp(lPotentialSwingLowPoint.BarStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
    IsConfirmedBuy = False
ElseIf mBarNumber - lPotentialSwingLowPoint.BarNumber = 1 Then
    lMessage = lMessage & getRejectMessage("signal bar is potential swing low")
    IsConfirmedBuy = False
ElseIf mBarNumber = lPotentialSwingLowPoint.BarNumber Then
    lMessage = lMessage & getRejectMessage("potential new swing low this bar")
    IsConfirmedBuy = False
Else
    IsConfirmedBuy = True
End If

If mBoll.InTopBand(mBars.HighValue(-1)).Value Then
    lMessage = lMessage & getRejectMessage("signal bar high is in top Bollinger band")
    IsConfirmedBuy = False
End If
If mTimestamp >= mLastHalfHour Then
    lMessage = lMessage & getRejectMessage("within last half-hour of session")
    IsConfirmedBuy = False
End If

If IsConfirmedBuy Then lMessage = lMessage & vbCrLf & Space(Len(MessagePrefix)) & "Confirmed buy"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IsConfirmedSell(lMessage) As Boolean
Const ProcName As String = "IsConfirmedSell"
On Error GoTo Err

Dim lSwingHighPoint As SwingPoint
lSwingHighPoint = mSwing.SwingHighPoint(0)

Dim lPotentialSwingHighPoint As SwingPoint
lPotentialSwingHighPoint = mSwing.PotentialSwingHighPoint

If mBarNumber - lSwingHighPoint.BarNumber <= 3 Then
    IsConfirmedSell = True
ElseIf lPotentialSwingHighPoint.Value = MinDouble Then
    lMessage = lMessage & getRejectMessage("more than 3 Bars since swing high at ") & FormatTimestamp(lSwingHighPoint.BarStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
    IsConfirmedSell = False
ElseIf mBarNumber - lPotentialSwingHighPoint.BarNumber > 3 Then
    lMessage = lMessage & getRejectMessage("more than 3 Bars since potential swing high at ") & FormatTimestamp(lPotentialSwingHighPoint.BarStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
    IsConfirmedSell = False
ElseIf mBarNumber - lPotentialSwingHighPoint.BarNumber = 1 Then
    lMessage = lMessage & getRejectMessage("signal bar is potential swing high")
    IsConfirmedSell = False
ElseIf mBarNumber = lPotentialSwingHighPoint.BarNumber Then
    lMessage = lMessage & getRejectMessage("potential new swing high this bar")
    IsConfirmedSell = False
Else
    IsConfirmedSell = True
End If

If mBoll.InBottomBand(mBars.LowValue(-1)).Value Then
    lMessage = lMessage & getRejectMessage("signal bar low is in bottom Bollinger band")
    IsConfirmedSell = False
End If
If mTimestamp >= mLastHalfHour Then
    lMessage = lMessage & getRejectMessage("within last half-hour of session")
    IsConfirmedSell = False
End If

If IsConfirmedSell Then lMessage = lMessage & vbCrLf & Space(Len(MessagePrefix)) & "Confirmed sell"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IsInEstablishedPosition() As Boolean
Const ProcName As String = "IsInEstablishedPosition"
On Error GoTo Err

IsInEstablishedPosition = mPositionManagerDummy.PositionSize <> 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IsPendingBuyOrders() As Boolean
Const ProcName As String = "IsPendingBuyOrders"
On Error GoTo Err

IsPendingBuyOrders = mOrderContextBuyDummy.PendingSize <> 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IsPendingSellOrders() As Boolean
Const ProcName As String = "IsPendingSellOrders"
On Error GoTo Err

IsPendingSellOrders = mOrderContextSellDummy.PendingSize <> 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IsPotentialBuy(lMessage) As Boolean
If mCurrColour = Green Then
    lMessage = MessagePrefix & "Potential buy: signs of strength: current colour=Green"
    IsPotentialBuy = True
ElseIf (mCurrColour = Amber And mPrevColour = Red) Then
    lMessage = MessagePrefix & "Potential buy: signs of strength: current colour=Amber; prev colour=Red"
    IsPotentialBuy = True
Else
    IsPotentialBuy = False
End If
End Function

Private Function IsPotentialSell(lMessage) As Boolean
If mCurrColour = Red Then
    lMessage = MessagePrefix & "Potential sell: signs of weakness: current colour=Red"
    IsPotentialSell = True
ElseIf (mCurrColour = Pink And mPrevColour = Green) Then
    lMessage = MessagePrefix & "Potential sell: signs of weakness: current colour=Pink; prev colour=Green"
    IsPotentialSell = True
Else
    IsPotentialSell = False
End If
End Function

Private Function IsSellSignal() As Boolean
Const ProcName As String = "IsSellSignal"
On Error GoTo Err

IsSellSignal = False

Dim lMessage As String

If Not IsPotentialSell(lMessage) Then Exit Function

If IsConfirmedSell(lMessage) Then IsSellSignal = True

If mOrderContextSellDummy.PendingSize = 0 Then mContext.LogTradeReason lMessage

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub ModifyPendingBuyOrders()
Const ProcName As String = "ModifyPendingBuyOrders"
On Error GoTo Err

Dim lTriggerPrice As Double
lTriggerPrice = getTriggerPriceLong
If lTriggerPrice >= mBracketOrderBuyDummy.EntryOrder.triggerPrice Then Exit Sub

Dim stopPrice As Double
Dim breakevenPrice As Double
Dim cancelPrice As Double
Dim risk As Double

calcPricesForBuy lTriggerPrice, stopPrice, breakevenPrice, cancelPrice, risk
mBracketOrderBuyDummy.SetNewEntryTriggerPrice lTriggerPrice
mBracketOrderBuyDummy.SetNewStopLossTriggerPrice stopPrice

mBracketOrderBuyDummy.Update
    
If mOrderContextBuy.PendingSize = 0 Then Exit Sub

mBracketOrderBuy.SetNewQuantity UnitSize(risk)
mBracketOrderBuy.SetNewEntryTriggerPrice lTriggerPrice
mBracketOrderBuy.SetNewEntryPrice lTriggerPrice - mEntryLimitOffsetTicks * mTickSize
mBracketOrderBuy.SetNewStopLossTriggerPrice stopPrice
mBracketOrderBuy.Update

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ModifyPendingSellOrders()
Const ProcName As String = "ModifyPendingSellOrders"
On Error GoTo Err

Dim lTriggerPrice As Double
lTriggerPrice = getTriggerPriceShort
If lTriggerPrice <= mBracketOrderSellDummy.EntryOrder.triggerPrice Then Exit Sub

Dim stopPrice As Double
Dim breakevenPrice As Double
Dim cancelPrice As Double
Dim risk As Double

calcPricesForSell lTriggerPrice, stopPrice, breakevenPrice, cancelPrice, risk
mBracketOrderSellDummy.SetNewEntryTriggerPrice lTriggerPrice
mBracketOrderSellDummy.SetNewStopLossTriggerPrice stopPrice

mBracketOrderSellDummy.Update
    
If mOrderContextSell.PendingSize = 0 Then Exit Sub

mBracketOrderSell.SetNewQuantity UnitSize(risk)
mBracketOrderSell.SetNewEntryTriggerPrice lTriggerPrice
mBracketOrderSell.SetNewEntryPrice lTriggerPrice + mEntryLimitOffsetTicks * mTickSize
mBracketOrderSell.SetNewStopLossTriggerPrice stopPrice
mBracketOrderSell.Update

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

        
Private Sub PlaceBuyOrders()
Const ProcName As String = "PlaceBuyOrders"
On Error GoTo Err

Dim lTriggerPrice As Double
lTriggerPrice = getTriggerPriceLong

Dim stopPrice As Double
Dim breakevenPrice As Double
Dim cancelPrice As Double
Dim risk As Double

calcPricesForBuy lTriggerPrice, stopPrice, breakevenPrice, cancelPrice, risk
If stopPrice = 0 Then
    mContext.LogTradeReason "    reject: trigger price too close to breakeven price"
    Exit Sub
End If

Set mBracketOrderBuyDummy = mOrderContextBuyDummy.Buy( _
                    Number:=1, _
                    BracketEntryType:=BracketEntryTypeStop, _
                    buyPrice:=0, _
                    buyOffset:=0, _
                    triggerPrice:=lTriggerPrice, _
                    BracketStopLossType:=BracketStopLossTypeStop, _
                    stopPrice:=stopPrice, _
                    stopOffset:=0, _
                    stopLimitPrice:=0, _
                    BracketTargetType:=BracketTargetTypeNone, _
                    targetPrice:=0, _
                    targetOffset:=0, _
                    targetTriggerPrice:=0, _
                    pCancelPrice:=cancelPrice)

mContext.ApplyPositionManagementStrategy mBracketOrderBuyDummy, mStopStrategyFactory.CreateStrategy(mContext)

If mBracketOrderBuyDummy Is Nothing Then
    mNoTradeThisBar = True
        ' prevents each subsequent tick being processed if money management has prevented
        ' further trades
    Exit Sub
End If

If mOrderContextBuy.ActiveSize <> 0 Or mOrderContextBuy.PendingSize <> 0 Then Exit Sub

If UnitSize(risk) = 0 Then
    mContext.LogTradeReason "    reject: allowable initial risk exceeded"
    Exit Sub
End If

Set mBracketOrderBuy = mOrderContextBuy.Buy( _
                    Number:=UnitSize(risk), _
                    BracketEntryType:=BracketEntryTypeStopLimit, _
                    buyPrice:=lTriggerPrice, _
                    buyOffset:=-mEntryLimitOffsetTicks, _
                    triggerPrice:=lTriggerPrice, _
                    BracketStopLossType:=BracketStopLossTypeStop, _
                    stopPrice:=stopPrice, _
                    stopOffset:=0, _
                    stopLimitPrice:=0, _
                    BracketTargetType:=BracketTargetTypeNone, _
                    targetPrice:=0, _
                    targetOffset:=0, _
                    targetTriggerPrice:=0, _
                    pCancelPrice:=cancelPrice, _
                    pCancelAfter:=300)
mContext.ApplyPositionManagementStrategy mBracketOrderBuy, mStopStrategyFactory.CreateStrategy(mContext)
'mContext.ApplyPositionManagementStrategy mBracketOrderBuy, mScalingStrategyFactory.CreateStrategy(mContext)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub PlaceSellOrders()
Const ProcName As String = "PlaceSellOrders"
On Error GoTo Err

Dim lTriggerPrice As Double
lTriggerPrice = getTriggerPriceShort

Dim stopPrice As Double
Dim breakevenPrice As Double
Dim cancelPrice As Double
Dim risk As Double

calcPricesForSell lTriggerPrice, stopPrice, breakevenPrice, cancelPrice, risk
If stopPrice = 0 Then
    mContext.LogTradeReason "    reject: trigger price too close to breakeven price"
    Exit Sub
End If

Set mBracketOrderSellDummy = mOrderContextSellDummy.Sell( _
                    Number:=1, _
                    BracketEntryType:=BracketEntryTypes.BracketEntryTypeStop, _
                    sellPrice:=0, _
                    sellOffset:=0, _
                    triggerPrice:=lTriggerPrice, _
                    BracketStopLossType:=BracketStopLossTypeStop, _
                    stopPrice:=stopPrice, _
                    stopOffset:=0, _
                    stopLimitPrice:=0, _
                    BracketTargetType:=BracketTargetTypeNone, _
                    targetPrice:=0, _
                    targetOffset:=0, _
                    targetTriggerPrice:=0, _
                    pCancelPrice:=cancelPrice)

mContext.ApplyPositionManagementStrategy mBracketOrderSellDummy, mStopStrategyFactory.CreateStrategy(mContext)

If mBracketOrderSellDummy Is Nothing Then
    mNoTradeThisBar = True
        ' prevents each subsequent tick being processed if money management has prevented
        ' further trades
    Exit Sub
End If

If mOrderContextSell.ActiveSize <> 0 Or mOrderContextSell.PendingSize <> 0 Then Exit Sub

If UnitSize(risk) = 0 Then
    mContext.LogTradeReason "    reject: allowable initial risk exceeded"
    Exit Sub
End If

Set mBracketOrderSell = mOrderContextSell.Sell( _
                    Number:=UnitSize(risk), _
                    BracketEntryType:=BracketEntryTypeStopLimit, _
                    sellPrice:=lTriggerPrice, _
                    sellOffset:=mEntryLimitOffsetTicks, _
                    triggerPrice:=lTriggerPrice, _
                    BracketStopLossType:=BracketStopLossTypeStop, _
                    stopPrice:=stopPrice, _
                    stopOffset:=0, _
                    stopLimitPrice:=0, _
                    BracketTargetType:=BracketTargetTypeNone, _
                    targetPrice:=0, _
                    targetOffset:=0, _
                    targetTriggerPrice:=0, _
                    pCancelPrice:=cancelPrice, _
                    pCancelAfter:=300)
mContext.ApplyPositionManagementStrategy mBracketOrderSell, mStopStrategyFactory.CreateStrategy(mContext)
'mContext.ApplyPositionManagementStrategy mBracketOrderSell, mScalingStrategyFactory.CreateStrategy(mContext)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcPricesForBuy( _
                ByVal triggerPrice As Double, _
                ByRef stopPrice As Double, _
                ByRef breakevenPrice As Double, _
                ByRef cancelPrice As Double, _
                ByRef risk As Double)
Const ProcName As String = "calcPricesForBuy"
On Error GoTo Err

Dim SwingLow As Double
SwingLow = IIf(mSwing.SwingLowPoint(0).Value < mSwing.PotentialSwingLowPoint.Value, mSwing.SwingLowPoint(0).Value, mSwing.PotentialSwingLowPoint.Value)
If mSwingToBreakeven <> 0 Then
    If (triggerPrice - SwingLow) > 0.5 * mSwingToBreakeven Then Exit Sub
    breakevenPrice = SwingLow + mSwingToBreakeven
End If

stopPrice = SwingLow - mStopBreakoutThresholdTicks * mTickSize

If stopPrice < triggerPrice - mMaxInitialStop Then
    stopPrice = triggerPrice - mMaxInitialStop
End If

risk = (triggerPrice + mEntryLimitOffsetTicks * mTickSize - stopPrice) * mTickValue / mTickSize

If mSwing.PotentialSwingLowPoint.Value = MaxDouble Then
    cancelPrice = mSwing.SwingLowPoint(0).Value
Else
    cancelPrice = mSwing.PotentialSwingLowPoint.Value
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub calcPricesForSell( _
                ByVal triggerPrice As Double, _
                ByRef stopPrice As Double, _
                ByRef breakevenPrice As Double, _
                ByRef cancelPrice As Double, _
                ByRef risk As Double)
Const ProcName As String = "calcPricesForSell"
On Error GoTo Err

Dim SwingHigh As Double
SwingHigh = IIf(mSwing.SwingHighPoint(0).Value > mSwing.PotentialSwingHighPoint.Value, mSwing.SwingHighPoint(0).Value, mSwing.PotentialSwingHighPoint.Value)
If mSwingToBreakeven <> 0 Then
    If (SwingHigh - triggerPrice) > 0.5 * mSwingToBreakeven Then Exit Sub
    breakevenPrice = SwingHigh - mSwingToBreakeven
End If

stopPrice = SwingHigh + mStopBreakoutThresholdTicks * mTickSize

If stopPrice > triggerPrice + mMaxInitialStop Then
    stopPrice = triggerPrice + mMaxInitialStop
End If

risk = (stopPrice - triggerPrice - mEntryLimitOffsetTicks * mTickSize) * mTickValue / mTickSize

If mSwing.PotentialSwingHighPoint.Value = MinDouble Then
    cancelPrice = mSwing.SwingHighPoint(0).Value
Else
    cancelPrice = mSwing.PotentialSwingHighPoint.Value
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub checkTradingOpportunity()
Const ProcName As String = "checkTradingOpportunity"
On Error GoTo Err

If IsInEstablishedPosition Then Exit Sub

InitialiseTickProcessing

If IsBuySignal Then
    If CancelPendingSellOrders Then
        ' do nothing till pending orders cancelled
    Else
        If Not IsPendingBuyOrders Then
            PlaceBuyOrders
        Else
            ModifyPendingBuyOrders
        End If
    End If
ElseIf IsSellSignal Then
    If CancelPendingBuyOrders Then
        ' do nothing till pending orders cancelled
    Else
        If Not IsPendingSellOrders Then
            PlaceSellOrders
        Else
            ModifyPendingSellOrders
        End If
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getColour(ByVal ref As Integer) As Colours
Const ProcName As String = "getColour"
On Error GoTo Err

Dim lMACDUp As Boolean
lMACDUp = mMACD.MACDUp(ref).Value

Dim lMACDDown As Boolean
lMACDDown = mMACD.MACDDown(ref).Value

Dim lMACDHistUp As Boolean
lMACDHistUp = mMACD.MACDHistUp(ref).Value

Dim lMACDHistDown As Boolean
lMACDHistDown = mMACD.MACDHistDown(ref).Value

If lMACDUp And lMACDHistUp Then
    getColour = Green
ElseIf lMACDUp And lMACDHistDown Then
    getColour = Pink
ElseIf lMACDDown And lMACDHistUp Then
    getColour = Amber
Else
    getColour = Red
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getTriggerPriceLong() As Double
getTriggerPriceLong = mBars.CloseValue(-1) + mEntryBreakoutThresholdTicks * mTickSize
End Function

Private Function getTriggerPriceShort() As Double
getTriggerPriceShort = mBars.CloseValue(-1) - mEntryBreakoutThresholdTicks * mTickSize
End Function

Private Sub processParameters(ByVal pParams As Parameters)
Const ProcName As String = "processParameters"
On Error GoTo Err

Set mParams = pParams
mEntryBreakoutThresholdTicks = mParams.GetParameterValue(ParamEntryBreakoutThresholdTicks)
mStopBreakoutThresholdTicks = mParams.GetParameterValue(ParamBreakoutThresholdTicks)
mEntryLimitOffsetTicks = mParams.GetParameterValue(ParamEntryLimitOffsetTicks)
mRiskUnitPercent = mParams.GetParameterValue(ParamRiskUnitPercent)
mMaxTradeSize = mParams.GetParameterValue(ParamMaxTradeSize)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getRejectMessage(pReason As String) As String
getRejectMessage = vbCrLf & Space(Len(MessagePrefix)) & "reject: " & pReason
End Function


Private Function UnitSize(ByVal risk As Double) As Long
Const ProcName As String = "UnitSize"
On Error GoTo Err

UnitSize = Int((mRiskUnitPercent * mOrderContextBuy.GetAccountBalance / 100) / risk)
If UnitSize > mMaxTradeSize Then UnitSize = mMaxTradeSize
'If UnitSize = 0 Then UnitSize = 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function





