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

Private mTriggerPrice As Double

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
                            params, _
                            3, _
                            , _
                            True)

Set params = New Parameters
params.SetParameterValue "Periods", mParams.GetParameterValue(ParamBollingerPeriods)
params.SetParameterValue "Standard deviations", 2
params.SetParameterValue "Mov avg type", mParams.GetParameterValue(ParamBollingerMovingAverageType)
params.SetParameterValue "Centre band width", mParams.GetParameterValue(ParamBollingerCentreBandWidthTicks)
params.SetParameterValue "Edge band width", mParams.GetParameterValue(ParamBollingerEdgeBandWidthTicks)
linputvaluenames(0) = "$default"
Set mBoll = pContext.AddStudy("Bollinger Bands", _
                        mTimeframe.BarStudy, _
                        linputvaluenames, _
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

If Not mContext.ReceivedFirstBidAndAsk Then Exit Sub
If mBars.TickVolume <> 1 Then Exit Sub
If pType <> TickTypes.TickTypeTrade Then Exit Sub

mTimestamp = mContext.TimeStamp

If mTimestamp >= mLastMinute Then
    mContext.PositionManager.ClosePositions
    Exit Sub
End If


mBarNumber = mBars.Bar.BarNumber
mNoTradeThisBar = False

Dim colour As Colours
colour = getColour(-1)
If colour <> mCurrColour Then
    mPrevColour = mCurrColour
    mCurrColour = colour
End If

If mCurrColour = Red Or _
    (mCurrColour = Pink And mPrevColour = Green) _
Then
    If mOrderContextBuyDummy.PendingSize <> 0 Then
        mBracketOrderBuyDummy.Cancel
        If mOrderContextBuy.PendingSize <> 0 Then
            mBracketOrderBuy.Cancel
        End If
    End If
ElseIf mCurrColour = Green Or _
        (mCurrColour = Amber And mPrevColour = Red) _
Then
    If mOrderContextSellDummy.PendingSize <> 0 Then
        mBracketOrderSellDummy.Cancel
        If mOrderContextSell.PendingSize <> 0 Then
            mBracketOrderSell.Cancel
        End If
    End If
End If

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
Set mOrderContextSellDummy = mContext.PositionManagerDummy.OrderContexts.Add("SELL")

Set mOrderContextBuy = mContext.PositionManager.OrderContexts.Add("BUY")
Set mOrderContextSell = mContext.PositionManager.OrderContexts.Add("SELL")

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
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkTradingOpportunity()
Const ProcName As String = "checkTradingOpportunity"
On Error GoTo Err

If mNoTradeThisBar Then Exit Sub
If mPositionManagerDummy.PositionSize <> 0 Then Exit Sub

If mCurrColour = Red Then
    sell getTriggerPriceShort, "signs of weakness: current colour=Red"
ElseIf (mCurrColour = Pink And mPrevColour = Green) Then
    sell getTriggerPriceShort, "signs of weakness: current colour=Pink; prev colour=Green"
ElseIf mCurrColour = Green Then
    buy getTriggerPriceLong, "signs of strength: current colour=Green"
ElseIf (mCurrColour = Amber And mPrevColour = Red) Then
    buy getTriggerPriceLong, "signs of strength: current colour=Amber; prev colour=Red"
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

Private Function okToBuy(ByVal triggerPrice As Double) As Boolean
Const ProcName As String = "okToBuy"
On Error GoTo Err

okToBuy = True

Dim lSwingLowPoint As SwingPoint
lSwingLowPoint = mSwing.SwingLowPoint(0)

Dim lPotentialSwingLowPoint As SwingPoint
lPotentialSwingLowPoint = mSwing.PotentialSwingLowPoint

If mBarNumber - lSwingLowPoint.BarNumber > 3 Then
    If lPotentialSwingLowPoint.Value = MaxDouble Then
        mContext.LogTradeReason "    reject: more than 3 Bars since swing low at " & FormatTimestamp(lSwingLowPoint.BarStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
        okToBuy = False
    ElseIf mBarNumber - lPotentialSwingLowPoint.BarNumber > 3 Then
        mContext.LogTradeReason "    reject: more than 3 Bars since potential swing low at " & FormatTimestamp(lPotentialSwingLowPoint.BarStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
        okToBuy = False
    ElseIf mBarNumber - lPotentialSwingLowPoint.BarNumber = 1 Then
        mContext.LogTradeReason "    reject: signal bar is potential swing low"
        okToBuy = False
    ElseIf mBarNumber = lPotentialSwingLowPoint.BarNumber Then
        mContext.LogTradeReason "    reject: potential new swing low this bar"
        okToBuy = False
    End If
End If

If mBoll.InTopBand(mBars.HighValue(-1)).Value Then
    mContext.LogTradeReason "    reject: signal bar high is in top Bollinger band"
    okToBuy = False
End If

'If mContext.lastPrice > mBoll.Top Then
'    mcontext.LogTradeReason "    reject: price is above top Bollinger band"
'    okToBuy = False
'End If

If mTimestamp >= mLastHalfHour Then
    mContext.LogTradeReason "    reject: within last half-hour of session"
    okToBuy = False
End If

If Not okToBuy Then mNoTradeThisBar = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function okToSell(ByVal triggerPrice As Double) As Boolean
Const ProcName As String = "okToSell"
On Error GoTo Err

okToSell = True

Dim lSwingHighPoint As SwingPoint
lSwingHighPoint = mSwing.SwingHighPoint(0)

Dim lPotentialSwingHighPoint As SwingPoint
lPotentialSwingHighPoint = mSwing.PotentialSwingHighPoint

If mBarNumber - lSwingHighPoint.BarNumber > 3 Then
    If lPotentialSwingHighPoint.Value = MinDouble Then
        mContext.LogTradeReason "    reject: more than 3 Bars since swing high at " & FormatTimestamp(lSwingHighPoint.BarStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
        okToSell = False
    ElseIf mBarNumber - lSwingHighPoint.BarNumber > 3 Then
        mContext.LogTradeReason "    reject: more than 3 Bars since potential swing high at " & FormatTimestamp(lPotentialSwingHighPoint.BarStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
        okToSell = False
    ElseIf mBarNumber - lSwingHighPoint.BarNumber = 1 Then
        mContext.LogTradeReason "    reject: signal bar is potential swing high"
        okToSell = False
    ElseIf mBarNumber = lSwingHighPoint.BarNumber Then
        mContext.LogTradeReason "    reject: potential new swing high this bar"
        okToSell = False
    End If
End If

If mBoll.InBottomBand(mBars.LowValue(-1)).Value Then
    mContext.LogTradeReason "    reject: signal bar low is in bottom Bollinger band"
    okToSell = False
End If

'If mContext.lastPrice < mBoll.bottom Then
'    mcontext.LogTradeReason "    reject: price is below bottom Bollinger band"
'    okToSell = False
'End If

If mTimestamp >= mLastHalfHour Then
    mContext.LogTradeReason "    reject: within last half-hour of session"
    okToSell = False
End If

If Not okToSell Then mNoTradeThisBar = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function stillOkToBuy(ByVal triggerPrice As Double) As Boolean
Const ProcName As String = "stillOkToBuy"
On Error GoTo Err

stillOkToBuy = True

If mBarNumber - mSwing.SwingLowPoint(0).BarNumber > 3 Then
    If mSwing.PotentialSwingLowPoint.Value <> MaxDouble And _
        mBarNumber = mSwing.PotentialSwingLowPoint.BarNumber _
    Then
        mContext.LogTradeReason "    reject: potential new swing low this bar"
        stillOkToBuy = False
    End If
End If

'If mBoll.intopBand(mBars.highPrice(-1)) Then
'    mcontext.LogTradeReason "    reject: signal bar high is in top Bollinger band"
'    stillOkToBuy = False
'End If

If mTimestamp >= mLastHalfHour Then
    mContext.LogTradeReason "    reject: within last half-hour of session"
    stillOkToBuy = False
End If

If Not stillOkToBuy Then mNoTradeThisBar = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function stillOkToSell(ByVal triggerPrice As Double) As Boolean
Const ProcName As String = "stillOkToSell"
On Error GoTo Err

stillOkToSell = True

If mBarNumber - mSwing.SwingHighPoint(0).BarNumber > 3 Then
    If mSwing.PotentialSwingHighPoint.Value <> MinDouble And _
        mBarNumber = mSwing.PotentialSwingHighPoint.BarNumber _
    Then
        mContext.LogTradeReason "    reject: potential new swing high this bar"
        stillOkToSell = False
    End If
End If

'If mBoll.inBottomBand(mBars.lowPrice(-1)) Then
'    mcontext.LogTradeReason "    reject: signal bar low is in bottom Bollinger band"
'    stillOkToSell = False
'End If

If mTimestamp >= mLastHalfHour Then
    mContext.LogTradeReason "    reject: within last half-hour of session"
    stillOkToSell = False
End If

If Not stillOkToSell Then mNoTradeThisBar = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub sell( _
    ByVal triggerPrice As Double, _
    ByVal reason As String)
Const ProcName As String = "sell"
On Error GoTo Err

Dim stopPrice As Double
Dim breakevenPrice As Double
Dim cancelPrice As Double
Dim risk As Double

If mOrderContextSellDummy.PendingSize <> 0 Then
    If Not stillOkToSell(triggerPrice) Then
        mBracketOrderSellDummy.Cancel
        If mOrderContextSell.PendingSize <> 0 Then mBracketOrderSell.Cancel
        Exit Sub
    End If
    
    If triggerPrice <= mBracketOrderSellDummy.EntryOrder.triggerPrice Then Exit Sub
        
    calcPricesForSell triggerPrice, stopPrice, breakevenPrice, cancelPrice, risk
    mBracketOrderSellDummy.SetNewEntryTriggerPrice triggerPrice
    mBracketOrderSellDummy.SetNewStopLossTriggerPrice stopPrice
    On Error Resume Next
    mBracketOrderSellDummy.Update
    If Err.Number = 0 Then
        On Error GoTo 0
        If mOrderContextSell.PendingSize = 0 Then Exit Sub
        mBracketOrderSell.SetNewQuantity UnitSize(risk)
        mBracketOrderSell.SetNewEntryTriggerPrice triggerPrice
        mBracketOrderSell.SetNewEntryPrice triggerPrice + mEntryLimitOffsetTicks * mTickSize
        mBracketOrderSell.SetNewStopLossTriggerPrice stopPrice
        On Error Resume Next
        mBracketOrderSell.Update
        If Err.Number = 0 Then Exit Sub
    End If
    ' the existing order plex cannot be amended - so cancel it
    ' and create a new one
    On Error GoTo 0
    mBracketOrderSellDummy.Cancel
    If mOrderContextSell.PendingSize <> 0 Then mBracketOrderSell.Cancel
End If

mContext.LogTradeReason ">>>>>>>>>>>>>>>> Potential sell: " & reason

If Not okToSell(triggerPrice) Then Exit Sub

calcPricesForSell triggerPrice, stopPrice, breakevenPrice, cancelPrice, risk
If stopPrice = 0 Then
    mContext.LogTradeReason "    reject: trigger price too close to breakeven price"
    Exit Sub
End If

Dim stopStrategyDummy As StopStrategy5
Set stopStrategyDummy = mStopStrategyFactory.CreateStrategy(mContext)

Set mBracketOrderSellDummy = mOrderContextSellDummy.sell( _
                    Number:=1, _
                    BracketEntryType:=BracketEntryTypes.BracketEntryTypeStop, _
                    sellPrice:=0, _
                    sellOffset:=0, _
                    triggerPrice:=triggerPrice, _
                    BracketStopLossType:=BracketStopLossTypeStop, _
                    stopPrice:=stopPrice, _
                    stopOffset:=0, _
                    stopLimitPrice:=0, _
                    BracketTargetType:=BracketTargetTypeNone, _
                    targetPrice:=0, _
                    targetOffset:=0, _
                    targetTriggerPrice:=0, _
                    pCancelPrice:=cancelPrice)
mContext.ApplyPositionManagementStrategy mBracketOrderSellDummy, stopStrategyDummy

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

Set mBracketOrderSell = mOrderContextSell.sell( _
                    Number:=UnitSize(risk), _
                    BracketEntryType:=BracketEntryTypeStopLimit, _
                    sellPrice:=triggerPrice, _
                    sellOffset:=mEntryLimitOffsetTicks, _
                    triggerPrice:=triggerPrice, _
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

Private Sub buy( _
    ByVal triggerPrice As Double, _
    ByVal reason As String)
Const ProcName As String = "buy"
On Error GoTo Err

Dim stopPrice As Double
Dim breakevenPrice As Double
Dim cancelPrice As Double
Dim risk As Double
Dim scaleStrategy As ScaleStrategy1

If mOrderContextBuyDummy.PendingSize <> 0 Then
    If Not stillOkToBuy(triggerPrice) Then
        mBracketOrderBuyDummy.Cancel
        If mOrderContextBuy.PendingSize <> 0 Then mBracketOrderBuy.Cancel
        Exit Sub
    End If
    
    If triggerPrice >= mBracketOrderBuyDummy.EntryOrder.triggerPrice Then Exit Sub
    
    calcPricesForBuy triggerPrice, stopPrice, breakevenPrice, cancelPrice, risk
    mBracketOrderBuyDummy.SetNewEntryTriggerPrice triggerPrice
    mBracketOrderBuyDummy.SetNewStopLossTriggerPrice stopPrice
    On Error Resume Next
    mBracketOrderBuyDummy.Update
    If Err.Number = 0 Then
        On Error GoTo 0
        If mOrderContextBuy.PendingSize = 0 Then Exit Sub
        mBracketOrderBuy.SetNewQuantity UnitSize(risk)
        mBracketOrderBuy.SetNewEntryTriggerPrice triggerPrice
        mBracketOrderBuy.SetNewEntryPrice triggerPrice - mEntryLimitOffsetTicks * mTickSize
        mBracketOrderBuy.SetNewStopLossTriggerPrice stopPrice
        On Error Resume Next
        mBracketOrderBuy.Update
        If Err.Number = 0 Then Exit Sub
    End If
    ' the existing order plex cannot be amended - so cancel it
    ' and create a new one
    On Error GoTo 0
    mBracketOrderBuyDummy.Cancel
    If mOrderContextBuy.PendingSize <> 0 Then mBracketOrderBuy.Cancel
End If

mContext.LogTradeReason ">>>>>>>>>>>>>>>> Potential buy: " & reason

If Not okToBuy(triggerPrice) Then Exit Sub

calcPricesForBuy triggerPrice, stopPrice, breakevenPrice, cancelPrice, risk
If stopPrice = 0 Then
    mContext.LogTradeReason "    reject: trigger price too close to breakeven price"
    Exit Sub
End If

Dim stopStrategyDummy As StopStrategy5
Set stopStrategyDummy = mStopStrategyFactory.CreateStrategy(mContext)

Set mBracketOrderBuyDummy = mOrderContextBuyDummy.buy( _
                    Number:=1, _
                    BracketEntryType:=BracketEntryTypeStop, _
                    buyPrice:=0, _
                    buyOffset:=0, _
                    triggerPrice:=triggerPrice, _
                    BracketStopLossType:=BracketStopLossTypeStop, _
                    stopPrice:=stopPrice, _
                    stopOffset:=0, _
                    stopLimitPrice:=0, _
                    BracketTargetType:=BracketTargetTypeNone, _
                    targetPrice:=0, _
                    targetOffset:=0, _
                    targetTriggerPrice:=0, _
                    pCancelPrice:=cancelPrice)
mContext.ApplyPositionManagementStrategy mBracketOrderBuyDummy, stopStrategyDummy

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

Set mBracketOrderBuy = mOrderContextBuy.buy( _
                    Number:=UnitSize(risk), _
                    BracketEntryType:=BracketEntryTypeStopLimit, _
                    buyPrice:=triggerPrice, _
                    buyOffset:=-mEntryLimitOffsetTicks, _
                    triggerPrice:=triggerPrice, _
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

Private Sub mBracketOrderBuyDummy_Completed()
Const ProcName As String = "mBracketOrderBuyDummy_Completed"
On Error GoTo Err

checkTradingOpportunity

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mBracketOrderBuyDummy_SelfCancelled()
Const ProcName As String = "mBracketOrderBuyDummy_SelfCancelled"
On Error GoTo Err

checkTradingOpportunity

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mBracketOrderSellDummy_Completed()
Const ProcName As String = "mBracketOrderSellDummy_Completed"
On Error GoTo Err

checkTradingOpportunity

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mBracketOrderSellDummy_SelfCancelled()
Const ProcName As String = "mBracketOrderSellDummy_SelfCancelled"
On Error GoTo Err

checkTradingOpportunity

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
gHandleUnexpectedError ProcName, ModuleName
End Sub

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





