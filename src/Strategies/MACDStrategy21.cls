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

''
' This is a strategy which was actually profitable from 2003 to 2005, but is now
' only used for testing purposes as it is quite complex and is a good exercise for
' the StrategyHost.
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IBracketOrderEventSink
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

Private Const MessagePrefix                         As String = ">>> "

'@================================================================================
' Member variables
'@================================================================================

Private mBracketOrderBuyDummy As ResourceIdentifier
Private mBracketOrderSellDummy As ResourceIdentifier

Private mMaxInitialStop As Double
Private mStopBreakoutThresholdTicks As Integer
Private mEntryBreakoutThresholdTicks As Integer
Private mEntryLimitOffsetTicks As Integer

Private mPrevColour As Colours
Private mCurrColour As Colours

Private mRiskUnitPercent As Double
Private mMaxTradeSize As Long

Private mBackstopMAOffset As Integer

Private mSwingToBreakeven As Double

Private mCurrentBarNumber As Long

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
' IBracketOrderEventSink Interface Members
'@================================================================================

Private Sub IBracketOrderEventSink_NotifyBracketOrderCompletion( _
                ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "IBracketOrderEventSink_NotifyBracketOrderCompletion"
On Error GoTo Err

ClearPrimaryBracketOrder
checkTradingOpportunity

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrderEventSink_NotifyBracketOrderFill(ByVal BracketOrder As ResourceIdentifier)
' nothing to do here
End Sub

Private Sub IBracketOrderEventSink_NotifyBracketOrderStopLossAdjusted(ByVal BracketOrder As ResourceIdentifier)
' nothing to do here
End Sub

'@================================================================================
' IStrategy Interface Members
'@================================================================================

Private Sub IStrategy_DefineDefaultParameters()
Const ProcName As String = "IStrategy_DefineDefaultParameters"
On Error GoTo Err

DefineParameter ParamBarLength, BarLength
DefineParameter ParamBarUnit, BarUnit
DefineParameter ParamIncludeBarsOutsideSession, IncludeBarsOutsideSession
DefineParameter ParamBollingerPeriods, BollingerPeriods
DefineParameter ParamBollingerMovingAverageType, BollingerMovingAverageType
DefineParameter ParamMinimumSwingTicks, MinimumSwingTicks
DefineParameter ParamEntryLimitOffsetTicks, EntryLimitOffsetTicks
DefineParameter ParamMaxInitialStopTicks, MaxInitialStopTicks
DefineParameter ParamEntryBreakoutThresholdTicks, EntryBreakoutThresholdTicks
DefineParameter ParamSwingToBreakevenTicks, SwingToBreakevenTicks
DefineParameter ParamRiskUnitPercent, RiskUnitPercent
DefineParameter ParamMaxTradeSize, MaxTradeSize
DefineParameter ParamBollingerCentreBandWidthTicks, BollCentreBandwidthTicks
DefineParameter ParamBollingerEdgeBandWidthTicks, BollEdgeBandwidthTicks

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub IStrategy_Initialise( _
                ByVal pContext As InitialisationContext)
Const ProcName As String = "IStrategy_Initialise"
On Error GoTo Err

processParameters

mMaxInitialStop = GetParameterValue(ParamMaxInitialStopTicks) * TickSize
mSwingToBreakeven = GetParameterValue(ParamSwingToBreakevenTicks) * TickSize

Dim lTimeframe As ResourceIdentifier
Set lTimeframe = DeclareTimeframe( _
                        GetParameterValue(ParamBarLength), _
                        GetParameterValue(ParamBarUnit), _
                        500, _
                        GetParameterValue(ParamIncludeBarsOutsideSession), _
                        True)

Dim lSwing As ResourceIdentifier
Set lSwing = DeclareStudy( _
                        "Swing", _
                        lTimeframe, _
                        "Minimum swing (ticks)=" & GetParameterValue(ParamMinimumSwingTicks) & _
                            ";Include implicit swing points=N", _
                        GetParameterValue(ParamIncludeBarsOutsideSession), _
                        3)
ChartStudyValue lSwing, SwingValueSwingLine

Dim lBoll As ResourceIdentifier
Set lBoll = DeclareStudy( _
                        "Bollinger Bands", _
                        lTimeframe, _
                        "Periods=" & GetParameterValue(ParamBollingerPeriods) & _
                            ";Standard deviations=2" & _
                            ";Mov avg type=" & GetParameterValue(ParamBollingerMovingAverageType) & _
                            ";Centre band width=" & GetParameterValue(ParamBollingerCentreBandWidthTicks) * TickSize & _
                            ";Edge band width=" & GetParameterValue(ParamBollingerEdgeBandWidthTicks) * TickSize, _
                        GetParameterValue(ParamIncludeBarsOutsideSession), _
                        3)
ChartStudyValue lBoll, BollingerBandsValueBottom
ChartStudyValue lBoll, BollingerBandsValueCentre
ChartStudyValue lBoll, BollingerBandsValueTop

Dim lMACD  As ResourceIdentifier
Set lMACD = DeclareStudy( _
                        "MACD", _
                        lTimeframe, _
                        "Short Periods=5" & _
                            ";Mov avg type=SMA" & _
                            ";Long Periods=34" & _
                            ";Smoothing Periods=5", _
                        GetParameterValue(ParamIncludeBarsOutsideSession), _
                        10)
ChartStudyValue lMACD, MACDValueMACD
ChartStudyValue lMACD, MACDValueMACDHist
ChartStudyValue lMACD, MACDValueMACDSignal

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStrategy_Name() As String
IStrategy_Name = "MACDStrategy21"
End Property

Private Sub IStrategy_NotifyNoLivePositions()
Const ProcName As String = "IStrategy_NotifyNoLivePositions"
On Error GoTo Err

checkTradingOpportunity

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategy_NotifyNoSimulatedPositions()
Const ProcName As String = "IStrategy_NotifyNoSimulatedPositions"
On Error GoTo Err

checkTradingOpportunity

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategy_NotifySimulatedTradingReadinessChange()

End Sub

Private Sub IStrategy_NotifyTick(ByVal pType As TickTypes)
Const ProcName As String = "IStrategy_NotifyTick"
On Error GoTo Err

If IsEndOfSession Then Exit Sub
If Not IsCurrentTickTrade Then Exit Sub
If Not isNewBar Then Exit Sub

checkTradingOpportunity

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategy_NotifyTradingReadinessChange()

End Sub

' occurs after historic price data has been loaded,
' but before the first tick is notified
Public Sub IStrategy_Start(ByVal pTradingContext As TradingContext)
Const ProcName As String = "IStrategy_Start"
On Error GoTo Err

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
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Private Function CancelPendingBuyOrders() As Boolean
Const ProcName As String = "CancelPendingBuyOrders"
On Error GoTo Err

If PendingBuySize = 0 And PendingSimulatedBuySize = 0 Then Exit Function

LogTradeMessage "Previous signal cancelled - cancelling pending buy orders"

If PendingSimulatedBuySize <> 0 Then CancelBracketOrder False, mBracketOrderBuyDummy
If PendingBuySize <> 0 Then CancelBracketOrder

CancelPendingBuyOrders = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function CancelPendingSellOrders() As Boolean
Const ProcName As String = "CancelPendingSellOrders"
On Error GoTo Err

If PendingSellSize = 0 And PendingSimulatedSellSize = 0 Then Exit Function

LogTradeMessage "Previous signal cancelled - cancelling pending sell orders"

If PendingSimulatedSellSize = 0 Then CancelBracketOrder False, mBracketOrderSellDummy
If PendingSellSize <> 0 Then CancelBracketOrder

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
LogTradeMessage lMessage

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IsConfirmedBuy(lMessage) As Boolean
Const ProcName As String = "IsConfirmedBuy"
On Error GoTo Err

If mCurrentBarNumber - SwingLowBarNumber() <= 3 Then
    IsConfirmedBuy = True
ElseIf PotentialSwingLow() = MaxDouble Then
    lMessage = lMessage & getRejectMessage("more than 3 Bars since swing low at ") & FormatTimestamp(SwingLowTime(), TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
    IsConfirmedBuy = False
ElseIf mCurrentBarNumber - PotentialSwingLowBarNumber() > 3 Then
    lMessage = lMessage & getRejectMessage("more than 3 Bars since potential swing low at ") & FormatTimestamp(PotentialSwingLowTime(), TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
    IsConfirmedBuy = False
ElseIf mCurrentBarNumber - PotentialSwingLowBarNumber() = 1 Then
    lMessage = lMessage & getRejectMessage("signal bar is potential swing low")
    IsConfirmedBuy = False
ElseIf mCurrentBarNumber = PotentialSwingLowBarNumber() Then
    lMessage = lMessage & getRejectMessage("potential new swing low this bar")
    IsConfirmedBuy = False
Else
    IsConfirmedBuy = True
End If

If BollingerInTopBand(HighValue(-1)) Then
    lMessage = lMessage & getRejectMessage("signal bar high is in top Bollinger band")
    IsConfirmedBuy = False
End If
If MinutesToEndOfTradingSession <= 30# Then
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

If mCurrentBarNumber - SwingHighBarNumber() <= 3 Then
    IsConfirmedSell = True
ElseIf PotentialSwingHigh() = MinDouble Then
    lMessage = lMessage & getRejectMessage("more than 3 Bars since swing high at ") & FormatTimestamp(SwingHighTime(), TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
    IsConfirmedSell = False
ElseIf mCurrentBarNumber - PotentialSwingHighBarNumber() > 3 Then
    lMessage = lMessage & getRejectMessage("more than 3 Bars since potential swing high at ") & FormatTimestamp(PotentialSwingHighTime(), TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
    IsConfirmedSell = False
ElseIf mCurrentBarNumber - PotentialSwingHighBarNumber() = 1 Then
    lMessage = lMessage & getRejectMessage("signal bar is potential swing high")
    IsConfirmedSell = False
ElseIf mCurrentBarNumber = PotentialSwingHighBarNumber() Then
    lMessage = lMessage & getRejectMessage("potential new swing high this bar")
    IsConfirmedSell = False
Else
    IsConfirmedSell = True
End If

If BollingerInBottomBand(LowValue(-1)) Then
    lMessage = lMessage & getRejectMessage("signal bar low is in bottom Bollinger band")
    IsConfirmedSell = False
End If
If MinutesToEndOfTradingSession <= 30# Then
    lMessage = lMessage & getRejectMessage("within last half-hour of session")
    IsConfirmedSell = False
End If

If IsConfirmedSell Then lMessage = lMessage & vbCrLf & Space(Len(MessagePrefix)) & "Confirmed sell"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IsEndOfSession() As Boolean
Const ProcName As String = "IsEndOfSession"
On Error GoTo Err

Static sSessionEnding As Boolean

If MinutesToEndOfTradingSession >= 1# Then
    sSessionEnding = False
    Exit Function
End If

If sSessionEnding Then
    IsEndOfSession = True
    Exit Function
End If

sSessionEnding = True
LogTradeMessage "Approaching end of session: closing any outstanding position"
CloseAllPositions
CloseAllSimulatedPositions
IsEndOfSession = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IsInEstablishedPosition() As Boolean
Const ProcName As String = "IsInEstablishedPosition"
On Error GoTo Err

IsInEstablishedPosition = TotalSimulatedPositionSize <> 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IsPendingBuyOrders() As Boolean
Const ProcName As String = "IsPendingBuyOrders"
On Error GoTo Err

IsPendingBuyOrders = PendingBuySize <> 0 And PendingSimulatedBuySize <> 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IsPendingSellOrders() As Boolean
Const ProcName As String = "IsPendingSellOrders"
On Error GoTo Err

IsPendingSellOrders = PendingSellSize <> 0 And PendingSimulatedSellSize <> 0

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
LogTradeMessage lMessage

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub ModifyPendingBuyOrders()
Const ProcName As String = "ModifyPendingBuyOrders"
On Error GoTo Err

If Not CanTrade Then
    LogTradeMessage "ModifyPendingBuyOrders: can't trade"
    Exit Sub
End If

Dim lTriggerPrice As Double
lTriggerPrice = getTriggerPriceLong
If lTriggerPrice >= EntryTriggerPrice(mBracketOrderBuyDummy) Then
    LogTradeMessage "ModifyPendingBuyOrders: no modification needed"
    Exit Sub
End If

Dim lStopPrice As Double
Dim lBreakevenPrice As Double
Dim lCancelPrice As Double
Dim lRisk As Double

calcPricesForBuy lTriggerPrice, lStopPrice, lBreakevenPrice, lCancelPrice, lRisk
SetNewEntryTriggerPrice lTriggerPrice, mBracketOrderBuyDummy
SetNewStopLossTriggerPrice lStopPrice, mBracketOrderBuyDummy

LogTradeMessage "ModifyPendingBuyOrders simulated"
Update mBracketOrderBuyDummy
    
If PendingBuySize = 0 Then Exit Sub

SetNewQuantity calcTradeSize(lRisk)
SetNewEntryTriggerPrice lTriggerPrice
SetNewEntryLimitPrice lTriggerPrice - mEntryLimitOffsetTicks * TickSize
SetNewStopLossTriggerPrice lStopPrice
LogTradeMessage "ModifyPendingBuyOrders"
Update

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ModifyPendingSellOrders()
Const ProcName As String = "ModifyPendingSellOrders"
On Error GoTo Err

If Not CanTrade Then
    LogTradeMessage "ModifyPendingSellOrders: can't trade"
    Exit Sub
End If

Dim lTriggerPrice As Double
lTriggerPrice = getTriggerPriceShort
If lTriggerPrice <= EntryTriggerPrice(mBracketOrderSellDummy) Then
    LogTradeMessage "ModifyPendingSellOrders: no modification needed"
    Exit Sub
End If

Dim lStopPrice As Double
Dim lBreakevenPrice As Double
Dim lCancelPrice As Double
Dim lRisk As Double

calcPricesForSell lTriggerPrice, lStopPrice, lBreakevenPrice, lCancelPrice, lRisk
SetNewEntryTriggerPrice lTriggerPrice, mBracketOrderSellDummy
SetNewStopLossTriggerPrice lStopPrice, mBracketOrderSellDummy

LogTradeMessage "ModifyPendingSellOrders simulated"
Update mBracketOrderSellDummy
    
If PendingSellSize = 0 Then Exit Sub

SetNewQuantity calcTradeSize(lRisk)
SetNewEntryTriggerPrice lTriggerPrice
SetNewEntryLimitPrice lTriggerPrice + mEntryLimitOffsetTicks * TickSize
SetNewStopLossTriggerPrice lStopPrice
LogTradeMessage "ModifyPendingSellOrders"
Update

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

        
Private Sub PlaceBuyOrders()
Const ProcName As String = "PlaceBuyOrders"
On Error GoTo Err

If Not (CanTrade And CanTradeSimulated) Then
    LogTradeMessage "PlaceBuyOrders: can't trade"
    Exit Sub
End If

Assert ActiveBuySize = 0 And PendingBuySize = 0, "ActiveBuySize <> 0 or PendingBuySize <> 0"

Dim lTriggerPrice As Double
lTriggerPrice = getTriggerPriceLong

Dim lStopPrice As Double
Dim lBreakevenPrice As Double
Dim lCancelPrice As Double
Dim lRisk As Double

If Not calcPricesForBuy(lTriggerPrice, lStopPrice, lBreakevenPrice, lCancelPrice, lRisk) Then
    Exit Sub
End If

Dim lSize As BoxedDecimal: Set lSize = calcTradeSize(lRisk)

LogTradeMessage "PlaceBuyOrders simulated"
Set mBracketOrderBuyDummy = BuySimulated( _
                                Quantity:=DecimalOne, _
                                EntryOrder:=DeclareSimulatedEntryOrder(EntryOrderTypeStop, TriggerPrice:=lTriggerPrice), _
                                StopLossOrder:=DeclareSimulatedStopLossOrder(StopLossOrderTypeStop, TriggerPrice:=lStopPrice), _
                                CancelPrice:=lCancelPrice, _
                                NotifyCompletion:=True)

Assert Not mBracketOrderBuyDummy Is Nothing, "Failed to create dummy buy bracket oder"

If lSize = 0 Then Exit Sub

LogTradeMessage "PlaceBuyOrders"
SetPrimaryBracketOrder Buy( _
                            Quantity:=lSize, _
                            EntryOrder:=DeclareEntryOrder( _
                                                        EntryOrderTypeStopLimit, _
                                                        LimitPrice:=lTriggerPrice, _
                                                        TriggerPrice:=lTriggerPrice), _
                            StopLossOrder:=DeclareStopLossOrder( _
                                                        StopLossOrderTypeStop, _
                                                        TriggerPrice:=lStopPrice, _
                                                        TimeInForce:=OrderTIFs.OrderTIFGoodTillCancelled), _
                            CancelPrice:=lCancelPrice, _
                            CancelAfter:=300, _
                            NotifyCompletion:=True)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub PlaceSellOrders()
Const ProcName As String = "PlaceSellOrders"
On Error GoTo Err

If Not (CanTrade And CanTradeSimulated) Then
    LogTradeMessage "PlaceSellOrders: can't trade"
    Exit Sub
End If

Assert ActiveSellSize = 0 And PendingSellSize = 0, "ActiveSellSize <> 0 or PendingSellSize <> 0"

Dim lTriggerPrice As Double
lTriggerPrice = getTriggerPriceShort

Dim lStopPrice As Double
Dim lBreakevenPrice As Double
Dim lCancelPrice As Double
Dim lRisk As Double

If Not calcPricesForSell(lTriggerPrice, lStopPrice, lBreakevenPrice, lCancelPrice, lRisk) Then
    Exit Sub
End If

Dim lSize As BoxedDecimal: Set lSize = calcTradeSize(lRisk)

LogTradeMessage "PlaceSellOrders simulated"
Set mBracketOrderSellDummy = SellSimulated( _
                                Quantity:=DecimalOne, _
                                EntryOrder:=DeclareSimulatedEntryOrder(EntryOrderTypeStop, TriggerPrice:=lTriggerPrice), _
                                StopLossOrder:=DeclareSimulatedStopLossOrder(StopLossOrderTypeStop, TriggerPrice:=lStopPrice), _
                                CancelPrice:=lCancelPrice, _
                                NotifyCompletion:=True)

Assert Not mBracketOrderSellDummy Is Nothing, "Failed to create dummy sell bracket oder"

If lSize = 0 Then Exit Sub

LogTradeMessage "PlaceSellOrders"
SetPrimaryBracketOrder Sell( _
                            Quantity:=lSize, _
                            EntryOrder:=DeclareEntryOrder( _
                                                        EntryOrderTypeStopLimit, _
                                                        LimitPrice:=lTriggerPrice, _
                                                        TriggerPrice:=lTriggerPrice), _
                            StopLossOrder:=DeclareStopLossOrder( _
                                                        StopLossOrderTypeStop, _
                                                        TriggerPrice:=lStopPrice, _
                                                        TimeInForce:=OrderTIFGoodTillCancelled), _
                            CancelPrice:=lCancelPrice, _
                            CancelAfter:=300, _
                            NotifyCompletion:=True)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcPricesForBuy( _
                ByVal pTriggerPrice As Double, _
                ByRef pStopPrice As Double, _
                ByRef pBreakevenPrice As Double, _
                ByRef pCancelPrice As Double, _
                ByRef pRisk As Double) As Boolean
Const ProcName As String = "calcPricesForBuy"
On Error GoTo Err

Dim lSwingLow As Double
lSwingLow = IIf(SwingLow < PotentialSwingLow, SwingLow, PotentialSwingLow)
If mSwingToBreakeven <> 0 Then
    If (pTriggerPrice - lSwingLow) > 0.5 * mSwingToBreakeven Then
        LogTradeMessage "    reject: trigger price too close to breakeven price"
        Exit Function
    End If
    pBreakevenPrice = lSwingLow + mSwingToBreakeven
End If

pStopPrice = lSwingLow - mStopBreakoutThresholdTicks * TickSize

If pStopPrice < pTriggerPrice - mMaxInitialStop Then
    pStopPrice = pTriggerPrice - mMaxInitialStop
End If

pRisk = (pTriggerPrice + mEntryLimitOffsetTicks * TickSize - pStopPrice) * TickValue / TickSize

If PotentialSwingLow = MaxDouble Then
    pCancelPrice = SwingLow
Else
    pCancelPrice = PotentialSwingLow
End If

calcPricesForBuy = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcPricesForSell( _
                ByVal pTriggerPrice As Double, _
                ByRef pStopPrice As Double, _
                ByRef pBreakevenPrice As Double, _
                ByRef pCancelPrice As Double, _
                ByRef pRisk As Double) As Boolean
Const ProcName As String = "calcPricesForSell"
On Error GoTo Err

Dim lSwingHigh As Double
lSwingHigh = IIf(SwingHigh > PotentialSwingHigh, SwingHigh, PotentialSwingHigh)
If mSwingToBreakeven <> 0 Then
    If (lSwingHigh - pTriggerPrice) > 0.5 * mSwingToBreakeven Then
        LogTradeMessage "    reject: trigger price too close to breakeven price"
        Exit Function
    End If
    pBreakevenPrice = lSwingHigh - mSwingToBreakeven
End If

pStopPrice = lSwingHigh + mStopBreakoutThresholdTicks * TickSize

If pStopPrice > pTriggerPrice + mMaxInitialStop Then
    pStopPrice = pTriggerPrice + mMaxInitialStop
End If

pRisk = (pStopPrice - pTriggerPrice - mEntryLimitOffsetTicks * TickSize) * TickValue / TickSize

If PotentialSwingHigh = MinDouble Then
    pCancelPrice = SwingHigh
Else
    pCancelPrice = PotentialSwingHigh
End If

calcPricesForSell = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcTradeSize(ByVal risk As Double) As BoxedDecimal
Const ProcName As String = "calcTradeSize"
On Error GoTo Err

Dim lSize As Long
lSize = Int((mRiskUnitPercent * AccountBalance / 100) / risk)
If lSize > mMaxTradeSize Then lSize = mMaxTradeSize
If lSize = 0 Then LogTradeMessage "    reject: allowable initial risk exceeded"

Set calcTradeSize = CreateBoxedDecimal(lSize)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub checkTradingOpportunity()
Const ProcName As String = "checkTradingOpportunity"
On Error GoTo Err

If IsEndOfSession Then Exit Sub
If IsInEstablishedPosition Then Exit Sub

If ActiveBuySize <> 0 Or ActiveSimulatedBuySize <> 0 Then Exit Sub
If ActiveSellSize <> 0 Or ActiveSimulatedSellSize <> 0 Then Exit Sub

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
lMACDUp = MACDValueUp(ref)

Dim lMACDDown As Boolean
lMACDDown = MACDValueDown(ref)

Dim lMACDHistUp As Boolean
lMACDHistUp = MACDHistUp(ref)

Dim lMACDHistDown As Boolean
lMACDHistDown = MACDHistDown(ref)

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
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    getColour = NoColour
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Function

Private Function getRejectMessage(pReason As String) As String
getRejectMessage = vbCrLf & Space(Len(MessagePrefix)) & "reject: " & pReason
End Function

Private Function getTriggerPriceLong() As Double
getTriggerPriceLong = CloseValue(-1) + mEntryBreakoutThresholdTicks * TickSize
End Function

Private Function getTriggerPriceShort() As Double
getTriggerPriceShort = CloseValue(-1) - mEntryBreakoutThresholdTicks * TickSize
End Function

Private Function isNewBar() As Boolean
Dim lBarNumber As Long: lBarNumber = CurrentBarNumber
If lBarNumber <> mCurrentBarNumber Then
    mCurrentBarNumber = lBarNumber
    isNewBar = True
End If
End Function

Private Sub processParameters()
Const ProcName As String = "processParameters"
On Error GoTo Err

mEntryBreakoutThresholdTicks = GetParameterValue(ParamEntryBreakoutThresholdTicks)
mStopBreakoutThresholdTicks = GetParameterValue(ParamBreakoutThresholdTicks)
mEntryLimitOffsetTicks = GetParameterValue(ParamEntryLimitOffsetTicks)
mRiskUnitPercent = GetParameterValue(ParamRiskUnitPercent)
mMaxTradeSize = GetParameterValue(ParamMaxTradeSize)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub





