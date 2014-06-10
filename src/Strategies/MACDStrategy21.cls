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

Private Const MAX_INITIAL_STOP_TICKS = 100
    ' max number of ticks from entry price for initial stop loss

Private Const STOP_BREAKOUT_THRESHOLD_TICKS = 1
    ' number of ticks below/above a low/high to set a stop

Private Const ENTRY_BREAKOUT_THRESHOLD_TICKS = 4

Private Const BOLL_CENTRE_BAND_WIDTH_TICKS As Single = 20
Private Const BOLL_EDGE_BAND_WIDTH_TICKS As Single = 10

Private Const MINIMUM_SWING_TICKS As Integer = 10

Private Const ENTRY_LIMIT_OFFSET_TICKS As Integer = 1

Private Const BreakevenBarLength As Long = 0
Private Const BreakevenBarThresholdTicks As Long = 0
Private Const breakEvenThresholdTicks As Long = 0
Private Const ConfirmedBarsForBreakEven As Long = 0
Private Const ConfirmedBarsForSwing As Long = 0
Private Const ConfirmedBarsForTrail As Long = 0
Private Const ConfirmedBarsForExit As Long = 0
Private Const RetracementStopThresholdTicks As Long = 0
Private Const RetracementStopPercent As Single = 0
Private Const RetraceFromExtremes As String = "Y"

Private Const SwingToBreakevenTicks As Long = 0

Private Const ATRPeriods As Long = 20
Private Const MaxIncrements As Long = 0
Private Const RiskUnitPercent As Double = 1#
Private Const RiskIncrementPercent As Double = 0.5
Private Const ScaleThresholdFactor As Double = 0.5
Private Const StopIncrementFactor As Double = 0.5
Private Const InitialStopFactor As Double = 2#
Private Const MaxTradeSize As Long = 1

Private Const BackstopMABarLength As Long = 0
Private Const BackstopMAPeriods As Long = 25
Private Const BackstopMAOffset As Long = 8

Private Const ParamBarLength                        As String = "Bar Length"
Private Const ParamTimePeriodUnit                   As String = "Time Period Unit"
Private Const ParamIncludeBarsOutsideSession        As String = "Include Bars Outside Session"
Private Const ParamBollingerPeriods                 As String = "Bollinger Periods"
Private Const ParamBollingerMovingAverageType       As String = "Bollinger Moving Avg Type"
Private Const ParamMinimumSwingTicks                As String = "Minimum Swing Ticks"
Private Const ParamRetraceFromExtremes              As String = "Retrace From Extremes"
Private Const ParamEntryLimitOffsetTicks            As String = "Entry Limit Offset Ticks"
Private Const ParamMaxInitialStopTicks              As String = "Max Initial Stop Ticks"
Private Const ParamEntryBreakoutThresholdTicks      As String = "Entry Breakout Threshold Ticks"
Private Const ParamStopBreakoutThresholdTicks       As String = "Stop Breakout Threshold Ticks"
Private Const ParamSwingToBreakevenTicks            As String = "Swing To Breakeven Ticks"
Private Const ParamATRPeriods                       As String = "ATR Periods"
Private Const ParamMaxIncrements                    As String = "Max Increments"
Private Const ParamRiskUnitPercent                  As String = "Risk Unit Percent"
Private Const ParamRiskIncrementPercent             As String = "Risk Increment Percent"
Private Const ParamScaleThresholdFactor             As String = "Scale Threshold Factor"
Private Const ParamStopIncrementFactor              As String = "Stop Increment Factor"
Private Const ParamInitialStopFactor                As String = "Initial Stop Factor"
Private Const ParamMaxTradeSize                     As String = "Max Trade Size"
Private Const ParamBackstopMABarLength              As String = "Backstop MA BarLength"
Private Const ParamBackstopMAPeriods                As String = "Backstop MA Periods"
Private Const ParamBackstopMAOffset                 As String = "Backstop MA Offset"
Private Const ParamBreakEvenBarLength               As String = "BreakEven Bar Length"
Private Const ParamRetracementStopPercent           As String = "Retracement Stop Percent"
Private Const ParamRetracementStopThresholdTicks    As String = "Retracement Stop Threshold Ticks"
Private Const ParamBreakEvenBarThresholdTicks       As String = "BreakEven Bar Threshold Ticks"
Private Const ParamBreakEvenThresholdTicks          As String = "BreakEven Threshold Ticks"
Private Const ParamConfirmedBarsForBreakEven        As String = "Confirmed Bars For BreakEven"
Private Const ParamConfirmedBarsForSwing            As String = "Confirmed Bars For Swing"
Private Const ParamConfirmedBarsForTrail            As String = "Confirmed Bars For Trail"
Private Const ParamConfirmedBarsForExit             As String = "Confirmed Bars For Exit"
Private Const ParamBollingerCentreBandWidthTicks    As String = "Bollinger Centre Band Width Ticks"
Private Const ParamBollingerEdgeBandWidthTicks      As String = "Bollinger Edge Band Width Ticks"

'@================================================================================
' Member variables
'@================================================================================

Private mContext As TradingContext
Private mPositionManager As PositionManager
Private WithEvents mPositionManagerSimulated As PositionManager
Attribute mPositionManagerSimulated.VB_VarHelpID = -1
Private mParams As Parameters

Private mOrderContextBuySimulated As OrderContext
Private mOrderContextSellSimulated As OrderContext
Private WithEvents mBracketOrderBuySimulated As BracketOrder
Attribute mBracketOrderBuySimulated.VB_VarHelpID = -1
Private WithEvents mBracketOrderSellSimulated As BracketOrder
Attribute mBracketOrderSellSimulated.VB_VarHelpID = -1
Private mOrderContextBuy As OrderContext
Private mOrderContextSell As OrderContext
Private WithEvents mBracketOrderBuy As BracketOrder
Attribute mBracketOrderBuy.VB_VarHelpID = -1
Private WithEvents mBracketOrderSell As BracketOrder
Attribute mBracketOrderSell.VB_VarHelpID = -1

Private mStopStrategyFactory As IStrategyFactory
Private mScalingStrategyFactory As IStrategyFactory

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
Private mGotFirstBidAndOffer As Boolean

' Scaling strategy studies
Private mATRStudy As IStudy

Private mATRPeriods As Long
Private mMaxIncrements As Long
Private mRiskUnitPercent As Double
Private mRiskIncrementPercent As Double
Private mScaleThresholdFactor As Double
Private mStopIncrementFactor As Double
Private mInitialStopFactor As Double
Private mMaxTradeSize As Long

Private mBackstopMABarLength As Integer
Private mBackstopMAPeriods As Integer
Private mBackstopMAOffset As Integer

Private mSwingToBreakeven As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mPrevColour = NoColour
mCurrColour = NoColour

' set default parameters
Set mParams = New Parameters
mParams.SetParameterValue ParamBarLength, "1"
mParams.SetParameterValue ParamTimePeriodUnit, "min"
mParams.SetParameterValue ParamBollingerPeriods, "34"
mParams.SetParameterValue ParamBollingerMovingAverageType, "SMA"
mParams.SetParameterValue ParamMinimumSwingTicks, MINIMUM_SWING_TICKS
mParams.SetParameterValue ParamRetraceFromExtremes, RetraceFromExtremes
mParams.SetParameterValue ParamEntryLimitOffsetTicks, ENTRY_LIMIT_OFFSET_TICKS
mParams.SetParameterValue ParamMaxInitialStopTicks, MAX_INITIAL_STOP_TICKS
mParams.SetParameterValue ParamEntryBreakoutThresholdTicks, ENTRY_BREAKOUT_THRESHOLD_TICKS
mParams.SetParameterValue ParamStopBreakoutThresholdTicks, STOP_BREAKOUT_THRESHOLD_TICKS
mParams.SetParameterValue ParamSwingToBreakevenTicks, SwingToBreakevenTicks
mParams.SetParameterValue ParamATRPeriods, ATRPeriods
mParams.SetParameterValue ParamMaxIncrements, MaxIncrements
mParams.SetParameterValue ParamRiskUnitPercent, RiskUnitPercent
mParams.SetParameterValue ParamRiskIncrementPercent, RiskIncrementPercent
mParams.SetParameterValue ParamScaleThresholdFactor, ScaleThresholdFactor
mParams.SetParameterValue ParamStopIncrementFactor, StopIncrementFactor
mParams.SetParameterValue ParamInitialStopFactor, InitialStopFactor
mParams.SetParameterValue ParamMaxTradeSize, MaxTradeSize
mParams.SetParameterValue ParamBackstopMABarLength, BackstopMABarLength
mParams.SetParameterValue ParamBackstopMAPeriods, BackstopMAPeriods
mParams.SetParameterValue ParamBackstopMAOffset, BackstopMAOffset
mParams.SetParameterValue ParamBreakEvenBarLength, BreakevenBarLength
mParams.SetParameterValue ParamRetracementStopPercent, RetracementStopPercent
mParams.SetParameterValue ParamRetracementStopThresholdTicks, RetracementStopThresholdTicks
mParams.SetParameterValue ParamBreakEvenBarThresholdTicks, BreakevenBarThresholdTicks
mParams.SetParameterValue ParamBreakEvenThresholdTicks, breakEvenThresholdTicks
mParams.SetParameterValue ParamConfirmedBarsForBreakEven, ConfirmedBarsForBreakEven
mParams.SetParameterValue ParamConfirmedBarsForSwing, ConfirmedBarsForSwing
mParams.SetParameterValue ParamConfirmedBarsForTrail, ConfirmedBarsForTrail
mParams.SetParameterValue ParamConfirmedBarsForExit, ConfirmedBarsForExit
mParams.SetParameterValue ParamBollingerCentreBandWidthTicks, BOLL_CENTRE_BAND_WIDTH_TICKS
mParams.SetParameterValue ParamBollingerEdgeBandWidthTicks, BOLL_EDGE_BAND_WIDTH_TICKS

processParameters mParams
End Sub

Private Sub Class_Terminate()
Debug.Print "MACDStrategy21 terminated"
End Sub

'@================================================================================
' IStrategy Interface Members
'@================================================================================

Private Sub IStrategy_NotifyTick(ByVal pType As TickTypes)
Dim colour As Colours

mTimestamp = mContext.Timestamp

If mTimestamp >= mLastMinute Then
    mPositionManager.ClosePositions
    Exit Sub
End If

If pType <> TickTypes.TickTypeTrade Then Exit Sub

If mBars.Bar.TickVolume = 1 Or (Not mGotFirstBidAndOffer) Then

    ' first make sure we don't attempt any trades before the
    ' first bid and offer have been received
    If Not mGotFirstBidAndOffer Then
        If mContext.BidQuote.price > 0 And mContext.AskQuote.price > 0 Then
            mGotFirstBidAndOffer = True
        Else
            Exit Sub
        End If
    End If
    
    mBarNumber = mBars.Bar.barNumber
    mNoTradeThisBar = False
    
    colour = getColour(1)
    If colour <> mCurrColour Then
        mPrevColour = mCurrColour
        mCurrColour = colour
    End If

    If mCurrColour = Red Or _
        (mCurrColour = Pink And mPrevColour = Green) _
    Then
        If mOrderContextBuySimulated.PendingSize <> 0 Then
            mBracketOrderBuySimulated.Cancel
            If mOrderContextBuy.PendingSize <> 0 Then
                mBracketOrderBuy.Cancel
            End If
        End If
    ElseIf mCurrColour = Green Or _
            (mCurrColour = Amber And mPrevColour = Red) _
    Then
        If mOrderContextSellSimulated.PendingSize <> 0 Then
            mBracketOrderSellSimulated.Cancel
            If mOrderContextSell.PendingSize <> 0 Then
                mBracketOrderSell.Cancel
            End If
        End If
    End If
    
    checkTradingOpportunity

End If
End Sub

Public Sub IStrategy_Initialise( _
                ByVal pContext As InitialisationContext)
processParameters mParams

Dim lContract As IContract
Set lContract = pContext.Contract
mTickSize = lContract.TickSize
mTickValue = lContract.TickValue
mLastMinute = lContract.SessionEndTime - (1 / 1440)
mLastHalfHour = lContract.SessionEndTime - (30 / 1440)

mMaxInitialStop = mParams.GetParameterValue(ParamMaxInitialStopTicks) * mTickSize
mSwingToBreakeven = mParams.GetParameterValue(ParamSwingToBreakevenTicks) * mTickSize

'--------------------------------------------------------------------------------
' Studies for the normal timeframe
'--------------------------------------------------------------------------------
Set mTimeframe = pContext.AddTimeframe(GetTimePeriod(mParams.GetParameterValue(ParamBarLength), _
                                                mParams.GetParameterValue(ParamTimePeriodUnit)), _
                                500, _
                                mParams.GetParameterValue(ParamIncludeBarsOutsideSession))

ReDim lInputValueNames(0) As String

Dim params
Set params = New Parameters
params.SetParameterValue "Minimum swing (ticks)", mParams.GetParameterValue(ParamMinimumSwingTicks) * mTickSize
params.SetParameterValue "Include implicit swing points", "N"
lInputValueNames(0) = "$default"
Set mSwing = pContext.AddStudy("Swing", _
                            mTimeframe, _
                            lInputValueNames, _
                            params, _
                            3)

Set params = New Parameters
params.SetParameterValue "Periods", mParams.GetParameterValue(ParamBollingerPeriods)
params.SetParameterValue "Standard deviations", 2
params.SetParameterValue "Mov avg type", mParams.GetParameterValue(ParamBollingerMovingAverageType)
params.SetParameterValue "Centre band width", mParams.GetParameterValue(ParamBollingerCentreBandWidthTicks)
params.SetParameterValue "Edge band width", mParams.GetParameterValue(ParamBollingerEdgeBandWidthTicks)
lInputValueNames(0) = "$default"
Set mBoll = pContext.AddStudy("BollingerBands", _
                        mTimeframe, _
                        lInputValueNames, _
                        params, _
                        3)

Set params = New Parameters
params.SetParameterValue "Short Periods", 5
params.SetParameterValue "Mov avg type", "SMA"
params.SetParameterValue "Long Periods", 34
params.SetParameterValue "Smoothing Periods", 5
lInputValueNames(0) = "$default"
Set mMACD = pContext.AddStudy("MACD", _
                        mTimeframe, _
                        lInputValueNames, _
                        params, _
                        10)

'--------------------------------------------------------------------------------
' Studies for the stop strategy
'--------------------------------------------------------------------------------

Set mStopStrategyFactory = New StopStrategyFactory5
mStopStrategyFactory.Initialise pContext

'--------------------------------------------------------------------------------
' Studies for the scaling strategy
'--------------------------------------------------------------------------------

Set mScalingStrategyFactory = New ScaleStrategyFactory1
mScalingStrategyFactory.Initialise pContext

'Set params = New Parameters
'params.SetParameterValue "Periods", mATRPeriods
'lInputValueNames(0) = "$default"
'Set mATRStudy = pContext.AddStudy("AverageTrueRange", _
'                            mTimeframe, _
'                            lInputValueNames, _
'                            params, _
'                            2)

'--------------------------------------------------------------------------------
' Other initialisation
'--------------------------------------------------------------------------------

Set mOrderContextBuySimulated = mPositionManagerSimulated.OrderContexts.Add("BUY")
Set mOrderContextSellSimulated = mPositionManagerSimulated.OrderContexts.Add("SELL")
Set mOrderContextBuy = mPositionManager.OrderContexts.Add("BUY")
Set mOrderContextSell = mPositionManager.OrderContexts.Add("SELL")

End Sub

Private Property Get IStrategy_Parameters() As Parameters
Set IStrategy_Parameters = mParams
End Property

' occurs after historic price data has been loaded,
' but before the first tick is notified
Public Sub IStrategy_start(ByVal pTradingContext As TradingContext)
Set mContext = pTradingContext

Set mBars = mContext.GetBars(mTimeframe)

mCurrColour = getColour(1)
mPrevColour = mCurrColour

Dim i As Integer
For i = 2 To 10
    mPrevColour = getColour(i)
    If mPrevColour <> mCurrColour Then Exit For
Next
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
Dim triggerPrice As Double

If mNoTradeThisBar Then Exit Sub
If mPositionManagerSimulated.PositionSize <> 0 Then Exit Sub
If mCurrColour = Red Or _
    (mCurrColour = Pink And mPrevColour = Green) _
Then
    triggerPrice = mBars.CloseValue(1) - mEntryBreakoutThresholdTicks * mTickSize
    sell triggerPrice, _
        "signs of weakness"
ElseIf mCurrColour = Green Or _
        (mCurrColour = Amber And mPrevColour = Red) _
Then
    triggerPrice = mBars.CloseValue(1) + mEntryBreakoutThresholdTicks * mTickSize
    buy triggerPrice, _
        "signs of strength"
End If
End Sub

Private Function getColour(ByVal ref As Integer) As Colours
Dim lMACDUp As Boolean
Dim lMACDDown As Boolean
Dim lMACDHistUp As Boolean
Dim lMACDHistDown As Boolean

lMACDUp = mMACD.MACDUp(ref).value
lMACDDown = mMACD.MACDDown(ref).value
lMACDHistUp = mMACD.MACDHistUp(ref).value
lMACDHistDown = mMACD.MACDHistDown(ref).value

If lMACDUp And lMACDHistUp Then
    getColour = Green
ElseIf lMACDUp And lMACDHistDown Then
    getColour = Pink
ElseIf lMACDDown And lMACDHistUp Then
    getColour = Amber
Else
    getColour = Red
End If
End Function

Private Function okToBuy(ByVal triggerPrice As Double) As Boolean

okToBuy = True

If mBarNumber - mSwing.SwingLow(1).barNumber > 3 Then
    If mSwing.SwingLow(0).value = MaxDouble Then
        mContext.LogTradeReason "reject: more than 3 Bars since swing low"
        okToBuy = False
    ElseIf mBarNumber - mSwing.SwingLow(0).barNumber > 3 Then
        mContext.LogTradeReason "reject: more than 3 Bars since potential swing low"
        okToBuy = False
    ElseIf mBarNumber - mSwing.SwingLow(0).barNumber = 1 Then
        mContext.LogTradeReason "reject: signal bar is potential swing low"
        okToBuy = False
    ElseIf mBarNumber = mSwing.SwingLow(0).barNumber Then
        mContext.LogTradeReason "reject: potential new swing low this bar"
        okToBuy = False
    End If
End If

If mBoll.InTopBand(mBars.HighValue(1)).value Then
    mContext.LogTradeReason "reject: signal bar high is in top Bollinger band"
    okToBuy = False
End If

'If mContext.lastPrice > mBoll.Top Then
'    mcontext.LogTradeReason "reject: price is above top Bollinger band"
'    okToBuy = False
'End If

If mTimestamp >= mLastHalfHour Then
    mContext.LogTradeReason "reject: within last half-hour of session"
    okToBuy = False
End If

If Not okToBuy Then mNoTradeThisBar = True
End Function

Private Function okToSell(ByVal triggerPrice As Double) As Boolean
okToSell = True

If mBarNumber - mSwing.SwingHigh(1).barNumber > 3 Then
    If mSwing.SwingHigh(0).value = MinDouble Then
        mContext.LogTradeReason "reject: more than 3 Bars since swing high"
        okToSell = False
    ElseIf mBarNumber - mSwing.SwingHigh(0).barNumber > 3 Then
        mContext.LogTradeReason "reject: more than 3 Bars since potential swing high"
        okToSell = False
    ElseIf mBarNumber - mSwing.SwingHigh(0).barNumber = 1 Then
        mContext.LogTradeReason "reject: signal bar is potential swing high"
        okToSell = False
    ElseIf mBarNumber = mSwing.SwingHigh(0).barNumber Then
        mContext.LogTradeReason "reject: potential new swing high this bar"
        okToSell = False
    End If
End If

If mBoll.InBottomBand(mBars.LowValue(1)).value Then
    mContext.LogTradeReason "reject: signal bar low is in bottom Bollinger band"
    okToSell = False
End If

'If mContext.lastPrice < mBoll.bottom Then
'    mcontext.LogTradeReason "reject: price is below bottom Bollinger band"
'    okToSell = False
'End If

If mTimestamp >= mLastHalfHour Then
    mContext.LogTradeReason "reject: within last half-hour of session"
    okToSell = False
End If

If Not okToSell Then mNoTradeThisBar = True
End Function

Private Function stillOkToBuy(ByVal triggerPrice As Double) As Boolean

stillOkToBuy = True

If mBarNumber - mSwing.SwingLow(1).barNumber > 3 Then
    If mSwing.SwingLow(0).value <> MaxDouble And _
        mBarNumber = mSwing.SwingLow(0).barNumber _
    Then
        mContext.LogTradeReason "reject: potential new swing low this bar"
        stillOkToBuy = False
    End If
End If

'If mBoll.intopBand(mBars.highPrice(1)) Then
'    mcontext.LogTradeReason "reject: signal bar high is in top Bollinger band"
'    stillOkToBuy = False
'End If

If mTimestamp >= mLastHalfHour Then
    mContext.LogTradeReason "reject: within last half-hour of session"
    stillOkToBuy = False
End If

If Not stillOkToBuy Then mNoTradeThisBar = True
End Function

Private Function stillOkToSell(ByVal triggerPrice As Double) As Boolean
stillOkToSell = True

If mBarNumber - mSwing.SwingHigh(1).barNumber > 3 Then
    If mSwing.SwingHigh(0).value <> MinDouble And _
        mBarNumber = mSwing.SwingHigh(0).barNumber _
    Then
        mContext.LogTradeReason "reject: potential new swing high this bar"
        stillOkToSell = False
    End If
End If

'If mBoll.inBottomBand(mBars.lowPrice(1)) Then
'    mcontext.LogTradeReason "reject: signal bar low is in bottom Bollinger band"
'    stillOkToSell = False
'End If

If mTimestamp >= mLastHalfHour Then
    mContext.LogTradeReason "reject: within last half-hour of session"
    stillOkToSell = False
End If

If Not stillOkToSell Then mNoTradeThisBar = True
End Function

Private Sub sell( _
    ByVal triggerPrice As Double, _
    ByVal reason As String)
Dim stopPrice As Double
Dim breakevenPrice As Double
Dim cancelPrice As Double
Dim risk As Double

If mOrderContextSellSimulated.PendingSize <> 0 Then
    If Not stillOkToSell(triggerPrice) Then
        mBracketOrderSellSimulated.Cancel
        If mOrderContextSell.PendingSize <> 0 Then mBracketOrderSell.Cancel
        Exit Sub
    End If
    
    If triggerPrice <= mBracketOrderSellSimulated.EntryOrder.triggerPrice Then Exit Sub
        
    calcPricesForSell triggerPrice, stopPrice, breakevenPrice, cancelPrice, risk
    mBracketOrderSellSimulated.SetNewEntryTriggerPrice triggerPrice
    mBracketOrderSellSimulated.SetNewStopLossPrice stopPrice
    On Error Resume Next
    mBracketOrderSellSimulated.Update
    If Err.Number = 0 Then
        On Error GoTo 0
        If mOrderContextSell.PendingSize = 0 Then Exit Sub
        mBracketOrderSell.SetNewQuantity UnitSize(risk)
        mBracketOrderSell.SetNewEntryTriggerPrice triggerPrice
        mBracketOrderSell.SetNewEntryPrice triggerPrice + mEntryLimitOffsetTicks * mTickSize
        mBracketOrderSell.SetNewStopLossPrice stopPrice
        On Error Resume Next
        mBracketOrderSell.Update
        If Err.Number = 0 Then Exit Sub
    End If
    ' the existing order plex cannot be amended - so cancel it
    ' and create a new one
    On Error GoTo 0
    mBracketOrderSellSimulated.Cancel
    If mOrderContextSell.PendingSize <> 0 Then mBracketOrderSell.Cancel
End If

mContext.LogTradeReason "Potential sell: " & reason

If Not okToSell(triggerPrice) Then Exit Sub

calcPricesForSell triggerPrice, stopPrice, breakevenPrice, cancelPrice, risk
If stopPrice = 0 Then
    mContext.LogTradeReason "reject: trigger price too close to breakeven price"
    Exit Sub
End If

Dim stopStrategySimulated As StopStrategy5
Set stopStrategySimulated = mStopStrategyFactory.CreateStrategy(mContext)
stopStrategySimulated.timePeriod = mParams.GetParameterValue(ParamBarLength)
stopStrategySimulated.TimePeriodUnit = mParams.GetParameterValue(ParamTimePeriodUnit)
stopStrategySimulated.minimumSwing = mParams.GetParameterValue(ParamMinimumSwingTicks) * mTickSize
stopStrategySimulated.breakoutThresholdTicks = mStopBreakoutThresholdTicks
stopStrategySimulated.RetraceFromExtremes = IIf(UCase$(mParams.GetParameterValue(ParamRetraceFromExtremes)) = "Y", True, False)

Set mBracketOrderSellSimulated = mOrderContextSellSimulated.sell( _
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
                    targetTriggerPrice:=0)
mBracketOrderSellSimulated.cancelPrice = cancelPrice
mContext.ApplyPositionManagementStrategy mBracketOrderSellSimulated, stopStrategySimulated

If mBracketOrderSellSimulated Is Nothing Then
    mNoTradeThisBar = True
        ' prevents each subsequent tick being processed if money management has prevented
        ' further trades
    Exit Sub
End If

If mOrderContextSell.ActiveSize <> 0 Or mOrderContextSell.PendingSize <> 0 Then Exit Sub

If UnitSize(risk) = 0 Then
    mContext.LogTradeReason "reject: allowable initial risk exceeded"
    Exit Sub
End If

Dim stopStrategy As StopStrategy5
Set stopStrategy = mStopStrategyFactory.CreateStrategy(mContext)
stopStrategy.timePeriod = mParams.GetParameterValue(ParamBarLength)
stopStrategy.TimePeriodUnit = mParams.GetParameterValue(ParamTimePeriodUnit)
stopStrategy.minimumSwing = mParams.GetParameterValue(ParamMinimumSwingTicks) * mTickSize
stopStrategy.breakoutThresholdTicks = mStopBreakoutThresholdTicks
stopStrategy.confirmedCountForBreakEven = mParams.GetParameterValue(ParamConfirmedBarsForBreakEven)
stopStrategy.RetraceFromExtremes = IIf(UCase$(mParams.GetParameterValue(ParamRetraceFromExtremes)) = "Y", True, False)
stopStrategy.BreakevenBarLength = mParams.GetParameterValue(ParamBreakEvenBarLength)
stopStrategy.BreakevenBarThresholdTicks = mParams.GetParameterValue(ParamBreakEvenBarThresholdTicks)
stopStrategy.breakEvenThresholdTicks = mParams.GetParameterValue(ParamBreakEvenThresholdTicks)
stopStrategy.confirmedCountForSwing = mParams.GetParameterValue(ParamConfirmedBarsForSwing)
stopStrategy.confirmedCountForTrailing = mParams.GetParameterValue(ParamConfirmedBarsForTrail)
stopStrategy.confirmedCountForExit = mParams.GetParameterValue(ParamConfirmedBarsForExit)
stopStrategy.RetracementStopPercent = mParams.GetParameterValue(ParamRetracementStopPercent)
stopStrategy.RetracementStopThresholdTicks = mParams.GetParameterValue(ParamRetracementStopThresholdTicks)
stopStrategy.breakevenPrice = breakevenPrice
stopStrategy.BackstopMABarLength = mBackstopMABarLength
stopStrategy.BackstopMAOffset = mBackstopMAOffset
stopStrategy.BackstopMAPeriods = mBackstopMAPeriods

Dim scaleStrategy As ScaleStrategy1
Set scaleStrategy = mScalingStrategyFactory.CreateStrategy(mContext)
scaleStrategy.MaxIncrements = mMaxIncrements
scaleStrategy.RiskIncrementPercent = mRiskIncrementPercent
scaleStrategy.ScaleThresholdFactor = mScaleThresholdFactor
scaleStrategy.RiskUnitPercent = mRiskUnitPercent
scaleStrategy.StopIncrementFactor = mStopIncrementFactor
scaleStrategy.timePeriod = mParams.GetParameterValue(ParamBarLength)
scaleStrategy.TimePeriodUnit = mParams.GetParameterValue(ParamTimePeriodUnit)
scaleStrategy.InitialStopFactor = mInitialStopFactor

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
                    targetTriggerPrice:=0)
mBracketOrderSell.cancelPrice = cancelPrice
mContext.ApplyPositionManagementStrategy mBracketOrderSell, stopStrategy
mContext.ApplyPositionManagementStrategy mBracketOrderSell, scaleStrategy

If Not mBracketOrderSell Is Nothing Then
    mBracketOrderSell.CancelAfter = 300
End If
End Sub

Private Sub calcPricesForSell( _
                ByVal triggerPrice As Double, _
                ByRef stopPrice As Double, _
                ByRef breakevenPrice As Double, _
                ByRef cancelPrice As Double, _
                ByRef risk As Double)
Dim SwingHigh As Double
SwingHigh = IIf(mSwing.SwingHigh(1).value > mSwing.SwingHigh(0).value, mSwing.SwingHigh(1).value, mSwing.SwingHigh(0).value)
If mSwingToBreakeven <> 0 Then
    If (SwingHigh - triggerPrice) > 0.5 * mSwingToBreakeven Then Exit Sub
    breakevenPrice = SwingHigh - mSwingToBreakeven
End If

stopPrice = SwingHigh + mStopBreakoutThresholdTicks * mTickSize

If stopPrice > triggerPrice + mMaxInitialStop Then
    stopPrice = triggerPrice + mMaxInitialStop
End If

risk = (stopPrice - triggerPrice - mEntryLimitOffsetTicks * mTickSize) * mTickValue / mTickSize

If mSwing.SwingHigh(0).value = MinDouble Then
    cancelPrice = mSwing.SwingHigh(1).value
Else
    cancelPrice = mSwing.SwingHigh(0).value
End If

End Sub

Private Sub buy( _
    ByVal triggerPrice As Double, _
    ByVal reason As String)
Dim stopPrice As Double
Dim breakevenPrice As Double
Dim cancelPrice As Double
Dim risk As Double
Dim scaleStrategy As ScaleStrategy1

If mOrderContextBuySimulated.PendingSize <> 0 Then
    If Not stillOkToBuy(triggerPrice) Then
        mBracketOrderBuySimulated.Cancel
        If mOrderContextBuy.PendingSize <> 0 Then mBracketOrderBuy.Cancel
        Exit Sub
    End If
    
    If triggerPrice >= mBracketOrderBuySimulated.EntryOrder.triggerPrice Then Exit Sub
    
    calcPricesForBuy triggerPrice, stopPrice, breakevenPrice, cancelPrice, risk
    mBracketOrderBuySimulated.SetNewEntryTriggerPrice triggerPrice
    mBracketOrderBuySimulated.SetNewStopLossPrice stopPrice
    On Error Resume Next
    mBracketOrderBuySimulated.Update
    If Err.Number = 0 Then
        On Error GoTo 0
        If mOrderContextBuy.PendingSize = 0 Then Exit Sub
        mBracketOrderBuy.SetNewQuantity UnitSize(risk)
        mBracketOrderBuy.SetNewEntryTriggerPrice triggerPrice
        mBracketOrderBuy.SetNewEntryPrice triggerPrice - mEntryLimitOffsetTicks * mTickSize
        mBracketOrderBuy.SetNewStopLossPrice stopPrice
        On Error Resume Next
        mBracketOrderBuy.Update
        If Err.Number = 0 Then Exit Sub
    End If
    ' the existing order plex cannot be amended - so cancel it
    ' and create a new one
    On Error GoTo 0
    mBracketOrderBuySimulated.Cancel
    If mOrderContextBuy.PendingSize <> 0 Then mBracketOrderBuy.Cancel
End If

mContext.LogTradeReason "Potential buy: " & reason

If Not okToBuy(triggerPrice) Then Exit Sub

calcPricesForBuy triggerPrice, stopPrice, breakevenPrice, cancelPrice, risk
If stopPrice = 0 Then
    mContext.LogTradeReason "reject: trigger price too close to breakeven price"
    Exit Sub
End If

Dim stopStrategySimulated As StopStrategy5
Set stopStrategySimulated = mStopStrategyFactory.CreateStrategy(mContext)
stopStrategySimulated.timePeriod = mParams.GetParameterValue(ParamBarLength)
stopStrategySimulated.TimePeriodUnit = mParams.GetParameterValue(ParamTimePeriodUnit)
stopStrategySimulated.minimumSwing = mParams.GetParameterValue(ParamMinimumSwingTicks) * mTickSize
stopStrategySimulated.breakoutThresholdTicks = mStopBreakoutThresholdTicks
stopStrategySimulated.RetraceFromExtremes = IIf(UCase$(mParams.GetParameterValue(ParamRetraceFromExtremes)) = "Y", True, False)

Set mBracketOrderBuySimulated = mOrderContextBuySimulated.buy( _
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
                    targetTriggerPrice:=0)
mBracketOrderBuySimulated.cancelPrice = cancelPrice
mContext.ApplyPositionManagementStrategy mBracketOrderBuySimulated, stopStrategySimulated

If mBracketOrderBuySimulated Is Nothing Then
    mNoTradeThisBar = True
        ' prevents each subsequent tick being processed if money management has prevented
        ' further trades
    Exit Sub
End If

If mOrderContextBuy.ActiveSize <> 0 Or mOrderContextBuy.PendingSize <> 0 Then Exit Sub

If UnitSize(risk) = 0 Then
    mContext.LogTradeReason "reject: allowable initial risk exceeded"
    Exit Sub
End If

Dim stopStrategy As StopStrategy5
Set stopStrategy = mStopStrategyFactory.CreateStrategy(mContext)
stopStrategy.timePeriod = mParams.GetParameterValue(ParamBarLength)
stopStrategy.TimePeriodUnit = mParams.GetParameterValue(ParamTimePeriodUnit)
stopStrategy.minimumSwing = mParams.GetParameterValue(ParamMinimumSwingTicks) * mTickSize
stopStrategy.breakoutThresholdTicks = mStopBreakoutThresholdTicks
stopStrategy.confirmedCountForBreakEven = mParams.GetParameterValue(ParamConfirmedBarsForBreakEven)
stopStrategy.RetraceFromExtremes = IIf(UCase$(mParams.GetParameterValue(ParamRetraceFromExtremes)) = "Y", True, False)
stopStrategy.BreakevenBarLength = mParams.GetParameterValue(ParamBreakEvenBarLength)
stopStrategy.BreakevenBarThresholdTicks = mParams.GetParameterValue(ParamBreakEvenBarThresholdTicks)
stopStrategy.breakEvenThresholdTicks = mParams.GetParameterValue(ParamBreakEvenThresholdTicks)
stopStrategy.confirmedCountForSwing = mParams.GetParameterValue(ParamConfirmedBarsForSwing)
stopStrategy.confirmedCountForTrailing = mParams.GetParameterValue(ParamConfirmedBarsForTrail)
stopStrategy.confirmedCountForExit = mParams.GetParameterValue(ParamConfirmedBarsForExit)
stopStrategy.RetracementStopPercent = mParams.GetParameterValue(ParamRetracementStopPercent)
stopStrategy.RetracementStopThresholdTicks = mParams.GetParameterValue(ParamRetracementStopThresholdTicks)
stopStrategy.breakevenPrice = breakevenPrice
stopStrategy.BackstopMABarLength = mBackstopMABarLength
stopStrategy.BackstopMAOffset = mBackstopMAOffset
stopStrategy.BackstopMAPeriods = mBackstopMAPeriods

Set scaleStrategy = mScalingStrategyFactory.CreateStrategy(mContext)
scaleStrategy.MaxIncrements = mMaxIncrements
scaleStrategy.RiskIncrementPercent = mRiskIncrementPercent
scaleStrategy.ScaleThresholdFactor = mScaleThresholdFactor
scaleStrategy.RiskUnitPercent = mRiskUnitPercent
scaleStrategy.StopIncrementFactor = mStopIncrementFactor
scaleStrategy.timePeriod = mParams.GetParameterValue(ParamBarLength)
scaleStrategy.TimePeriodUnit = mParams.GetParameterValue(ParamTimePeriodUnit)
scaleStrategy.InitialStopFactor = mInitialStopFactor

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
                    targetTriggerPrice:=0)
mBracketOrderBuy.cancelPrice = cancelPrice
mContext.ApplyPositionManagementStrategy mBracketOrderBuy, stopStrategy
mContext.ApplyPositionManagementStrategy mBracketOrderBuy, scaleStrategy

If Not mBracketOrderBuy Is Nothing Then
    mBracketOrderBuy.CancelAfter = 300
End If
End Sub

Private Sub calcPricesForBuy( _
                ByVal triggerPrice As Double, _
                ByRef stopPrice As Double, _
                ByRef breakevenPrice As Double, _
                ByRef cancelPrice As Double, _
                ByRef risk As Double)
Dim SwingLow As Double
SwingLow = IIf(mSwing.SwingLow(1).value < mSwing.SwingLow(0).value, mSwing.SwingLow(1).value, mSwing.SwingLow(0).value)
If mSwingToBreakeven <> 0 Then
    If (triggerPrice - SwingLow) > 0.5 * mSwingToBreakeven Then Exit Sub
    breakevenPrice = SwingLow + mSwingToBreakeven
End If

stopPrice = SwingLow - mStopBreakoutThresholdTicks * mTickSize

If stopPrice < triggerPrice - mMaxInitialStop Then
    stopPrice = triggerPrice - mMaxInitialStop
End If

risk = (triggerPrice + mEntryLimitOffsetTicks * mTickSize - stopPrice) * mTickValue / mTickSize

If mSwing.SwingLow(0).value = MaxDouble Then
    cancelPrice = mSwing.SwingLow(1).value
Else
    cancelPrice = mSwing.SwingLow(0).value
End If

End Sub

Private Sub processParameters(ByVal pParams As Parameters)
Set mParams = pParams
mEntryBreakoutThresholdTicks = mParams.GetParameterValue(ParamEntryBreakoutThresholdTicks)
mStopBreakoutThresholdTicks = mParams.GetParameterValue(ParamStopBreakoutThresholdTicks)
mEntryLimitOffsetTicks = mParams.GetParameterValue(ParamEntryLimitOffsetTicks)
mATRPeriods = mParams.GetParameterValue(ParamATRPeriods)
mMaxIncrements = mParams.GetParameterValue(ParamMaxIncrements)
mRiskUnitPercent = mParams.GetParameterValue(ParamRiskUnitPercent)
mRiskIncrementPercent = mParams.GetParameterValue(ParamRiskIncrementPercent)
mScaleThresholdFactor = mParams.GetParameterValue(ParamScaleThresholdFactor)
mStopIncrementFactor = mParams.GetParameterValue(ParamStopIncrementFactor)
mInitialStopFactor = mParams.GetParameterValue(ParamInitialStopFactor)
mMaxTradeSize = mParams.GetParameterValue(ParamMaxTradeSize)
mBackstopMABarLength = mParams.GetParameterValue(ParamBackstopMABarLength)
mBackstopMAPeriods = mParams.GetParameterValue(ParamBackstopMAPeriods)
mBackstopMAOffset = mParams.GetParameterValue(ParamBackstopMAOffset)
End Sub

Private Sub mBracketOrderBuySimulated_Completed()
checkTradingOpportunity
End Sub

Private Sub mBracketOrderBuySimulated_SelfCancelled()
checkTradingOpportunity
End Sub

Private Sub mBracketOrderSellSimulated_Completed()
checkTradingOpportunity
End Sub

Private Sub mBracketOrderSellSimulated_SelfCancelled()
checkTradingOpportunity
End Sub

Private Sub mPositionManagerSimulated_PositionFlat()
checkTradingOpportunity
End Sub

Private Function UnitSize(ByVal risk As Double) As Long
UnitSize = Int((mRiskUnitPercent * mOrderContextBuy.GetAccountBalance / 100) / risk)
If UnitSize > mMaxTradeSize Then UnitSize = mMaxTradeSize
'If UnitSize = 0 Then UnitSize = 1
End Function





