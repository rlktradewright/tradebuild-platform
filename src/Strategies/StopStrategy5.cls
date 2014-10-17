VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StopStrategy5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IPositionMgmtStrategy

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "StopStrategy5"

Private Const DummyLowCurr                          As Double = 922337203685477.5807@
Private Const DummyHighCurr                         As Double = -922337203685477.5807@

'@================================================================================
' Member variables
'@================================================================================

Private mContext As TradingContext
Private WithEvents mBracketOrder As BracketOrder
Attribute mBracketOrder.VB_VarHelpID = -1
Private mStopOrder As IOrder

Private mEntryOrderFilled As Boolean
Private mEntryPrice As Double
Private mInitialStopPrice As Double
Private mInitialReferencePrice As Double
Private mReferencePrice As Double
Private mBreakevenPrice As Double
Private mReferenceBar As Long
Private mStopChanged As Boolean

Private mInitialSwingStartPrice As Double

Private S                                           As StopStrategy5Static

Private mBars As Bars
Private mLongBars As Bars
Private mStopBars As Bars
Private mBreakevenBars As Bars

Private mLongPosition As Boolean

Private mTickSize As Double

Private mBarNumber As Long

Private mHighest As Double
Private mLowest As Double
Private mTrailStop As Boolean
Private mLongTrailStop As Boolean
Private mSwingCount As Long
Private mRestartCount As Long

Private mBreakevenEntryBar As Long

Private mTargetPrice As Double

Private mNextSwingStopCandidate As Double
Private mLastActivatedSwingStopCandidate As Double

Private mReceivingTicks As Boolean


'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "StopStrategy5 terminated"
End Sub

'================================================================================
' IPositionMgmtStrategy Interface Members
'================================================================================

Private Property Get IPositionMgmtStrategy_Name() As String
IPositionMgmtStrategy_Name = "StopStrategy5"
End Property

Private Sub IPositionMgmtStrategy_NotifyTick(ByVal pType As TickTypes)
Const ProcName As String = "IPositionMgmtStrategy_NotifyTick"
On Error GoTo Err

If pType <> TickTypes.TickTypeTrade Then Exit Sub

Dim price As Double
price = mContext.TradeQuote.price

Dim isNewBar  As Boolean
isNewBar = newbar(mBars.Bar(0).BarNumber)   ' need to do this first to maintain the
                                            ' bar number

Assert mEntryOrderFilled, "Tick not expected - entry order not yet filled"

If Not S.UseIntermediateStops Then
    If mReferencePrice <> mInitialReferencePrice Then
        If mLongPosition Then
            If price < mReferencePrice Then
                ' can only happen when the first contra-swing low is broken
                ' start over again
                mReferencePrice = mInitialReferencePrice
                mSwingCount = 0
                mRestartCount = mRestartCount + 1
                mBracketOrder.StopReason = "IS/" & mRestartCount
                mHighest = price
            End If
        Else
            If price > mReferencePrice Then
                ' can only happen when the first contra-swing high is broken
                ' start over again
                mReferencePrice = mInitialReferencePrice
                mSwingCount = 0
                mRestartCount = mRestartCount + 1
                mBracketOrder.StopReason = "IS/" & mRestartCount
                mLowest = price
            End If
        End If
    End If
End If

'If mBracketOrder.State = BracketOrderStateClosed Then
'    mContext.StopTickData Me
'    Set mBracketOrder = Nothing
'    Exit Sub
'End If

' NB: can't maintain current stop in a static
' variable, since stops may be altered by other
' code outside this strategy
Dim currentStop As Double
currentStop = mStopOrder.triggerPrice

Dim newStop As Double
newStop = currentStop

If Not S.LongBoll Is Nothing Then
    If mLongPosition Then
        If price >= S.LongBoll.Top.Value Then
            setTrailingStop price, "LB/" & mRestartCount
            Set S.LongBoll = Nothing ' prevent coming in here next time
            mTrailStop = True
            Exit Sub
        End If
'        If price >= mLongBoll.centre Then
'            If mEntryPrice > newStop Then
'                newStop = mEntryPrice
'                mBracketOrder.stopReason = "LBE/" & mRestartCount
'            End If
'        End If
    Else
        If price <= S.LongBoll.Bottom.Value Then
            setTrailingStop price, "LB/" & mRestartCount
            Set S.LongBoll = Nothing ' prevent coming in here next time
            mTrailStop = True
            Exit Sub
        End If
'        If price <= mLongBoll.centre Then
'            If mEntryPrice < newStop Then
'                newStop = mEntryPrice
'                mBracketOrder.stopReason = "LBE/" & mRestartCount
'            End If
'        End If
    End If
End If

If S.RewardToRiskRatio <> 0 Then
    If mLongPosition Then
        If price >= mTargetPrice Then
            setTrailingStop price, "RR/" & mRestartCount
            S.RewardToRiskRatio = 0  ' prevent coming in here next time
            mTrailStop = True
            Exit Sub
        End If
    Else
        If price <= mTargetPrice Then
            setTrailingStop price, "RR/" & mRestartCount
            S.RewardToRiskRatio = 0  ' prevent coming in here next time
            mTrailStop = True
            Exit Sub
        End If
    End If
End If

If S.SwingToMoveStopTicks > 0 And mNextSwingStopCandidate <> 0 Then
    If mLongPosition Then
        If price < mNextSwingStopCandidate Then
            calcNextSwingStopCandidate price
        ElseIf (price - mNextSwingStopCandidate) >= S.SwingToMoveStopTicks * mTickSize Then
            If (mNextSwingStopCandidate - S.BreakoutThresholdTicks * mTickSize) > newStop Then
                newStop = mNextSwingStopCandidate - S.BreakoutThresholdTicks * mTickSize
                mBracketOrder.StopReason = "SM/" & mRestartCount
            End If
            mLastActivatedSwingStopCandidate = mNextSwingStopCandidate
            calcNextSwingStopCandidate price
        End If
    Else
        If price > mNextSwingStopCandidate Then
            calcNextSwingStopCandidate price
        ElseIf (mNextSwingStopCandidate - price) >= S.BreakoutThresholdTicks * mTickSize Then
            If (mNextSwingStopCandidate + S.BreakoutThresholdTicks * mTickSize) < newStop Then
                newStop = mNextSwingStopCandidate + S.BreakoutThresholdTicks * mTickSize
                mBracketOrder.StopReason = "SM/" & mRestartCount
            End If
            mLastActivatedSwingStopCandidate = mNextSwingStopCandidate
            calcNextSwingStopCandidate price
        End If
    End If
End If
    
If mLongPosition Then
    If mBreakevenPrice > 0 And price >= mBreakevenPrice Then
        If mEntryPrice > newStop Then
            newStop = mEntryPrice
            mBracketOrder.StopReason = "SB/" & mRestartCount
        End If
    End If
Else
    If mBreakevenPrice > 0 And price <= mBreakevenPrice Then
        If mEntryPrice < newStop Then
            newStop = mEntryPrice
            mBracketOrder.StopReason = "SB/" & mRestartCount
        End If
    End If
End If


If mLongBars Is Nothing Or S.LongMA Is Nothing Then
ElseIf mLongPosition Then
    If mLongTrailStop Then
        If mLongBars.LowValue(1) > S.LongMA.ma(0).Value Then
            ' now that we've had one long bar whose low is above the long
            ' MA, we'll stop trailing and keep the stop below the bar
            ' before that one.
            newStop = mLongBars.LowValue(2) - S.BreakoutThresholdTicks * mTickSize
            mBracketOrder.StopReason = "LMA/" & mRestartCount
            mLongTrailStop = False
        End If
    Else
        If Abs(price - S.LongMA.ma(0).Value) <= 2 * mTickSize And _
            S.LongMA.ma(0).Value < S.LongMA.ma(2).Value _
        Then
            ' we're close to the long ma, which is falling, so start
            ' trailing the long bars
            On Error Resume Next
            newStop = mLongBars.LowerLowValue(2) - S.BreakoutThresholdTicks * mTickSize
            If Err.Number <> 0 Then
                Err.Clear
                newStop = mLongBars.LowerLowValue(1) - S.BreakoutThresholdTicks * mTickSize
            End If
            If Err.Number = 0 Then
                On Error GoTo 0
                mBracketOrder.StopReason = "LMA/" & mRestartCount
                mLongTrailStop = True
            Else
                On Error GoTo 0
            End If
        End If
    End If
Else
    If mLongTrailStop Then
        If mLongBars.HighValue(1) < S.LongMA.ma(0).Value Then
            ' now that we've had one long bar whose high is below the long
            ' MA, we'll stop trailing and keep the stop above the bar
            ' before that one.
            newStop = mLongBars.HighValue(2) + S.BreakoutThresholdTicks * mTickSize
            mBracketOrder.StopReason = "LMA/" & mRestartCount
            mLongTrailStop = False
        End If
    Else
        If Abs(price - S.LongMA.ma(0).Value) <= 2 * mTickSize And _
            S.LongMA.ma(0).Value > S.LongMA.ma(2).Value _
        Then
            ' we're close to the long ma, which is rising, so start
            ' trailing the long bars
            On Error Resume Next
            newStop = mLongBars.HigherHighValue(2) + S.BreakoutThresholdTicks * mTickSize
            If Err.Number <> 0 Then
                Err.Clear
                newStop = mLongBars.HigherHighValue(1) + S.BreakoutThresholdTicks * mTickSize
            End If
            If Err.Number = 0 Then
                On Error GoTo 0
                mBracketOrder.StopReason = "LMA/" & mRestartCount
                mLongTrailStop = True
            Else
                On Error GoTo 0
            End If
        End If
    End If
End If

' for other stop mechanisms, we only need to check the stop at the start
' of each bar so exit if not a new bar
If Not isNewBar Then
    If newStop <> currentStop Then
        mBracketOrder.AdjustStop newStop, mBracketOrder.EntryOrder.QuantityFilled
        mStopChanged = True
    End If
    Exit Sub
End If

Dim backStop As Double
Static bollTouched As Boolean

If mLongPosition Then
    If Not S.StopMA Is Nothing Then
        If Not bollTouched Then
            If mStopBars.HighValue(1) >= S.Boll.Top(1).Value Then bollTouched = True
        End If
        If bollTouched Then
            backStop = (Int(S.StopMA.ma.Value / mTickSize) - 1) * mTickSize
            If backStop > newStop Then
                newStop = backStop
                mBracketOrder.StopReason = "BS/" & mRestartCount
            End If
        End If
    End If
Else
    If Not S.StopMA Is Nothing Then
        If Not bollTouched Then
            If mStopBars.LowValue(1) <= S.Boll.Bottom(1).Value Then bollTouched = True
        End If
        If bollTouched Then
            backStop = (Int((S.StopMA.ma.Value / mTickSize) + 0.99999) + 1) * mTickSize
            If backStop < newStop Then
                newStop = backStop
                mBracketOrder.StopReason = "BS/" & mRestartCount
            End If
        End If
    End If
End If

Dim retracementStop As Double

If Not mTrailStop Then
    If mLongPosition Then
    
        If S.BreakoutThresholdTicks > 0 And mNextSwingStopCandidate = 0 Then
            calcNextSwingStopCandidate price
        End If
        
        If S.TicksSwingToTrail > 0 And mBars.HighValue(1) > (mReferencePrice + S.TicksSwingToTrail * mTickSize) Then
            mBracketOrder.StopReason = "ST/" & mRestartCount
            mTrailStop = True
        End If
        
        If Not mBreakevenBars Is Nothing Then
            If mBreakevenBars.Bar.BarNumber > mBreakevenEntryBar + 1 Then
                If mBreakevenBars.LowValue(1) >= mEntryPrice + S.BreakevenBarThresholdTicks * mTickSize Then
                    If mEntryPrice > newStop Then
                        newStop = mEntryPrice
                        mBracketOrder.StopReason = "BB/" & mRestartCount
                    End If
                End If
            End If
        End If
            
        If S.ConfirmedBarsForBreakEven <> 0 Then
            If S.MACDStudy.strength(1).Value = MACDConfirmedStrength And _
                S.MACDStudy.strengthCount.Value = S.ConfirmedBarsForBreakEven And _
                price >= mEntryPrice + S.BreakEvenThresholdTicks * mTickSize _
            Then
                If mEntryPrice > newStop Then
                    newStop = mEntryPrice
                    mBracketOrder.StopReason = "BE/" & mRestartCount
                End If
            End If
        End If
        
        If S.ConfirmedBarsForSwing <> 0 Then
            If S.MACDStudy.strength(1).Value = MACDConfirmedStrength And _
                S.MACDStudy.strengthCount.Value = S.ConfirmedBarsForSwing _
            Then
                If S.Swing.SwingLowPoint(0).Value > newStop Then
                    newStop = S.Swing.SwingLowPoint(0).Value - S.BreakoutThresholdTicks * mTickSize
                    mBracketOrder.StopReason = "CS/" & mRestartCount
                End If
            End If
        End If
        
        If S.ConfirmedBarsForExit <> 0 Then
            If S.MACDStudy.strength(1).Value = MACDConfirmedWeakness And _
                S.MACDStudy.strengthCount.Value = S.ConfirmedBarsForExit _
            Then
                newStop = mBars.LowValue(1) - S.BreakoutThresholdTicks * mTickSize
                mBracketOrder.StopReason = "CE/" & mRestartCount
            End If
        End If
        
        If mSwingCount >= 1 And _
            S.ConfirmedBarsForTrailing <> 0 _
        Then
            If S.MACDStudy.strength(1).Value = MACDConfirmedStrength And _
                S.MACDStudy.strengthCount.Value = S.ConfirmedBarsForTrailing _
            Then
                mTrailStop = True
                mBracketOrder.StopReason = "CT/" & mRestartCount
            End If
        End If
        
        If mBars.HighValue(1) > mHighest Then
            mHighest = mBars.HighValue(1)
            If mHighest - mInitialStopPrice >= S.RetracementStopThresholdTicks * mTickSize And _
                S.RetracementStopFactor <> 0 _
            Then
                retracementStop = mInitialStopPrice + _
                                    (Int(S.RetracementStopFactor * (mHighest - mInitialStopPrice) / mTickSize) - 1) * mTickSize
                If retracementStop > newStop Then
                    newStop = retracementStop
                    mBracketOrder.StopReason = "RS/" & mRestartCount
                End If
            End If
        End If
        
        If S.Swing.SwingLowPoint(-1).BarNumber > mReferenceBar And _
            S.Swing.SwingLowPoint(-1).Value - mReferencePrice > S.MaxContraSwingFactor * _
                                                    (IIf(S.RetraceFromExtremes, mHighest, mBars.HighValue(1)) - mReferencePrice) _
        Then
            mReferencePrice = S.Swing.SwingLowPoint(-1).Value
            mReferenceBar = S.Swing.SwingLowPoint(-1).BarNumber
            mSwingCount = mSwingCount + 1
            If mSwingCount >= 2 Then
'                If mTrailStopSMABarLength = 0 Or _
'                    mTrailStopSMAPeriods = 0 _
'                Then
                    mTrailStop = True
                    mBracketOrder.StopReason = "TS/" & mRestartCount
'                ElseIf mReferencePrice <= S.StopMA.ma Then
'                    mTrailStop = True
'                    mBracketOrder.stopReason = "TM/" & mRestartCount
'                Else
'                    If mReferencePrice - S.BreakoutThresholdTicks * mTickSize > newStop Then
'                        newStop = mReferencePrice - S.BreakoutThresholdTicks * mTickSize
'                        mBracketOrder.stopReason = "S" & mSwingCount & "/" & mRestartCount
'                    End If
'                End If
            Else
                If S.UseIntermediateStops And _
                    (mReferencePrice - S.BreakoutThresholdTicks * mTickSize) > newStop Then
                    newStop = mReferencePrice - S.BreakoutThresholdTicks * mTickSize
                    mBracketOrder.StopReason = "S1/" & mRestartCount
                End If
            End If
        End If
    
    ElseIf Not mLongPosition Then
        
        If S.BreakoutThresholdTicks > 0 And mNextSwingStopCandidate = 0 Then
            calcNextSwingStopCandidate price
        End If
        
        If S.TicksSwingToTrail > 0 And mBars.LowValue(1) < (mReferencePrice - S.TicksSwingToTrail * mTickSize) Then
            mBracketOrder.StopReason = "ST/" & mRestartCount
            mTrailStop = True
        End If
        
        If Not mBreakevenBars Is Nothing Then
            If mBreakevenBars.Bar.BarNumber > mBreakevenEntryBar + 1 Then
                If mBreakevenBars.HighValue(1) <= mEntryPrice - S.BreakevenBarThresholdTicks * mTickSize Then
                    If mEntryPrice < newStop Then
                        newStop = mEntryPrice
                        mBracketOrder.StopReason = "BB/" & mRestartCount
                    End If
                End If
            End If
        End If
            
        If S.ConfirmedBarsForBreakEven <> 0 Then
            If S.MACDStudy.strength(1).Value = MACDConfirmedWeakness And _
                S.MACDStudy.strengthCount.Value = S.ConfirmedBarsForBreakEven And _
                price <= mEntryPrice - S.BreakEvenThresholdTicks * mTickSize _
            Then
                If mEntryPrice < newStop Then
                    newStop = mEntryPrice
                    mBracketOrder.StopReason = "BE/" & mRestartCount
                End If
            End If
        End If
        
        If S.ConfirmedBarsForSwing <> 0 Then
            If S.MACDStudy.strength(1).Value = MACDConfirmedWeakness And _
                S.MACDStudy.strengthCount.Value = S.ConfirmedBarsForSwing _
            Then
                If S.Swing.SwingHighPoint(0).Value > newStop Then
                    newStop = S.Swing.SwingHighPoint(0).Value + S.BreakoutThresholdTicks * mTickSize
                    mBracketOrder.StopReason = "CS/" & mRestartCount
                End If
            End If
        End If
        
        If S.ConfirmedBarsForExit <> 0 Then
            If S.MACDStudy.strength(1).Value = MACDConfirmedStrength And _
                S.MACDStudy.strengthCount.Value = S.ConfirmedBarsForExit _
            Then
                newStop = mBars.HighValue(1) + S.BreakoutThresholdTicks * mTickSize
                mBracketOrder.StopReason = "CE/" & mRestartCount
            End If
        End If
        
        If mSwingCount >= 1 And _
            S.ConfirmedBarsForTrailing <> 0 _
        Then
            If S.MACDStudy.strength(1).Value = MACDConfirmedWeakness And _
                S.MACDStudy.strengthCount.Value = S.ConfirmedBarsForTrailing _
            Then
                mTrailStop = True
                mBracketOrder.StopReason = "CT/" & mRestartCount
            End If
        End If
        
        If mBars.LowValue(1) < mLowest Then
            mLowest = mBars.LowValue(1)
            If mInitialStopPrice - mLowest >= S.RetracementStopThresholdTicks * mTickSize And _
                S.RetracementStopFactor <> 0 _
            Then
                retracementStop = mInitialStopPrice - _
                                    (Int(S.RetracementStopFactor * (mInitialStopPrice - mLowest) / mTickSize) + 1) * mTickSize
                If retracementStop < newStop Then
                    newStop = retracementStop
                    mBracketOrder.StopReason = "RS/" & mRestartCount
                End If
            End If
        End If
        
        If S.Swing.SwingHighPoint(-1).BarNumber > mReferenceBar And _
            mReferencePrice - S.Swing.SwingHighPoint(-1).Value > S.MaxContraSwingFactor * _
                                                    (mReferencePrice - IIf(S.RetraceFromExtremes, mLowest, mBars.LowValue(1))) _
        Then
            mReferencePrice = S.Swing.SwingHighPoint(-1).Value
            mReferenceBar = S.Swing.SwingHighPoint(-1).BarNumber
            mSwingCount = mSwingCount + 1
            If mSwingCount = 2 Then
'                If mTrailStopSMABarLength = 0 Or _
'                    mTrailStopSMAPeriods = 0 _
'                Then
                    mTrailStop = True
                    mBracketOrder.StopReason = "TS/" & mRestartCount
'                ElseIf mReferencePrice >= S.StopMA.ma Then
'                    mTrailStop = True
'                    mBracketOrder.stopReason = "TM/" & mRestartCount
'                Else
'                    If mReferencePrice + S.BreakoutThresholdTicks * mTickSize < newStop Then
'                        newStop = mReferencePrice + S.BreakoutThresholdTicks * mTickSize
'                        mBracketOrder.stopReason = "S" & mSwingCount & "/" & mRestartCount
'                    End If
'                End If
            Else
                If S.UseIntermediateStops And _
                    (mReferencePrice + S.BreakoutThresholdTicks * mTickSize) < newStop Then
                    newStop = mReferencePrice + S.BreakoutThresholdTicks * mTickSize
                    mBracketOrder.StopReason = "S1/" & mRestartCount
                End If
            End If
        End If
    
    End If
End If

Dim lowStop As Double
Dim highStop As Double

If mTrailStop Then
    If mLongPosition Then
        If mBars.LowValue(1) > mBars.LowValue(2) Then
            lowStop = mBars.LowValue(2) - S.BreakoutThresholdTicks * mTickSize
            If lowStop > currentStop And lowStop > newStop Then
                newStop = lowStop
            End If
        End If
    Else
        If mBars.HighValue(1) < mBars.HighValue(2) Then
            highStop = mBars.HighValue(2) + S.BreakoutThresholdTicks * mTickSize
            If highStop < currentStop And highStop < newStop Then
                newStop = highStop
            End If
        End If
    End If
ElseIf mLongTrailStop Then
    If mLongPosition Then
        If mLongBars.LowValue(1) > mLongBars.LowValue(2) Then
            lowStop = mLongBars.LowValue(2) - S.BreakoutThresholdTicks * mTickSize
            If lowStop > currentStop And lowStop > newStop Then
                newStop = lowStop
                mBracketOrder.StopReason = "LMA/" & mRestartCount
            End If
        End If
    Else
        If mLongBars.HighValue(1) < mLongBars.HighValue(2) Then
            highStop = mLongBars.HighValue(2) + S.BreakoutThresholdTicks * mTickSize
            If highStop < currentStop And highStop < newStop Then
                newStop = highStop
                mBracketOrder.StopReason = "LMA/" & mRestartCount
            End If
        End If
    End If
End If

If newStop <> currentStop Then
    mBracketOrder.AdjustStop newStop, mBracketOrder.EntryOrder.QuantityFilled
    mStopChanged = True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IPositionMgmtStrategy_Start( _
                ByVal pContext As TradingContext, _
                ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IPositionMgmtStrategy_Start"
On Error GoTo Err

Set mContext = pContext

mTickSize = mContext.Contract.TickSize

Set mBars = mContext.GetBars(S.Timeframe)
If Not S.LongTimeframe Is Nothing Then Set mLongBars = mContext.GetBars(S.LongTimeframe)
If Not S.StopTimeframe Is Nothing Then Set mStopBars = mContext.GetBars(S.StopTimeframe)
If Not S.BreakevenTimeframe Is Nothing Then Set mBreakevenBars = mContext.GetBars(S.BreakevenTimeframe)


Set mBracketOrder = pBracketOrder
mLongPosition = IIf(mBracketOrder.EntryOrder.Action = OrderActionBuy, True, False)
Set mStopOrder = mBracketOrder.StopLossOrder
mInitialStopPrice = mStopOrder.triggerPrice

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' mBracketOrder Event Handlers
'================================================================================

Private Sub mBracketOrder_Completed()
Const ProcName As String = "mBracketOrder_Completed"
On Error GoTo Err

If mReceivingTicks Then mContext.StopTickData Me
mReceivingTicks = False
Set mBracketOrder = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mBracketOrder_EntryOrderPartiallyFilled()
Const ProcName As String = "mBracketOrder_EntryOrderPartiallyFilled"
On Error GoTo Err

Dim swingLow1 As Double
Dim swingLow2 As Double
Dim swingHigh1 As Double
Dim swingHigh2 As Double

mEntryOrderFilled = True
mBracketOrder.StopReason = "IS/0"
mEntryPrice = mBracketOrder.EntryOrder.AveragePrice

If S.RewardToRiskRatio <> 0 Then
    If mLongPosition Then
        mTargetPrice = mEntryPrice + S.RewardToRiskRatio * (mEntryPrice - mInitialStopPrice)
    Else
        mTargetPrice = mEntryPrice - S.RewardToRiskRatio * (mInitialStopPrice - mEntryPrice)
    End If
End If

If Not mReceivingTicks Then
    mReceivingTicks = True
    mContext.StartTickData Me
    mHighest = mBars.CloseValue
    mLowest = mHighest
End If

If Not mBreakevenBars Is Nothing And mBreakevenEntryBar <> 0 Then
    mBreakevenEntryBar = mBreakevenBars.Bar.BarNumber
End If

If mReferenceBar <> 0 Then Exit Sub ' means reference point was set via the ReferencePoint
                                    ' property or this is not the first partial fill
                                    

mReferenceBar = mBars.Bar.BarNumber

If mTrailStop Then Exit Sub

If mLongPosition Then
    swingLow1 = IIf(IsEmpty(S.Swing.SwingLow(1)), DummyLowCurr, S.Swing.SwingLowPoint(-1).Value)
    swingLow2 = IIf(IsEmpty(S.Swing.SwingLow(2)), DummyLowCurr, S.Swing.SwingLowPoint(-2).Value)
    If S.Swing.SwingLowPoint(0).Value = MaxDouble Then
        mNextSwingStopCandidate = swingLow1
        If Not S.UseIntermediateStops Then
            mReferencePrice = swingLow1
'        ElseIf swingLow2 < swingLow1 Then
'            mReferencePrice = swingLow2
        Else
            mReferencePrice = swingLow1
        End If
    Else
        mInitialSwingStartPrice = S.Swing.SwingLowPoint(0).Value
        If Not S.UseIntermediateStops Then
            mReferencePrice = S.Swing.SwingLowPoint(0).Value
'        ElseIf swingLow1 < S.Swing.SwingLowPoint(0).Value Then
''            mReferencePrice = swingLow2
''                        ' this should really be swingLow1 but testing seems to give
''                        ' better results with swingLow2
'            If swingLow2 < swingLow1 Then
'                mReferencePrice = swingLow2
'            Else
'                mReferencePrice = swingLow1
'            End If
        Else
            mReferencePrice = S.Swing.SwingLowPoint(0).Value
        End If
    End If
Else
    swingHigh1 = IIf(IsEmpty(S.Swing.SwingHigh(1)), DummyHighCurr, S.Swing.SwingHighPoint(-1).Value)
    swingHigh2 = IIf(IsEmpty(S.Swing.SwingHigh(2)), DummyHighCurr, S.Swing.SwingHighPoint(-2).Value)
    If S.Swing.SwingHighPoint(0).Value = MinDouble Then
        mInitialSwingStartPrice = swingHigh1
        If Not S.UseIntermediateStops Then
            mReferencePrice = swingHigh1
'        ElseIf swingHigh2 > swingHigh1 Then
'            mReferencePrice = swingHigh2
        Else
            mReferencePrice = swingHigh1
        End If
    Else
        mInitialSwingStartPrice = S.Swing.SwingHighPoint(0).Value
        If Not S.UseIntermediateStops Then
            mReferencePrice = S.Swing.SwingHighPoint(0).Value
'        ElseIf swingHigh1 > S.Swing.SwingHighPoint(0).Value Then
''            mReferencePrice = swingHigh2
''                        ' this should really be swingHigh1 but testing seems to give
''                        ' better results with swingHigh2
'            If swingHigh2 > swingHigh1 Then
'                mReferencePrice = swingHigh2
'            Else
'                mReferencePrice = swingHigh1
'            End If
        Else
            mReferencePrice = S.Swing.SwingHighPoint(0).Value
        End If
    End If
End If
mInitialReferencePrice = mReferencePrice
If mInitialReferencePrice = 0 Then Stop

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let noContraSwing(ByVal val As Boolean)
If val Then mTrailStop = True
End Property

Public Property Let ReferencePoint(ByRef Value As SwingPoint)
Const ProcName As String = "ReferencePoint"
On Error GoTo Err

mReferencePrice = Value.Value
mReferenceBar = Value.BarNumber
mInitialReferencePrice = Value.Value
mLowest = mReferencePrice
mHighest = mReferencePrice
mSwingCount = 0
mRestartCount = 0

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Initialise(ByVal pStatic As StopStrategy5Static)
Set S = pStatic
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcNextSwingStopCandidate(ByVal price As Double)
Const ProcName As String = "calcNextSwingStopCandidate"
On Error GoTo Err

Static prevSwingStopCandidate As Double

Dim i As Long

If mNextSwingStopCandidate <> 0 Then
    prevSwingStopCandidate = mNextSwingStopCandidate
End If

If mLongPosition Then
    i = -1
    If S.Swing.SwingLowPoint(0).Value = MaxDouble Then
        i = -2
        mNextSwingStopCandidate = S.Swing.SwingLowPoint(-1).Value
    Else
        mNextSwingStopCandidate = S.Swing.SwingLowPoint(0).Value
    End If
    Do While (price - S.Swing.SwingLowPoint(i).Value) < S.BreakoutThresholdTicks * mTickSize And _
            S.Swing.SwingLowPoint(i).Value < mNextSwingStopCandidate And _
            S.Swing.SwingLowPoint(i).Value > mLastActivatedSwingStopCandidate
        mNextSwingStopCandidate = S.Swing.SwingLowPoint(i).Value
        i = i - 1
    Loop
Else
    i = -1
    If S.Swing.SwingHighPoint(0).Value = MinDouble Then
        i = -2
        mNextSwingStopCandidate = S.Swing.SwingHighPoint(-1).Value
    Else
        mNextSwingStopCandidate = S.Swing.SwingHighPoint(0).Value
    End If
    Do While (S.Swing.SwingHighPoint(i).Value - price) < S.BreakoutThresholdTicks * mTickSize And _
            S.Swing.SwingHighPoint(i).Value > mNextSwingStopCandidate And _
            S.Swing.SwingHighPoint(i).Value < mLastActivatedSwingStopCandidate
        mNextSwingStopCandidate = S.Swing.SwingHighPoint(i).Value
        i = i - 1
    Loop
End If

If prevSwingStopCandidate = mNextSwingStopCandidate Then
    mNextSwingStopCandidate = 0
Else
    Debug.Print "Next swing stop candidate: " & mNextSwingStopCandidate
End If


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function newbar(ByVal BarNumber As Long) As Boolean
If BarNumber > mBarNumber Then
    mBarNumber = BarNumber
    newbar = True
Else
    newbar = False
End If
End Function

Private Sub setTrailingStop(ByVal price As Double, ByVal reason As String)
Const ProcName As String = "setTrailingStop"
On Error GoTo Err

Dim i As Long
Dim firstLow As Double
Dim firstHigh As Double
Dim newStop As Double
Dim currentStop As Double

If mLongPosition Then
    If mBars.LowValue(1) < mBars.LowValue(0) Then
        firstLow = mBars.LowValue(1)
        i = 2
    Else
        firstLow = mBars.LowValue(0)
        i = 1
    End If
    Do While mBars.LowValue(i) >= firstLow
        i = i + 1
    Loop
    newStop = mBars.LowValue(i)
    currentStop = mStopOrder.triggerPrice
    If newStop > currentStop Then
        mBracketOrder.AdjustStop newStop, mBracketOrder.EntryOrder.QuantityFilled
        mBracketOrder.StopReason = reason
        mStopChanged = True
    End If
Else
    If mBars.HighValue(1) > mBars.HighValue(0) Then
        firstHigh = mBars.HighValue(1)
        i = 2
    Else
        firstHigh = mBars.HighValue(0)
        i = 1
    End If
    Do While mBars.HighValue(i) <= firstHigh
        i = i + 1
    Loop
    newStop = mBars.HighValue(i)
    currentStop = mStopOrder.triggerPrice
    If newStop < currentStop Then
        mBracketOrder.AdjustStop newStop, mBracketOrder.EntryOrder.QuantityFilled
        mBracketOrder.StopReason = reason
        mStopChanged = True
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub








