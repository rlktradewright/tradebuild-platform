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

Implements IBracketOrderEventSink
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

'@================================================================================
' Member variables
'@================================================================================

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
' IBracketOrderEventSink Interface Members
'================================================================================

Private Sub IBracketOrderEventSink_NotifyBracketOrderCompletion(ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "IBracketOrderEventSink_NotifyBracketOrderCompletion"
On Error GoTo Err

If mReceivingTicks Then StopTickData Me
mReceivingTicks = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrderEventSink_NotifyBracketOrderFill(ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "IBracketOrderEventSink_NotifyBracketOrderFill"
On Error GoTo Err

mEntryOrderFilled = True
SetStopReason "IS/0"
mEntryPrice = AveragePrice

If S.RewardToRiskRatio <> 0 Then
    If mLongPosition Then
        mTargetPrice = mEntryPrice + S.RewardToRiskRatio * (mEntryPrice - mInitialStopPrice)
    Else
        mTargetPrice = mEntryPrice - S.RewardToRiskRatio * (mInitialStopPrice - mEntryPrice)
    End If
End If

If Not mReceivingTicks Then
    mReceivingTicks = True
    StartTickData Me
    mHighest = CloseValue(0)
    mLowest = mHighest
End If

If Not S.BreakevenTimeframe Is Nothing And mBreakevenEntryBar <> 0 Then
    mBreakevenEntryBar = BarNumber(0, S.BreakevenTimeframe)
End If

If mReferenceBar <> 0 Then Exit Sub ' means reference point was set via the ReferencePoint
                                    ' property or this is not the first partial fill
                                    

mReferenceBar = BarNumber(0)

If mTrailStop Then Exit Sub

Dim swingLow1 As Double
Dim swingLow2 As Double
Dim swingHigh1 As Double
Dim swingHigh2 As Double

If mLongPosition Then
    swingLow1 = IIf(IsEmpty(SwingLow()), MaxDouble, SwingLow())
    swingLow2 = IIf(IsEmpty(SwingLow(-1)), MaxDouble, SwingLow(-1))
    If PotentialSwingLow() = MaxDouble Then
        mNextSwingStopCandidate = swingLow1
        If Not S.UseIntermediateStops Then
            mReferencePrice = swingLow1
        ElseIf swingLow2 < swingLow1 Then
            mReferencePrice = swingLow2
        Else
            mReferencePrice = swingLow1
        End If
    Else
        mInitialSwingStartPrice = PotentialSwingLow()
        If Not S.UseIntermediateStops Then
            mReferencePrice = PotentialSwingLow()
        ElseIf swingLow1 < PotentialSwingLow() Then
'            mReferencePrice = swingLow2
'                        ' this should really be swingLow1 but testing seems to give
'                        ' better results with swingLow2
            If swingLow2 < swingLow1 Then
                mReferencePrice = swingLow2
            Else
                mReferencePrice = swingLow1
            End If
        Else
            mReferencePrice = PotentialSwingLow()
        End If
    End If
Else
    swingHigh1 = IIf(IsEmpty(SwingHigh()), MinDouble, SwingHigh())
    swingHigh2 = IIf(IsEmpty(SwingHigh(-1)), MinDouble, SwingHigh(-1))
    If PotentialSwingHigh() = MinDouble Then
        mInitialSwingStartPrice = swingHigh1
        If Not S.UseIntermediateStops Then
            mReferencePrice = swingHigh1
        ElseIf swingHigh2 > swingHigh1 Then
            mReferencePrice = swingHigh2
        Else
            mReferencePrice = swingHigh1
        End If
    Else
        mInitialSwingStartPrice = PotentialSwingHigh()
        If Not S.UseIntermediateStops Then
            mReferencePrice = PotentialSwingHigh()
        ElseIf swingHigh1 > PotentialSwingHigh() Then
'            mReferencePrice = swingHigh2
'                        ' this should really be swingHigh1 but testing seems to give
'                        ' better results with swingHigh2
            If swingHigh2 > swingHigh1 Then
                mReferencePrice = swingHigh2
            Else
                mReferencePrice = swingHigh1
            End If
        Else
            mReferencePrice = PotentialSwingHigh()
        End If
    End If
End If
mInitialReferencePrice = mReferencePrice
If mInitialReferencePrice = 0 Then Stop

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrderEventSink_NotifyBracketOrderStopLossAdjusted(ByVal BracketOrder As ResourceIdentifier)

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

If Not IsCurrentTickTrade Then Exit Sub

Dim price As Double
price = TradePrice

Dim isNewBar  As Boolean
isNewBar = newbar   ' need to do this first to maintain the bar number

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
                SetStopReason "IS/" & mRestartCount
                mHighest = price
            End If
        Else
            If price > mReferencePrice Then
                ' can only happen when the first contra-swing high is broken
                ' start over again
                mReferencePrice = mInitialReferencePrice
                mSwingCount = 0
                mRestartCount = mRestartCount + 1
                SetStopReason "IS/" & mRestartCount
                mLowest = price
            End If
        End If
    End If
End If

If IsCompleted Then
    StopTickData Me
    ClearPrimaryBracketOrder
    Exit Sub
End If

' NB: can't maintain current stop in a static
' variable, since stops may be altered by other
' code outside this strategy
Dim currentStop As Double
currentStop = StopLossTriggerPrice

Dim newStop As Double
newStop = currentStop

If Not S.LongBoll Is Nothing Then
    If mLongPosition Then
        If price >= BollingerTop(, S.LongBoll) Then
            setTrailingStop price, "LB/" & mRestartCount
            Set S.LongBoll = Nothing ' prevent coming in here next time
            mTrailStop = True
            Exit Sub
        End If
'        If price >= mLongBoll.centre Then
'            If mEntryPrice > newStop Then
'                newStop = mEntryPrice
'                SetStopReason  "LBE/" & mRestartCount
'            End If
'        End If
    Else
        If price <= BollingerBottom(, S.LongBoll) Then
            setTrailingStop price, "LB/" & mRestartCount
            Set S.LongBoll = Nothing ' prevent coming in here next time
            mTrailStop = True
            Exit Sub
        End If
'        If price <= mLongBoll.centre Then
'            If mEntryPrice < newStop Then
'                newStop = mEntryPrice
'                SetStopReason  "LBE/" & mRestartCount
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
                SetStopReason "SM/" & mRestartCount
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
                SetStopReason "SM/" & mRestartCount
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
            SetStopReason "SB/" & mRestartCount
        End If
    End If
Else
    If mBreakevenPrice > 0 And price <= mBreakevenPrice Then
        If mEntryPrice < newStop Then
            newStop = mEntryPrice
            SetStopReason "SB/" & mRestartCount
        End If
    End If
End If


If S.LongTimeframe Is Nothing Or S.LongMA Is Nothing Then
ElseIf mLongPosition Then
    If mLongTrailStop Then
        If LowValue(-1, S.LongTimeframe) > SMA(0, S.LongMA) Then
            ' now that we've had one long bar whose low is above the long
            ' MA, we'll stop trailing and keep the stop below the bar
            ' before that one.
            newStop = LowValue(-2, S.LongTimeframe) - S.BreakoutThresholdTicks * mTickSize
            SetStopReason "LMA/" & mRestartCount
            mLongTrailStop = False
        End If
    Else
        If Abs(price - SMA(0, S.LongMA)) <= 2 * mTickSize And _
            SMA(0, S.LongMA) < SMA(-2, S.LongMA) _
        Then
            ' we're close to the long ma, which is falling, so start
            ' trailing the long bars
            On Error Resume Next
            newStop = LowerLowValue(2, , S.LongTimeframe) - S.BreakoutThresholdTicks * mTickSize
            If Err.Number <> 0 Then
                Err.Clear
                newStop = LowerLowValue(1, , S.LongTimeframe) - S.BreakoutThresholdTicks * mTickSize
            End If
            If Err.Number = 0 Then
                On Error GoTo 0
                SetStopReason "LMA/" & mRestartCount
                mLongTrailStop = True
            Else
                On Error GoTo 0
            End If
        End If
    End If
Else
    If mLongTrailStop Then
        If HighValue(-1, S.LongTimeframe) < SMA(0, S.LongMA) Then
            ' now that we've had one long bar whose high is below the long
            ' MA, we'll stop trailing and keep the stop above the bar
            ' before that one.
            newStop = HighValue(-2, S.LongTimeframe) + S.BreakoutThresholdTicks * mTickSize
            SetStopReason "LMA/" & mRestartCount
            mLongTrailStop = False
        End If
    Else
        If Abs(price - SMA(0, S.LongMA)) <= 2 * mTickSize And _
            SMA(0, S.LongMA) > SMA(-2, S.LongMA) _
        Then
            ' we're close to the long ma, which is rising, so start
            ' trailing the long bars
            On Error Resume Next
            newStop = HigherHighValue(2, , S.LongTimeframe) + S.BreakoutThresholdTicks * mTickSize
            If Err.Number <> 0 Then
                Err.Clear
                newStop = HigherHighValue(1, , S.LongTimeframe) + S.BreakoutThresholdTicks * mTickSize
            End If
            If Err.Number = 0 Then
                On Error GoTo 0
                SetStopReason "LMA/" & mRestartCount
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
        If CanTrade Then AdjustStop newStop, QuantityFilled
        mStopChanged = True
    End If
    Exit Sub
End If

Dim backStop As Double
Static bollTouched As Boolean

If mLongPosition Then
    If Not S.StopMA Is Nothing Then
        If Not bollTouched Then
            If HighValue(-1, S.StopTimeframe) >= BollingerTop(-1, S.Boll) Then bollTouched = True
        End If
        If bollTouched Then
            backStop = (Int(SMA(0, S.StopMA) / mTickSize) - 1) * mTickSize
            If backStop > newStop Then
                newStop = backStop
                SetStopReason "BS/" & mRestartCount
            End If
        End If
    End If
Else
    If Not S.StopMA Is Nothing Then
        If Not bollTouched Then
            If LowValue(-1, S.StopTimeframe) <= BollingerBottom(-1, S.Boll) Then bollTouched = True
        End If
        If bollTouched Then
            backStop = (Int((SMA(, S.StopMA) / mTickSize) + 0.99999) + 1) * mTickSize
            If backStop < newStop Then
                newStop = backStop
                SetStopReason "BS/" & mRestartCount
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
        
        If S.TicksSwingToTrail > 0 And HighValue(-1) > (mReferencePrice + S.TicksSwingToTrail * mTickSize) Then
            SetStopReason "ST/" & mRestartCount
            mTrailStop = True
        End If
        
        If Not S.BreakevenTimeframe Is Nothing Then
            If BarNumber(0, S.BreakevenTimeframe) > mBreakevenEntryBar + 1 Then
                If LowValue(-1, S.BreakevenTimeframe) >= mEntryPrice + S.BreakevenBarThresholdTicks * mTickSize Then
                    If mEntryPrice > newStop Then
                        newStop = mEntryPrice
                        SetStopReason "BB/" & mRestartCount
                    End If
                End If
            End If
        End If
            
        If S.ConfirmedBarsForBreakEven <> 0 Then
            If MACDStrengthConfirmed(-1) And _
                MACDStrengthCount(0) = S.ConfirmedBarsForBreakEven And _
                price >= mEntryPrice + S.BreakEvenThresholdTicks * mTickSize _
            Then
                If mEntryPrice > newStop Then
                    newStop = mEntryPrice
                    SetStopReason "BE/" & mRestartCount
                End If
            End If
        End If
        
        If S.ConfirmedBarsForSwing <> 0 Then
            If MACDStrengthConfirmed(-1) And _
                MACDStrengthCount(0) = S.ConfirmedBarsForSwing _
            Then
                If SwingLow(0) > newStop Then
                    newStop = SwingLow(0) - S.BreakoutThresholdTicks * mTickSize
                    SetStopReason "CS/" & mRestartCount
                End If
            End If
        End If
        
        If S.ConfirmedBarsForExit <> 0 Then
            If MACDWeaknessConfirmed(-1) And _
                MACDStrengthCount(0) = S.ConfirmedBarsForExit _
            Then
                newStop = LowValue(-1) - S.BreakoutThresholdTicks * mTickSize
                SetStopReason "CE/" & mRestartCount
            End If
        End If
        
        If mSwingCount >= 1 And _
            S.ConfirmedBarsForTrailing <> 0 _
        Then
            If MACDStrengthConfirmed(-1) And _
                MACDStrengthCount(0) = S.ConfirmedBarsForTrailing _
            Then
                mTrailStop = True
                SetStopReason "CT/" & mRestartCount
            End If
        End If
        
        If HighValue(-1) > mHighest Then
            mHighest = HighValue(-1)
            If mHighest - mInitialStopPrice >= S.RetracementStopThresholdTicks * mTickSize And _
                S.RetracementStopFactor <> 0 _
            Then
                retracementStop = mInitialStopPrice + _
                                    (Int(S.RetracementStopFactor * (mHighest - mInitialStopPrice) / mTickSize) - 1) * mTickSize
                If retracementStop > newStop Then
                    newStop = retracementStop
                    SetStopReason "RS/" & mRestartCount
                End If
            End If
        End If
        
        If SwingLowBarNumber > mReferenceBar And _
            SwingLow - mReferencePrice > S.MaxContraSwingFactor * _
                                                    (IIf(S.RetraceFromExtremes, mHighest, HighValue(-1)) - mReferencePrice) _
        Then
            mReferencePrice = SwingLow
            mReferenceBar = SwingLowBarNumber
            mSwingCount = mSwingCount + 1
            If mSwingCount >= 2 Then
'                If mTrailStopSMABarLength = 0 Or _
'                    mTrailStopSMAPeriods = 0 _
'                Then
                    mTrailStop = True
                    SetStopReason "TS/" & mRestartCount
'                ElseIf mReferencePrice <= S.StopMA.ma Then
'                    mTrailStop = True
'                    SetStopReason  "TM/" & mRestartCount
'                Else
'                    If mReferencePrice - S.BreakoutThresholdTicks * mTickSize > newStop Then
'                        newStop = mReferencePrice - S.BreakoutThresholdTicks * mTickSize
'                        SetStopReason  "S" & mSwingCount & "/" & mRestartCount
'                    End If
'                End If
            Else
                If S.UseIntermediateStops And _
                    (mReferencePrice - S.BreakoutThresholdTicks * mTickSize) > newStop Then
                    newStop = mReferencePrice - S.BreakoutThresholdTicks * mTickSize
                    SetStopReason "S1/" & mRestartCount
                End If
            End If
        End If
    
    ElseIf Not mLongPosition Then
        
        If S.BreakoutThresholdTicks > 0 And mNextSwingStopCandidate = 0 Then
            calcNextSwingStopCandidate price
        End If
        
        If S.TicksSwingToTrail > 0 And LowValue(-1) < (mReferencePrice - S.TicksSwingToTrail * mTickSize) Then
            SetStopReason "ST/" & mRestartCount
            mTrailStop = True
        End If
        
        If Not S.BreakevenTimeframe Is Nothing Then
            If BarNumber(0, S.BreakevenTimeframe) > mBreakevenEntryBar + 1 Then
                If HighValue(-1, S.BreakevenTimeframe) <= mEntryPrice - S.BreakevenBarThresholdTicks * mTickSize Then
                    If mEntryPrice < newStop Then
                        newStop = mEntryPrice
                        SetStopReason "BB/" & mRestartCount
                    End If
                End If
            End If
        End If
            
        If S.ConfirmedBarsForBreakEven <> 0 Then
            If MACDWeaknessConfirmed(-1) And _
                MACDStrengthCount(0) = S.ConfirmedBarsForBreakEven And _
                price <= mEntryPrice - S.BreakEvenThresholdTicks * mTickSize _
            Then
                If mEntryPrice < newStop Then
                    newStop = mEntryPrice
                    SetStopReason "BE/" & mRestartCount
                End If
            End If
        End If
        
        If S.ConfirmedBarsForSwing <> 0 Then
            If MACDWeaknessConfirmed(-1) And _
                MACDStrengthCount(0) = S.ConfirmedBarsForSwing _
            Then
                If SwingHigh(0) > newStop Then
                    newStop = SwingHigh(0) + S.BreakoutThresholdTicks * mTickSize
                    SetStopReason "CS/" & mRestartCount
                End If
            End If
        End If
        
        If S.ConfirmedBarsForExit <> 0 Then
            If MACDStrengthConfirmed(-1) And _
                MACDStrengthCount(0) = S.ConfirmedBarsForExit _
            Then
                newStop = HighValue(-1) + S.BreakoutThresholdTicks * mTickSize
                SetStopReason "CE/" & mRestartCount
            End If
        End If
        
        If mSwingCount >= 1 And _
            S.ConfirmedBarsForTrailing <> 0 _
        Then
            If MACDWeaknessConfirmed(-1) And _
                MACDStrengthCount(0) = S.ConfirmedBarsForTrailing _
            Then
                mTrailStop = True
                SetStopReason "CT/" & mRestartCount
            End If
        End If
        
        If LowValue(-1) < mLowest Then
            mLowest = LowValue(-1)
            If mInitialStopPrice - mLowest >= S.RetracementStopThresholdTicks * mTickSize And _
                S.RetracementStopFactor <> 0 _
            Then
                retracementStop = mInitialStopPrice - _
                                    (Int(S.RetracementStopFactor * (mInitialStopPrice - mLowest) / mTickSize) + 1) * mTickSize
                If retracementStop < newStop Then
                    newStop = retracementStop
                    SetStopReason "RS/" & mRestartCount
                End If
            End If
        End If
        
        If SwingHighBarNumber > mReferenceBar And _
            mReferencePrice - SwingHigh > S.MaxContraSwingFactor * _
                                                    (mReferencePrice - IIf(S.RetraceFromExtremes, mLowest, LowValue(-1))) _
        Then
            mReferencePrice = SwingHigh
            mReferenceBar = SwingHighBarNumber
            mSwingCount = mSwingCount + 1
            If mSwingCount = 2 Then
'                If mTrailStopSMABarLength = 0 Or _
'                    mTrailStopSMAPeriods = 0 _
'                Then
                    mTrailStop = True
                    SetStopReason "TS/" & mRestartCount
'                ElseIf mReferencePrice >= S.StopMA.ma Then
'                    mTrailStop = True
'                    SetStopReason  "TM/" & mRestartCount
'                Else
'                    If mReferencePrice + S.BreakoutThresholdTicks * mTickSize < newStop Then
'                        newStop = mReferencePrice + S.BreakoutThresholdTicks * mTickSize
'                        SetStopReason  "S" & mSwingCount & "/" & mRestartCount
'                    End If
'                End If
            Else
                If S.UseIntermediateStops And _
                    (mReferencePrice + S.BreakoutThresholdTicks * mTickSize) < newStop Then
                    newStop = mReferencePrice + S.BreakoutThresholdTicks * mTickSize
                    SetStopReason "S1/" & mRestartCount
                End If
            End If
        End If
    
    End If
End If

Dim lowStop As Double
Dim highStop As Double

If mTrailStop Then
    If mLongPosition Then
        If LowValue(-1) > LowValue(-2) Then
            lowStop = LowValue(-2) - S.BreakoutThresholdTicks * mTickSize
            If lowStop > currentStop And lowStop > newStop Then
                newStop = lowStop
            End If
        End If
    Else
        If HighValue(-1) < HighValue(-2) Then
            highStop = HighValue(-2) + S.BreakoutThresholdTicks * mTickSize
            If highStop < currentStop And highStop < newStop Then
                newStop = highStop
            End If
        End If
    End If
ElseIf mLongTrailStop Then
    If mLongPosition Then
        If LowValue(-1, S.LongTimeframe) > LowValue(-2, S.LongTimeframe) Then
            lowStop = LowValue(-2, S.LongTimeframe) - S.BreakoutThresholdTicks * mTickSize
            If lowStop > currentStop And lowStop > newStop Then
                newStop = lowStop
                SetStopReason "LMA/" & mRestartCount
            End If
        End If
    Else
        If HighValue(-1, S.LongTimeframe) < HighValue(-2, S.LongTimeframe) Then
            highStop = HighValue(-2, S.LongTimeframe) + S.BreakoutThresholdTicks * mTickSize
            If highStop < currentStop And highStop < newStop Then
                newStop = highStop
                SetStopReason "LMA/" & mRestartCount
            End If
        End If
    End If
End If

If newStop <> currentStop Then
    If CanTrade Then AdjustStop newStop, QuantityFilled
    mStopChanged = True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IPositionMgmtStrategy_Start( _
                ByVal TradingContext As TradingContext, _
                ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "IPositionMgmtStrategy_Start"
On Error GoTo Err

mTickSize = TickSize

SetPrimaryBracketOrder BracketOrder
mLongPosition = IsBuy
mInitialStopPrice = StopLossTriggerPrice

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
    If PotentialSwingLow() = MaxDouble Then
        i = -1
        mNextSwingStopCandidate = SwingLow(0)
    Else
        i = 0
        mNextSwingStopCandidate = PotentialSwingLow()
    End If
    Do While (price - SwingLow(i)) < S.BreakoutThresholdTicks * mTickSize And _
            SwingLow(i) < mNextSwingStopCandidate And _
            SwingLow(i) > mLastActivatedSwingStopCandidate
        mNextSwingStopCandidate = SwingLow(i)
        i = i - 1
    Loop
Else
    If PotentialSwingHigh() = MinDouble Then
        i = -1
        mNextSwingStopCandidate = SwingHigh(0)
    Else
        i = 0
        mNextSwingStopCandidate = PotentialSwingHigh()
    End If
    Do While (SwingHigh(i) - price) < S.BreakoutThresholdTicks * mTickSize And _
            SwingHigh(i) > mNextSwingStopCandidate And _
            SwingHigh(i) < mLastActivatedSwingStopCandidate
        mNextSwingStopCandidate = SwingHigh(i)
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

Private Function newbar() As Boolean
If CurrentBarNumber > mBarNumber Then
    mBarNumber = CurrentBarNumber
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
    If LowValue(-1) < LowValue(0) Then
        firstLow = LowValue(-1)
        i = -2
    Else
        firstLow = LowValue(0)
        i = -1
    End If
    Do While LowValue(i) >= firstLow
        i = i - 1
    Loop
    newStop = LowValue(i)
    currentStop = StopLossTriggerPrice
    If newStop > currentStop Then
        If CanTrade Then AdjustStop newStop, QuantityFilled
        SetStopReason reason
        mStopChanged = True
    End If
Else
    If HighValue(-1) > HighValue(0) Then
        firstHigh = HighValue(-1)
        i = -2
    Else
        firstHigh = HighValue(0)
        i = -1
    End If
    Do While HighValue(i) <= firstHigh
        i = i - 1
    Loop
    newStop = HighValue(i)
    currentStop = StopLossTriggerPrice
    If newStop < currentStop Then
        If CanTrade Then AdjustStop newStop, QuantityFilled
        SetStopReason reason
        mStopChanged = True
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub








