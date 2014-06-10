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
Attribute VB_Creatable = False
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

Event StoppedOut(ByVal longStop As Boolean, ByVal atInitialStop As Boolean)
Event StopChanged(ByVal newStop As Double)

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

Private mBars As Bars
Private mLongBars As Bars
Private mSwing As Swing
Private mMACDStudy As MACD
Private mListenerKey As String

Private mStopBars As Bars
Private mStopMA As EMA
Private mBoll As BollingerBands

Private mLongBoll As BollingerBands
Private mLongBollPeriods As Long

Private mBreakevenBars As Bars

Private mTimePeriod As Integer
Private mTimePeriodUnit As String
Private mLongTimePeriod As Integer
Private mLongTimePeriodUnit As String

Private mLongMA As SMA
Private mLongMAPeriods As Long

Private mUseIntermediateStops As Boolean

Private mMinimumSwing As Double

Private mLongPosition As Boolean

Private mBreakoutThreshold As Integer
Private mTickSize As Double

Private mBarNumber As Long

Private mHighest As Double
Private mLowest As Double
Private mTrailStop As Boolean
Private mLongTrailStop As Boolean
Private mSwingCount As Long
Private mRestartCount As Long
Private mRetraceFromExtremes As Boolean

Private mMaxContraSwingFactor As Single
Private mConfirmedCountForBreakEven As Long
Private mConfirmedCountForSwing As Long
Private mConfirmedCountForTrailing As Long
Private mConfirmedCountForExit As Long

Private mBackstopMABarLength As Integer
Private mBackstopMAPeriods As Integer
Private mBackstopMAOffset As Integer

Private mBreakevenBarLength As Integer
Private mBreakevenBarThresholdTicks As Integer
Private mBreakevenEntryBar As Long

Private mBreakEvenThresholdTicks As Long

Private mRetracementStopFactor As Single
Private mRetracementStopThresholdTicks As Long

Private mTicksSwingToTrail As Long

Private mRewardToRiskRatio As Single
Private mTargetPrice As Double

Private mTicksSwingToMoveStop As Long
Private mNextSwingStopCandidate As Double
Private mLastActivatedSwingStopCandidate As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mUseIntermediateStops = True
mMaxContraSwingFactor = 0.5!
mRetraceFromExtremes = True
mLongMAPeriods = 34
End Sub

Private Sub Class_Terminate()
Debug.Print "StopStrategy5 terminated"
End Sub

'================================================================================
' IListener Interface Members
'================================================================================

Public Sub IListener_notify(ByVal listenerKey As String, _
                            ByVal newval As Variant, _
                            ByVal valuetype As String, _
                            ByVal barNumber As Long, _
                            ByVal Timestamp As Date)

Dim price As Double
Dim currentStop As Double
Dim newStop As Double
Dim lowStop As Double
Dim highStop As Double
Dim isNewBar  As Boolean
Dim retracementStop As Double
Static bollTouched As Boolean
Dim backStop As Double

price = newval

isNewBar = newbar(barNumber)    ' need to do this first to maintain the
                                ' bar number

Assert mEntryOrderFilled, "Tick not expected - entry order not yet filled"

If Not mUseIntermediateStops Then
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

If mBracketOrder.closedOut Then
    ' can now remove ourself from the listeners list - should get no more
    ' notifications
    mBars.removeListener mListenerKey
    Set mBracketOrder = Nothing
    Exit Sub
End If

currentStop = mStopOrder.triggerPrice
                            ' NB: can't maintain current stop in a static
                            ' variable, since stops may be altered by other
                            ' code outside this strategy
newStop = currentStop

If mLongBollPeriods <> 0 Then
    If mLongPosition Then
        If price >= mLongBoll.Top Then
            setTrailingStop price, "LB/" & mRestartCount
            mLongBollPeriods = 0    ' prevent coming in here next time
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
        If price <= mLongBoll.Bottom Then
            setTrailingStop price, "LB/" & mRestartCount
            mLongBollPeriods = 0    ' prevent coming in here next time
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

If mRewardToRiskRatio <> 0 Then
    If mLongPosition Then
        If price >= mTargetPrice Then
            setTrailingStop price, "RR/" & mRestartCount
            mRewardToRiskRatio = 0  ' prevent coming in here next time
            mTrailStop = True
            Exit Sub
        End If
    Else
        If price <= mTargetPrice Then
            setTrailingStop price, "RR/" & mRestartCount
            mRewardToRiskRatio = 0  ' prevent coming in here next time
            mTrailStop = True
            Exit Sub
        End If
    End If
End If

If mTicksSwingToMoveStop > 0 And mNextSwingStopCandidate <> 0 Then
    If mLongPosition Then
        If price < mNextSwingStopCandidate Then
            calcNextSwingStopCandidate price
        ElseIf (price - mNextSwingStopCandidate) >= mTicksSwingToMoveStop * mTickSize Then
            If (mNextSwingStopCandidate - mBreakoutThreshold * mTickSize) > newStop Then
                newStop = mNextSwingStopCandidate - mBreakoutThreshold * mTickSize
                mBracketOrder.StopReason = "SM/" & mRestartCount
            End If
            mLastActivatedSwingStopCandidate = mNextSwingStopCandidate
            calcNextSwingStopCandidate price
        End If
    Else
        If price > mNextSwingStopCandidate Then
            calcNextSwingStopCandidate price
        ElseIf (mNextSwingStopCandidate - price) >= mTicksSwingToMoveStop * mTickSize Then
            If (mNextSwingStopCandidate + mBreakoutThreshold * mTickSize) < newStop Then
                newStop = mNextSwingStopCandidate + mBreakoutThreshold * mTickSize
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

If mLongPosition And _
    mLongTimePeriod <> 0 And _
    mLongMAPeriods <> 0 _
    Then
    If mLongTrailStop Then
        If mLongBars.lowPrice(1) > mLongMA.ma(0) Then
            ' now that we've had one long bar whose low is above the long
            ' MA, we'll stop trailing and keep the stop below the bar
            ' before that one.
            newStop = mLongBars.lowPrice(2) - mBreakoutThreshold * mTickSize
            mBracketOrder.StopReason = "LMA/" & mRestartCount
            mLongTrailStop = False
        End If
    Else
        If Abs(price - mLongMA.ma(0)) <= 2 * mTickSize And _
            mLongMA.ma(0) < mLongMA.ma(2) _
        Then
            ' we're close to the long ma, which is falling, so start
            ' trailing the long bars
            On Error Resume Next
            newStop = mLongBars.lowerLowPrice(2) - mBreakoutThreshold * mTickSize
            If Err.Number <> 0 Then
                Err.Clear
                newStop = mLongBars.lowerLowPrice(1) - mBreakoutThreshold * mTickSize
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
ElseIf (Not mLongPosition) And _
        mLongTimePeriod <> 0 And _
        mLongMAPeriods <> 0 _
Then
    If mLongTrailStop Then
        If mLongBars.highPrice(1) < mLongMA.ma(0) Then
            ' now that we've had one long bar whose high is below the long
            ' MA, we'll stop trailing and keep the stop above the bar
            ' before that one.
            newStop = mLongBars.highPrice(2) + mBreakoutThreshold * mTickSize
            mBracketOrder.StopReason = "LMA/" & mRestartCount
            mLongTrailStop = False
        End If
    Else
        If Abs(price - mLongMA.ma(0)) <= 2 * mTickSize And _
            mLongMA.ma(0) > mLongMA.ma(2) _
        Then
            ' we're close to the long ma, which is rising, so start
            ' trailing the long bars
            On Error Resume Next
            newStop = mLongBars.higherHighPrice(2) + mBreakoutThreshold * mTickSize
            If Err.Number <> 0 Then
                Err.Clear
                newStop = mLongBars.higherHighPrice(1) + mBreakoutThreshold * mTickSize
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
        RaiseEvent StopChanged(newStop)
    End If
    Exit Sub
End If

If mLongPosition Then
    If Not mStopMA Is Nothing Then
        If Not bollTouched Then
            If mStopBars.highPrice(1) >= mBoll.Top(1) Then bollTouched = True
        End If
        If bollTouched Then
            backStop = (Int(mStopMA.ma / mTickSize) - 1) * mTickSize
            If backStop > newStop Then
                newStop = backStop
                mBracketOrder.StopReason = "BS/" & mRestartCount
            End If
        End If
    End If
Else
    If Not mStopMA Is Nothing Then
        If Not bollTouched Then
            If mStopBars.lowPrice(1) <= mBoll.Bottom(1) Then bollTouched = True
        End If
        If bollTouched Then
            backStop = (Int((mStopMA.ma / mTickSize) + 0.99999) + 1) * mTickSize
            If backStop < newStop Then
                newStop = backStop
                mBracketOrder.StopReason = "BS/" & mRestartCount
            End If
        End If
    End If
End If

If Not mTrailStop Then
    If mLongPosition Then
    
        If mTicksSwingToMoveStop > 0 And mNextSwingStopCandidate = 0 Then
            calcNextSwingStopCandidate price
        End If
        
        If mTicksSwingToTrail > 0 And mBars.highPrice(1) > (mReferencePrice + mTicksSwingToTrail * mTickSize) Then
            mBracketOrder.StopReason = "ST/" & mRestartCount
            mTrailStop = True
        End If
        
        If mBreakevenBarLength > 0 Then
            If mBreakevenBars.barNumber > mBreakevenEntryBar + 1 Then
                If mBreakevenBars.lowPrice(1) >= mEntryPrice + mBreakevenBarThresholdTicks * mTickSize Then
                    If mEntryPrice > newStop Then
                        newStop = mEntryPrice
                        mBracketOrder.StopReason = "BB/" & mRestartCount
                    End If
                End If
            End If
        End If
            
        If mConfirmedCountForBreakEven <> 0 Then
            If mMACDStudy.strength(1) = MACDConfirmedStrength And _
                mMACDStudy.strengthCount = mConfirmedCountForBreakEven And _
                price >= mEntryPrice + mBreakEvenThresholdTicks * mTickSize _
            Then
                If mEntryPrice > newStop Then
                    newStop = mEntryPrice
                    mBracketOrder.StopReason = "BE/" & mRestartCount
                End If
            End If
        End If
        
        If mConfirmedCountForSwing <> 0 Then
            If mMACDStudy.strength(1) = MACDConfirmedStrength And _
                mMACDStudy.strengthCount = mConfirmedCountForSwing _
            Then
                If mSwing.SwingLow(0) > newStop Then
                    newStop = mSwing.SwingLow(0) - mBreakoutThreshold * mTickSize
                    mBracketOrder.StopReason = "CS/" & mRestartCount
                End If
            End If
        End If
        
        If mConfirmedCountForExit <> 0 Then
            If mMACDStudy.strength(1) = MACDConfirmedWeakness And _
                mMACDStudy.strengthCount = mConfirmedCountForExit _
            Then
                newStop = mBars.lowPrice(1) - mBreakoutThreshold * mTickSize
                mBracketOrder.StopReason = "CE/" & mRestartCount
            End If
        End If
        
        If mSwingCount >= 1 And _
            mConfirmedCountForTrailing <> 0 _
        Then
            If mMACDStudy.strength(1) = MACDConfirmedStrength And _
                mMACDStudy.strengthCount = mConfirmedCountForTrailing _
            Then
                mTrailStop = True
                mBracketOrder.StopReason = "CT/" & mRestartCount
            End If
        End If
        
        If mBars.highPrice(1) > mHighest Then
            mHighest = mBars.highPrice(1)
            If mHighest - mInitialStopPrice >= mRetracementStopThresholdTicks * mTickSize And _
                mRetracementStopFactor <> 0 _
            Then
                retracementStop = mInitialStopPrice + _
                                    (Int(mRetracementStopFactor * (mHighest - mInitialStopPrice) / mTickSize) - 1) * mTickSize
                If retracementStop > newStop Then
                    newStop = retracementStop
                    mBracketOrder.StopReason = "RS/" & mRestartCount
                End If
            End If
        End If
        
        If mSwing.swingLowBarNumber(1) > mReferenceBar And _
            mSwing.SwingLow(1) - mReferencePrice > mMaxContraSwingFactor * _
                                                    (IIf(mRetraceFromExtremes, mHighest, mBars.highPrice(1)) - mReferencePrice) _
        Then
            mReferencePrice = mSwing.SwingLow(1)
            mReferenceBar = mSwing.swingLowBarNumber(1)
            mSwingCount = mSwingCount + 1
            If mSwingCount >= 2 Then
'                If mTrailStopSMABarLength = 0 Or _
'                    mTrailStopSMAPeriods = 0 _
'                Then
                    mTrailStop = True
                    mBracketOrder.StopReason = "TS/" & mRestartCount
'                ElseIf mReferencePrice <= mStopMA.ma Then
'                    mTrailStop = True
'                    mBracketOrder.stopReason = "TM/" & mRestartCount
'                Else
'                    If mReferencePrice - mBreakoutThreshold * mTickSize > newStop Then
'                        newStop = mReferencePrice - mBreakoutThreshold * mTickSize
'                        mBracketOrder.stopReason = "S" & mSwingCount & "/" & mRestartCount
'                    End If
'                End If
            Else
                If mUseIntermediateStops And _
                    (mReferencePrice - mBreakoutThreshold * mTickSize) > newStop Then
                    newStop = mReferencePrice - mBreakoutThreshold * mTickSize
                    mBracketOrder.StopReason = "S1/" & mRestartCount
                End If
            End If
        End If
    
    ElseIf Not mLongPosition Then
        
        If mTicksSwingToMoveStop > 0 And mNextSwingStopCandidate = 0 Then
            calcNextSwingStopCandidate price
        End If
        
        If mTicksSwingToTrail > 0 And mBars.lowPrice(1) < (mReferencePrice - mTicksSwingToTrail * mTickSize) Then
            mBracketOrder.StopReason = "ST/" & mRestartCount
            mTrailStop = True
        End If
        
        If mBreakevenBarLength > 0 Then
            If mBreakevenBars.barNumber > mBreakevenEntryBar + 1 Then
                If mBreakevenBars.highPrice(1) <= mEntryPrice - mBreakevenBarThresholdTicks * mTickSize Then
                    If mEntryPrice < newStop Then
                        newStop = mEntryPrice
                        mBracketOrder.StopReason = "BB/" & mRestartCount
                    End If
                End If
            End If
        End If
            
        If mConfirmedCountForBreakEven <> 0 Then
            If mMACDStudy.strength(1) = MACDConfirmedWeakness And _
                mMACDStudy.strengthCount = mConfirmedCountForBreakEven And _
                price <= mEntryPrice - mBreakEvenThresholdTicks * mTickSize _
            Then
                If mEntryPrice < newStop Then
                    newStop = mEntryPrice
                    mBracketOrder.StopReason = "BE/" & mRestartCount
                End If
            End If
        End If
        
        If mConfirmedCountForSwing <> 0 Then
            If mMACDStudy.strength(1) = MACDConfirmedWeakness And _
                mMACDStudy.strengthCount = mConfirmedCountForSwing _
            Then
                If mSwing.SwingHigh(0) > newStop Then
                    newStop = mSwing.SwingHigh(0) + mBreakoutThreshold * mTickSize
                    mBracketOrder.StopReason = "CS/" & mRestartCount
                End If
            End If
        End If
        
        If mConfirmedCountForExit <> 0 Then
            If mMACDStudy.strength(1) = MACDConfirmedStrength And _
                mMACDStudy.strengthCount = mConfirmedCountForExit _
            Then
                newStop = mBars.highPrice(1) + mBreakoutThreshold * mTickSize
                mBracketOrder.StopReason = "CE/" & mRestartCount
            End If
        End If
        
        If mSwingCount >= 1 And _
            mConfirmedCountForTrailing <> 0 _
        Then
            If mMACDStudy.strength(1) = MACDConfirmedWeakness And _
                mMACDStudy.strengthCount = mConfirmedCountForTrailing _
            Then
                mTrailStop = True
                mBracketOrder.StopReason = "CT/" & mRestartCount
            End If
        End If
        
        If mBars.lowPrice(1) < mLowest Then
            mLowest = mBars.lowPrice(1)
            If mInitialStopPrice - mLowest >= mRetracementStopThresholdTicks * mTickSize And _
                mRetracementStopFactor <> 0 _
            Then
                    retracementStop = mInitialStopPrice - _
                                    (Int(mRetracementStopFactor * (mInitialStopPrice - mLowest) / mTickSize) + 1) * mTickSize
                If retracementStop < newStop Then
                    newStop = retracementStop
                    mBracketOrder.StopReason = "RS/" & mRestartCount
                End If
            End If
        End If
        
        If mSwing.swingHighBarNumber(1) > mReferenceBar And _
            mReferencePrice - mSwing.SwingHigh(1) > mMaxContraSwingFactor * _
                                                    (mReferencePrice - IIf(mRetraceFromExtremes, mLowest, mBars.lowPrice(1))) _
        Then
            mReferencePrice = mSwing.SwingHigh(1)
            mReferenceBar = mSwing.swingHighBarNumber(1)
            mSwingCount = mSwingCount + 1
            If mSwingCount = 2 Then
'                If mTrailStopSMABarLength = 0 Or _
'                    mTrailStopSMAPeriods = 0 _
'                Then
                    mTrailStop = True
                    mBracketOrder.StopReason = "TS/" & mRestartCount
'                ElseIf mReferencePrice >= mStopMA.ma Then
'                    mTrailStop = True
'                    mBracketOrder.stopReason = "TM/" & mRestartCount
'                Else
'                    If mReferencePrice + mBreakoutThreshold * mTickSize < newStop Then
'                        newStop = mReferencePrice + mBreakoutThreshold * mTickSize
'                        mBracketOrder.stopReason = "S" & mSwingCount & "/" & mRestartCount
'                    End If
'                End If
            Else
                If mUseIntermediateStops And _
                    (mReferencePrice + mBreakoutThreshold * mTickSize) < newStop Then
                    newStop = mReferencePrice + mBreakoutThreshold * mTickSize
                    mBracketOrder.StopReason = "S1/" & mRestartCount
                End If
            End If
        End If
    
    End If
End If

If mTrailStop Then
    If mLongPosition Then
        If mBars.lowPrice(1) > mBars.lowPrice(2) Then
            lowStop = mBars.lowPrice(2) - mBreakoutThreshold * mTickSize
            If lowStop > currentStop And lowStop > newStop Then
                newStop = lowStop
            End If
        End If
    Else
        If mBars.highPrice(1) < mBars.highPrice(2) Then
            highStop = mBars.highPrice(2) + mBreakoutThreshold * mTickSize
            If highStop < currentStop And highStop < newStop Then
                newStop = highStop
            End If
        End If
    End If
ElseIf mLongTrailStop Then
    If mLongPosition Then
        If mLongBars.lowPrice(1) > mLongBars.lowPrice(2) Then
            lowStop = mLongBars.lowPrice(2) - mBreakoutThreshold * mTickSize
            If lowStop > currentStop And lowStop > newStop Then
                newStop = lowStop
                mBracketOrder.StopReason = "LMA/" & mRestartCount
            End If
        End If
    Else
        If mLongBars.highPrice(1) < mLongBars.highPrice(2) Then
            highStop = mLongBars.highPrice(2) + mBreakoutThreshold * mTickSize
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
    RaiseEvent StopChanged(newStop)
End If
End Sub

'================================================================================
' IPositionMgmtStrategy Interface Members
'================================================================================

Private Sub IPositionMgmtStrategy_NotifyTick(ByVal pType As TickTypes)
Const ProcName As String = "IPositionMgmtStrategy_NotifyTick"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IPositionMgmtStrategy_Start(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IPositionMgmtStrategy_Start"
On Error GoTo Err

Set mBracketOrder = pBracketOrder
mLongPosition = IIf(mBracketOrder.EntryOrder.Action = ActionBuy, True, False)
Set mStopOrder = mBracketOrder.stoporder
mInitialStopPrice = mStopOrder.triggerPrice

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' mBracketOrder Event Handlers
'================================================================================

Private Sub mBracketOrder_Completed()
If mListenerKey <> "" Then mBars.removeListener mListenerKey
Set mBracketOrder = Nothing
End Sub

Private Sub mBracketOrder_EntryOrderPartiallyFilled()
Dim swingLow1 As Double
Dim swingLow2 As Double
Dim swingHigh1 As Double
Dim swingHigh2 As Double

mEntryOrderFilled = True
mBracketOrder.StopReason = "IS/0"
mEntryPrice = mBracketOrder.EntryOrder.AveragePrice

If mRewardToRiskRatio <> 0 Then
    If mLongPosition Then
        mTargetPrice = mEntryPrice + mRewardToRiskRatio * (mEntryPrice - mInitialStopPrice)
    Else
        mTargetPrice = mEntryPrice - mRewardToRiskRatio * (mInitialStopPrice - mEntryPrice)
    End If
End If

If mListenerKey = "" Then
    mListenerKey = mBars.addListener(Me, "lastprice")
    mHighest = mBars.closePrice
    mLowest = mHighest
End If

If mBreakevenBarLength > 0 And _
    mBreakevenEntryBar <> 0 _
Then
    mBreakevenEntryBar = mBreakevenBars.barNumber
End If

If mReferenceBar <> 0 Then Exit Sub ' means reference point was set via the ReferencePoint
                                    ' property or this is not the first partial fill
                                    

mReferenceBar = mBars.barNumber

If mTrailStop Then Exit Sub

If mLongPosition Then
    swingLow1 = IIf(IsEmpty(mSwing.SwingLow(1)), DummyLowCurr, mSwing.SwingLow(1))
    swingLow2 = IIf(IsEmpty(mSwing.SwingLow(2)), DummyLowCurr, mSwing.SwingLow(2))
    If mSwing.SwingLow(0) = DummyLow Then
        mNextSwingStopCandidate = swingLow1
        If Not mUseIntermediateStops Then
            mReferencePrice = swingLow1
'        ElseIf swingLow2 < swingLow1 Then
'            mReferencePrice = swingLow2
        Else
            mReferencePrice = swingLow1
        End If
    Else
        mInitialSwingStartPrice = mSwing.SwingLow(0)
        If Not mUseIntermediateStops Then
            mReferencePrice = mSwing.SwingLow(0)
'        ElseIf swingLow1 < mSwing.swingLow(0) Then
''            mReferencePrice = swingLow2
''                        ' this should really be swingLow1 but testing seems to give
''                        ' better results with swingLow2
'            If swingLow2 < swingLow1 Then
'                mReferencePrice = swingLow2
'            Else
'                mReferencePrice = swingLow1
'            End If
        Else
            mReferencePrice = mSwing.SwingLow(0)
        End If
    End If
Else
    swingHigh1 = IIf(IsEmpty(mSwing.SwingHigh(1)), DummyHighCurr, mSwing.SwingHigh(1))
    swingHigh2 = IIf(IsEmpty(mSwing.SwingHigh(2)), DummyHighCurr, mSwing.SwingHigh(2))
    If mSwing.SwingHigh(0) = DummyHigh Then
        mInitialSwingStartPrice = swingHigh1
        If Not mUseIntermediateStops Then
            mReferencePrice = swingHigh1
'        ElseIf swingHigh2 > swingHigh1 Then
'            mReferencePrice = swingHigh2
        Else
            mReferencePrice = swingHigh1
        End If
    Else
        mInitialSwingStartPrice = mSwing.SwingHigh(0)
        If Not mUseIntermediateStops Then
            mReferencePrice = mSwing.SwingHigh(0)
'        ElseIf swingHigh1 > mSwing.swingHigh(0) Then
''            mReferencePrice = swingHigh2
''                        ' this should really be swingHigh1 but testing seems to give
''                        ' better results with swingHigh2
'            If swingHigh2 > swingHigh1 Then
'                mReferencePrice = swingHigh2
'            Else
'                mReferencePrice = swingHigh1
'            End If
        Else
            mReferencePrice = mSwing.SwingHigh(0)
        End If
    End If
End If
mInitialReferencePrice = mReferencePrice
If mInitialReferencePrice = 0 Then Stop
End Sub

Private Sub mBracketOrder_StopOrderFilled()
RaiseEvent StoppedOut(Not mLongPosition, Not mStopChanged)
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let BackstopMABarLength(ByVal value As Integer)
mBackstopMABarLength = value
End Property

Public Property Let BackstopMAPeriods(ByVal value As Integer)
mBackstopMAPeriods = value
End Property

Public Property Let BackstopMAOffset(ByVal value As Integer)
mBackstopMAOffset = value
End Property

Public Property Let BreakevenBarLength(ByVal value As Long)
mBreakevenBarLength = value
End Property

Public Property Let BreakevenBarThresholdTicks(ByVal value As Long)
mBreakevenBarThresholdTicks = value
End Property

Public Property Let breakevenPrice(ByVal value As Integer)
mBreakevenPrice = value
End Property

Public Property Let breakEvenThresholdTicks(ByVal value As Long)
mBreakEvenThresholdTicks = value
End Property

Public Property Let breakoutThresholdTicks(ByVal value As Integer)
mBreakoutThreshold = value
End Property

Public Property Let confirmedCountForBreakEven(ByVal value As Long)
mConfirmedCountForBreakEven = value
End Property

Public Property Let confirmedCountForExit(ByVal value As Long)
mConfirmedCountForExit = value
End Property

Public Property Let confirmedCountForSwing(ByVal value As Long)
mConfirmedCountForSwing = value
End Property

Public Property Let confirmedCountForTrailing(ByVal value As Long)
mConfirmedCountForTrailing = value
End Property

Public Property Let longBollPeriods(ByVal newValue As Integer)
mLongBollPeriods = newValue
End Property

Public Property Let longMAPeriods(ByVal newValue As Integer)
mLongMAPeriods = newValue
End Property

Public Property Let longTimePeriod(ByVal newValue As Integer)
mLongTimePeriod = newValue
End Property

Public Property Let longTimePeriodUnit(ByVal newValue As String)
mLongTimePeriodUnit = newValue
End Property

Public Property Let MaxContraSwingPercent(ByVal value As Single)
If value = 0 Then
    mMaxContraSwingFactor = 0
Else
    mMaxContraSwingFactor = 1! - value / 100
End If
End Property

Public Property Let minimumSwing(ByVal value As Double)
mMinimumSwing = value
End Property

Public Property Let noContraSwing(ByVal val As Boolean)
If val Then mTrailStop = True
End Property

Public Property Let ReferencePoint(ByVal value As SwingPoint)
mReferencePrice = value.value
mReferenceBar = value.barNumber
mInitialReferencePrice = value.value
mLowest = mReferencePrice
mHighest = mReferencePrice
mSwingCount = 0
mRestartCount = 0
End Property

Public Property Let RetraceFromExtremes(ByVal value As Long)
mRetraceFromExtremes = value
End Property

Public Property Let RetracementStopPercent(ByVal value As Single)
If value = 0 Then
    mRetracementStopFactor = 0
Else
    mRetracementStopFactor = 1! - value / 100
End If
End Property

Public Property Let RetracementStopThresholdTicks(ByVal value As Long)
mRetracementStopThresholdTicks = value
End Property

Public Property Let RewardToRiskRatio(ByVal value As Long)
mRewardToRiskRatio = value
End Property

Public Property Let TicksSwingToBreakeven(ByVal value As Integer)
' property no longer used
Err.Raise vbObjectError + 512, _
            "Strategies2.cStopStrategy5::TicksSwingToBreakeven", _
            "Property not implemented"
End Property

Public Property Let TicksSwingToMoveStop(ByVal value As Long)
mTicksSwingToMoveStop = value
End Property

Public Property Let TicksSwingToTrail(ByVal value As Integer)
mTicksSwingToTrail = value
End Property

Public Property Let timePeriod(ByVal newValue As Integer)
mTimePeriod = newValue
End Property

Public Property Let TimePeriodUnit(ByVal newValue As String)
mTimePeriodUnit = newValue
End Property

Public Property Let UseIntermediateStops(ByVal value As Boolean)
mUseIntermediateStops = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Initialise( _
                ByVal pTradingContext As TradingContext)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContext = pTradingContext
mTickSize = pTradingContext.Contract.TickSize

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcNextSwingStopCandidate(ByVal price As Double)
Dim i As Long
Static prevSwingStopCandidate As Double

If mNextSwingStopCandidate <> 0 Then
    prevSwingStopCandidate = mNextSwingStopCandidate
End If

If mLongPosition Then
    i = 1
    If mSwing.SwingLow(0) = DummyLow Then
        i = 2
        mNextSwingStopCandidate = mSwing.SwingLow(1)
    Else
        mNextSwingStopCandidate = mSwing.SwingLow(0)
    End If
    Do While (price - mSwing.SwingLow(i)) < mTicksSwingToMoveStop * mTickSize And _
            mSwing.SwingLow(i) < mNextSwingStopCandidate And _
            mSwing.SwingLow(i) > mLastActivatedSwingStopCandidate
        mNextSwingStopCandidate = mSwing.SwingLow(i)
        i = i + 1
    Loop
Else
    i = 1
    If mSwing.SwingHigh(0) = DummyHigh Then
        i = 2
        mNextSwingStopCandidate = mSwing.SwingHigh(1)
    Else
        mNextSwingStopCandidate = mSwing.SwingHigh(0)
    End If
    Do While (mSwing.SwingHigh(i) - price) < mTicksSwingToMoveStop * mTickSize And _
            mSwing.SwingHigh(i) > mNextSwingStopCandidate And _
            mSwing.SwingHigh(i) < mLastActivatedSwingStopCandidate
        mNextSwingStopCandidate = mSwing.SwingHigh(i)
        i = i + 1
    Loop
End If

If prevSwingStopCandidate = mNextSwingStopCandidate Then
    mNextSwingStopCandidate = 0
Else
    Debug.Print "Next swing stop candidate: " & mNextSwingStopCandidate
End If


End Sub

Private Function newbar(ByVal barNumber As Long) As Boolean
If barNumber > mBarNumber Then
    mBarNumber = barNumber
    newbar = True
Else
    newbar = False
End If
End Function

Private Sub setTrailingStop(ByVal price As Double, ByVal reason As String)
Dim i As Long
Dim firstLow As Double
Dim firstHigh As Double
Dim newStop As Double
Dim currentStop As Double

If mLongPosition Then
    If mBars.lowPrice(1) < mBars.lowPrice(0) Then
        firstLow = mBars.lowPrice(1)
        i = 2
    Else
        firstLow = mBars.lowPrice(0)
        i = 1
    End If
    Do While mBars.lowPrice(i) >= firstLow
        i = i + 1
    Loop
    newStop = mBars.lowPrice(i)
    currentStop = mStopOrder.triggerPrice
    If newStop > currentStop Then
        mBracketOrder.AdjustStop newStop, mBracketOrder.EntryOrder.QuantityFilled
        mBracketOrder.StopReason = reason
        mStopChanged = True
        RaiseEvent StopChanged(newStop)
    End If
Else
    If mBars.highPrice(1) > mBars.highPrice(0) Then
        firstHigh = mBars.highPrice(1)
        i = 2
    Else
        firstHigh = mBars.highPrice(0)
        i = 1
    End If
    Do While mBars.highPrice(i) <= firstHigh
        i = i + 1
    Loop
    newStop = mBars.highPrice(i)
    currentStop = mStopOrder.triggerPrice
    If newStop < currentStop Then
        mBracketOrder.AdjustStop newStop, mBracketOrder.EntryOrder.QuantityFilled
        mBracketOrder.StopReason = reason
        mStopChanged = True
        RaiseEvent StopChanged(newStop)
    End If
End If
        
End Sub








