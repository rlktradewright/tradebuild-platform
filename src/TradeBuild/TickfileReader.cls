VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements TradeBuildSP.IStreamingDataConsumer

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private mUnProcessed            As Boolean
Private mTimed                  As Boolean

'@================================================================================
' Member variables
'@================================================================================

Private mCurrentTick As GenericTick

Private mReader As TradeBuildSP.ITickfileReader

Private mDataConsumer As TradeBuildSP.IStreamingDataConsumer

Private mReplaySpeed As Long

Private mState As TickfileStateCodes

Private mBaseTimestampUTC As Double
Private mFirstTickTime As Double

Private mSkip As Boolean

Private mTimestampAdjustmentStart As Double
Private mTimestampAdjustmentEnd As Double
Private mPercentComplete As Single

Private mTickReceivedSynchronously As Boolean

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1
Private mElapsedTimer As ElapsedTimer

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mState = TickfileStateCodes.TickfileNotPlaying
End Sub

Private Sub Class_Terminate()
Debug.Print "TickfileReader terminated"
End Sub

'@================================================================================
' TradeBuildSP.IStreamingDataConsumer Interface Members
'@================================================================================

Private Sub IStreamingDataConsumer_Tick(Tick As TickUtils26.GenericTick)
mCurrentTick = Tick
mUnProcessed = True
mCurrentTick.Timestamp = AdjustTimestamp(mCurrentTick.Timestamp)
HandleTick
End Sub

'@================================================================================
' mTimer Event Handlers
'@================================================================================

Private Sub mTimer_TimerExpired()
Play
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Contract(ByVal value As Contract)
mReader.Contract = value
End Property

Friend Property Get Contract() As Contract
Set Contract = mReader.Contract
End Property

Friend Property Get ContractSpecifier() As ContractSpecifier
Set ContractSpecifier = mReader.ContractSpecifier
End Property

Friend Property Let dataConsumer(ByVal value As IStreamingDataConsumer)
Set mDataConsumer = value
End Property

Friend Property Get firstTickFileTime() As Date
firstTickFileTime = mReader.firstTickFileTime
End Property

Friend Property Let percentComplete(ByVal value As Single)
mPercentComplete = value
End Property

Friend Property Let ReplayProgressEventFrequency(ByVal value As Long)
mReader.ReplayProgressEventFrequency = value
End Property

Friend Property Get ReplayProgressEventFrequency() As Long
ReplayProgressEventFrequency = mReader.ReplayProgressEventFrequency
End Property

Friend Property Let ReplaySpeed(ByVal value As Long)
mReplaySpeed = value
If mState = TickfileStateCodes.TickfilePlaying Then
    setBaseTimes
    Set mTimer = CreateIntervalTimer(1)
    mTimer.StartTimer
End If
End Property

Friend Property Get ReplaySpeed() As Long
ReplaySpeed = mReplaySpeed
End Property

Friend Property Let ServiceProviderReader(ByVal value As TradeBuildSP.ITickfileReader)
Set mReader = value
End Property

Public Property Get State() As TickfileStateCodes
State = mState
End Property

Public Property Get TickfileSizeBytes() As Long
TickfileSizeBytes = mReader.TickfileSizeBytes
End Property

Friend Property Let TimestampAdjustmentStart(ByVal value As Double)
mTimestampAdjustmentStart = value / 86400
End Property

Friend Property Get TimestampAdjustmentStart() As Double
TimestampAdjustmentStart = mTimestampAdjustmentStart * 86400
End Property

Friend Property Let TimestampAdjustmentEnd(ByVal value As Double)
mTimestampAdjustmentEnd = value / 86400
End Property

Friend Property Get TimestampAdjustmentEnd() As Double
TimestampAdjustmentEnd = mTimestampAdjustmentEnd * 86400
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub PauseReplay()
If Not mTimer Is Nothing Then mTimer.StopTimer
mState = TickfileStateCodes.TickfilePaused
'gPauseTime
End Sub

Friend Sub ReplayTickfile(ByVal tfSpec As TickfileSpecifier)
mReader.ReplayTickfile tfSpec
End Sub

Public Sub SkipFile()
mSkip = True
End Sub

Friend Sub StartReplay()

If mState = TickfileStateCodes.TickfilePlaying Then Exit Sub

mState = TickfileStateCodes.TickfilePlaying

Set mElapsedTimer = New ElapsedTimer

' start replay using the timer to unwind the current stack
Set mTimer = CreateIntervalTimer(1)
mTimer.StartTimer

End Sub

Friend Sub StopReplay()
mUnProcessed = False
'If mState = TickfileStateCodes.TickfileNotPlaying Then Exit Sub
If Not mTimer Is Nothing Then mTimer.StopTimer
mState = TickfileStateCodes.TickfileNotPlaying
mBaseTimestampUTC = 0
mReader.CloseInputFile
Set mReader = Nothing
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function AdjustTimestamp(ByVal Timestamp As Date) As Date
AdjustTimestamp = Timestamp + _
                mTimestampAdjustmentStart + _
                (mTimestampAdjustmentEnd - mTimestampAdjustmentStart) * (mPercentComplete / 100)

End Function

Private Sub CloseInputFile()
mReader.CloseInputFile
Set mReader = Nothing
mState = TickfileStateCodes.TickfileNotPlaying
End Sub

Private Sub HandleTick()
Dim delayMillisecs As Long

mTimed = False

If mBaseTimestampUTC = 0 Then
    setBaseTimes
End If

If mReplaySpeed > 0 Then
    delayMillisecs = CLng(((mCurrentTick.Timestamp - mFirstTickTime) - mReplaySpeed * (GetTimestampUTC - mBaseTimestampUTC)) * 86400# * 1000# / CDbl(mReplaySpeed))
'    delayMillisecs = CLng((mCurrentTick.Timestamp - mCurrentTimestamp) * 86400# * 1000# / CDbl(mReplaySpeed))
Else
    delayMillisecs = -mReplaySpeed
End If

If delayMillisecs < 5 Then
    ' process immediately. Note that if the tick was received synchronously, this
    ' is achieved simply by exiting from this procedure and winding back down into the
    ' Play routine. For ticks received asynchronously, we use the timer with a very short
    ' delay: this is to prevent calling back into the Service Provider before it has
    ' completed the function that raised the tick event.
    If Not mTickReceivedSynchronously Then
        mTimed = True
        Set mTimer = CreateIntervalTimer(1)
        mTimer.StartTimer
    End If
Else
    mTimed = True
    Set mTimer = CreateIntervalTimer(delayMillisecs)
    mTimer.StartTimer
End If

End Sub

Private Sub Play()

If mUnProcessed Then
    ProcessTickData
End If
    
mTickReceivedSynchronously = True
mReader.FireNextTick
mTickReceivedSynchronously = False

mElapsedTimer.StartTiming

Do While mUnProcessed And _
    Not mTimed And _
    (mElapsedTimer.ElapsedTimeMicroseconds / 1000) <= MultiTaskingTimeQuantumMillisecs
    
    ProcessTickData

    If mState = TickfileStateCodes.TickfileNotPlaying Or _
        mState = TickfileStateCodes.TickfilePaused _
    Then
        Exit Sub
    End If
    
    mTickReceivedSynchronously = True
    mReader.FireNextTick
    mTickReceivedSynchronously = False

Loop

If mUnProcessed And _
    Not mTimed _
Then
    Set mTimer = CreateIntervalTimer(1)
    mTimer.StartTimer
End If

End Sub

Private Sub ProcessTickData()
mDataConsumer.Tick mCurrentTick
mUnProcessed = False
End Sub

Private Sub setBaseTimes()
mBaseTimestampUTC = GetTimestampUTC
' note that we don't bother to get the tick time in UTC - this is because
' it's not likely that anyone would replay a tickfile that spans a standard
' time to daylight time transition at normal speed. On the other hand
' the base Timestamp is in UTC because someone might replay a tickfile
' during a period that spans such a transition
mFirstTickTime = mCurrentTick.Timestamp
End Sub
