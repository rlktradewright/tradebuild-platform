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

Private Const ModuleName                As String = "TickfileReader"

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
Const ProcName As String = "IStreamingDataConsumer_Tick"
On Error GoTo Err

mCurrentTick = Tick
mUnProcessed = True
mCurrentTick.Timestamp = AdjustTimestamp(mCurrentTick.Timestamp)
HandleTick

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' mTimer Event Handlers
'@================================================================================

Private Sub mTimer_TimerExpired()
Const ProcName As String = "mTimer_TimerExpired"
On Error GoTo Err

Play

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Contract(ByVal value As Contract)
Const ProcName As String = "Contract"
On Error GoTo Err

mReader.Contract = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get Contract() As Contract
Const ProcName As String = "Contract"
On Error GoTo Err

Set Contract = mReader.Contract

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get ContractSpecifier() As ContractSpecifier
Const ProcName As String = "ContractSpecifier"
On Error GoTo Err

Set ContractSpecifier = mReader.ContractSpecifier

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let dataConsumer(ByVal value As IStreamingDataConsumer)
Const ProcName As String = "dataConsumer"
On Error GoTo Err

Set mDataConsumer = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get firstTickFileTime() As Date
Const ProcName As String = "firstTickFileTime"
On Error GoTo Err

firstTickFileTime = mReader.firstTickFileTime

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let percentComplete(ByVal value As Single)
Const ProcName As String = "percentComplete"
On Error GoTo Err

mPercentComplete = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let ReplayProgressEventFrequency(ByVal value As Long)
Const ProcName As String = "ReplayProgressEventFrequency"
On Error GoTo Err

mReader.ReplayProgressEventFrequency = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get ReplayProgressEventFrequency() As Long
Const ProcName As String = "ReplayProgressEventFrequency"
On Error GoTo Err

ReplayProgressEventFrequency = mReader.ReplayProgressEventFrequency

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let ReplaySpeed(ByVal value As Long)
Const ProcName As String = "ReplaySpeed"
On Error GoTo Err

mReplaySpeed = value
If mState = TickfileStateCodes.TickfilePlaying Then
    setBaseTimes
    Set mTimer = CreateIntervalTimer(1)
    mTimer.StartTimer
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get ReplaySpeed() As Long
Const ProcName As String = "ReplaySpeed"
On Error GoTo Err

ReplaySpeed = mReplaySpeed

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let ServiceProviderReader(ByVal value As TradeBuildSP.ITickfileReader)
Const ProcName As String = "ServiceProviderReader"
On Error GoTo Err

Set mReader = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get State() As TickfileStateCodes
Const ProcName As String = "State"
On Error GoTo Err

State = mState

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TickfileSizeBytes() As Long
Const ProcName As String = "TickfileSizeBytes"
On Error GoTo Err

TickfileSizeBytes = mReader.TickfileSizeBytes

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let TimestampAdjustmentStart(ByVal value As Double)
Const ProcName As String = "TimestampAdjustmentStart"
On Error GoTo Err

mTimestampAdjustmentStart = value / 86400

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get TimestampAdjustmentStart() As Double
Const ProcName As String = "TimestampAdjustmentStart"
On Error GoTo Err

TimestampAdjustmentStart = mTimestampAdjustmentStart * 86400

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let TimestampAdjustmentEnd(ByVal value As Double)
Const ProcName As String = "TimestampAdjustmentEnd"
On Error GoTo Err

mTimestampAdjustmentEnd = value / 86400

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get TimestampAdjustmentEnd() As Double
Const ProcName As String = "TimestampAdjustmentEnd"
On Error GoTo Err

TimestampAdjustmentEnd = mTimestampAdjustmentEnd * 86400

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub PauseReplay()
Const ProcName As String = "PauseReplay"
On Error GoTo Err

If Not mTimer Is Nothing Then mTimer.StopTimer
mState = TickfileStateCodes.TickfilePaused
'gPauseTime

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub ReplayTickfile(ByVal tfSpec As TickfileSpecifier)
Const ProcName As String = "ReplayTickfile"
On Error GoTo Err

mReader.ReplayTickfile tfSpec

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub SkipFile()
Const ProcName As String = "SkipFile"
On Error GoTo Err

mSkip = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub StartReplay()

Const ProcName As String = "StartReplay"
On Error GoTo Err

If mState = TickfileStateCodes.TickfilePlaying Then Exit Sub

mState = TickfileStateCodes.TickfilePlaying

Set mElapsedTimer = New ElapsedTimer

' start replay using the timer to unwind the current stack
Set mTimer = CreateIntervalTimer(1)
mTimer.StartTimer

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Friend Sub StopReplay()
Const ProcName As String = "StopReplay"
On Error GoTo Err

mUnProcessed = False
'If mState = TickfileStateCodes.TickfileNotPlaying Then Exit Sub
If Not mTimer Is Nothing Then mTimer.StopTimer
mState = TickfileStateCodes.TickfileNotPlaying
mBaseTimestampUTC = 0
mReader.CloseInputFile
Set mReader = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function AdjustTimestamp(ByVal Timestamp As Date) As Date
Const ProcName As String = "AdjustTimestamp"
On Error GoTo Err

AdjustTimestamp = Timestamp + _
                mTimestampAdjustmentStart + _
                (mTimestampAdjustmentEnd - mTimestampAdjustmentStart) * (mPercentComplete / 100)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Private Sub CloseInputFile()
Const ProcName As String = "CloseInputFile"
On Error GoTo Err

mReader.CloseInputFile
Set mReader = Nothing
mState = TickfileStateCodes.TickfileNotPlaying

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub HandleTick()
Dim delayMillisecs As Long

Const ProcName As String = "HandleTick"
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub Play()

Const ProcName As String = "Play"
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub ProcessTickData()
Const ProcName As String = "ProcessTickData"
On Error GoTo Err

mDataConsumer.Tick mCurrentTick
mUnProcessed = False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setBaseTimes()
Const ProcName As String = "setBaseTimes"
On Error GoTo Err

mBaseTimestampUTC = GetTimestampUTC
' note that we don't bother to get the tick time in UTC - this is because
' it's not likely that anyone would replay a tickfile that spans a standard
' time to daylight time transition at normal speed. On the other hand
' the base Timestamp is in UTC because someone might replay a tickfile
' during a period that spans such a transition
mFirstTickTime = mCurrentTick.Timestamp

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub
