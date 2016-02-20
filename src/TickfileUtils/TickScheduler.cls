VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type PlayerEntry
    Player              As TickfilePlayer
    Tick                As GenericTick
    UTCTimestamp        As Date
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TickScheduler"

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileReplayController                   As ReplayController
Private mPlayers()                                  As PlayerEntry

Private mNumberOfTicksAwaited                       As Long

Private mReplaySpeed                                As Long

Private mTimerList                                  As TimerList
Private WithEvents mProgressTimerListItem           As TimerListItem
Attribute mProgressTimerListItem.VB_VarHelpID = -1

Private WithEvents mTickTimer                       As IntervalTimer
Attribute mTickTimer.VB_VarHelpID = -1

Private mBaseRealTimeUTC                            As Date
Private mBaseSimulatedTimeUTC                       As Date

Private mStarted                                    As Boolean

Private mReplayProgressEventInterval                As Long

Private mTicksPlayed                                As Long

Private mClock                                      As Clock
Private mClockNeedsSetting                          As Boolean

Private mSingleSource                               As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mTimerList = GetGlobalTimerList
ReDim mPlayers(3) As PlayerEntry
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mProgressTimerListItem Event Handlers
'@================================================================================

Private Sub mProgressTimerListItem_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mProgressTimerListItem_StateChange"
On Error GoTo Err

Dim lState As TimerListItemStates
lState = ev.State
If lState <> TimerListItemStateExpired Then Exit Sub

fireProgress
Set mProgressTimerListItem = mTimerList.Add(Empty, mReplayProgressEventInterval, ExpiryTimeUnitMilliseconds)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTickTimer Event Handlers
'@================================================================================

Private Sub mTickTimer_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mTickTimer_TimerExpired"
On Error GoTo Err

Dim lIndex As Long
lIndex = CLng(ev.Data)
Set mTickTimer = Nothing
fireTick lIndex

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let ReplaySpeed(ByVal Value As Long)
Const ProcName As String = "ReplaySpeed"
On Error GoTo Err

mReplaySpeed = Value

undispatch

If Not mClock Is Nothing Then
    mBaseSimulatedTimeUTC = mClock.TimestampUTC
Else
    mBaseSimulatedTimeUTC = 0
End If
mBaseRealTimeUTC = GetTimestampUTC

schedule

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddPlayer(ByVal pPlayer As TickfilePlayer)
Const ProcName As String = "AddPlayer"
On Error GoTo Err

Dim lIndex As Long
lIndex = getPlayerIndex(pPlayer)

Do While lIndex > UBound(mPlayers)
    ReDim Preserve mPlayers(2 * (UBound(mPlayers) + 1) - 1) As PlayerEntry
Loop

Set mPlayers(lIndex).Player = pPlayer
mPlayers(lIndex).UTCTimestamp = MaxDate
mNumberOfTicksAwaited = mNumberOfTicksAwaited + 1

If mSingleSource Then
    Set mClock = pPlayer.Clock
ElseIf mClock Is Nothing Then
    Set mClock = CreateSimulatedClock(mReplaySpeed, "")
    mClockNeedsSetting = True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTickfileReplayController As ReplayController, _
                ByVal pReplaySpeed As Long, _
                ByVal pReplayProgressEventInterval As Long, _
                ByVal pSingleSource As Boolean)
Set mTickfileReplayController = pTickfileReplayController
mReplaySpeed = pReplaySpeed
mReplayProgressEventInterval = pReplayProgressEventInterval
mSingleSource = pSingleSource
End Sub

Friend Sub NotifyTick(ByRef pTick As GenericTick, ByVal pPlayer As TickfilePlayer)
Const ProcName As String = "NotifyTick"
On Error GoTo Err

Static sFirstTickNotified As Boolean
Static sFirstTickTimeUTC As Date

If sFirstTickTimeUTC = 0# Then sFirstTickTimeUTC = MaxDate

Dim lIndex As Long
lIndex = getPlayerIndex(pPlayer)
mPlayers(lIndex).Tick = pTick
mPlayers(lIndex).UTCTimestamp = ConvertDateTzToUTC(pTick.Timestamp, pPlayer.TimeZone)

mNumberOfTicksAwaited = mNumberOfTicksAwaited - 1

If Not sFirstTickNotified Then
    If mPlayers(lIndex).UTCTimestamp < sFirstTickTimeUTC Then sFirstTickTimeUTC = mPlayers(lIndex).UTCTimestamp
    If mNumberOfTicksAwaited = 0 Then
        sFirstTickNotified = True
        notifyFirstTickTime sFirstTickTimeUTC
    End If
End If

schedule

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub PauseReplay()
Const ProcName As String = "PauseReplay"
On Error GoTo Err

mClock.PauseTime
undispatch
unProgress
fireProgress
'debug.print FormatTimestamp(GetTimestamp, TimestampDateAndTimeISO8601) & "  Paused: " & FormatTimestamp(mClock.Timestamp, TimestampDateAndTimeISO8601)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemovePlayer(ByVal pPlayer As TickfilePlayer)
Const ProcName As String = "RemovePlayer"
On Error GoTo Err

Dim lIndex As Long
lIndex = getPlayerIndex(pPlayer)

Assert Not mPlayers(lIndex).Player Is Nothing

Set mPlayers(lIndex).Player = Nothing
If mPlayers(lIndex).UTCTimestamp = MaxDate Then
    ' a tick is awaited from this player
    mNumberOfTicksAwaited = mNumberOfTicksAwaited - 1
End If

schedule

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ResumeReplay()
    Const ProcName As String = "ResumeReplay"
On Error GoTo Err

mClock.ResumeTime
mBaseSimulatedTimeUTC = mClock.TimestampUTC
mBaseRealTimeUTC = GetTimestampUTC
'debug.print FormatTimestamp(GetTimestamp, TimestampDateAndTimeISO8601) & "  Resumed: " & FormatTimestamp(mClock.Timestamp, TimestampDateAndTimeISO8601)
Set mProgressTimerListItem = mTimerList.Add(Empty, mReplayProgressEventInterval, ExpiryTimeUnitMilliseconds)
schedule

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub StartScheduling()
Const ProcName As String = "StartScheduling"
On Error GoTo Err

mStarted = True
Set mProgressTimerListItem = mTimerList.Add(Empty, mReplayProgressEventInterval, ExpiryTimeUnitMilliseconds)
schedule

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub StopScheduling()
Const ProcName As String = "StopScheduling"
On Error GoTo Err

fireProgress
mStarted = False
undispatch
unProgress

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub dispatchTick(ByVal pIndex As Long)
Const ProcName As String = "dispatchTick"
On Error GoTo Err

Dim lCurrTime As Date
lCurrTime = GetTimestampUTC

If mBaseSimulatedTimeUTC = 0 Then
    mBaseSimulatedTimeUTC = mPlayers(pIndex).UTCTimestamp
    mBaseRealTimeUTC = lCurrTime
End If

Dim delayMillisecs As Currency
If mReplaySpeed > 0 Then
    delayMillisecs = CCur(((mPlayers(pIndex).UTCTimestamp - mBaseSimulatedTimeUTC) - _
                mReplaySpeed * (lCurrTime - mBaseRealTimeUTC)) * 86400# * 1000# / CDbl(mReplaySpeed))
Else
    delayMillisecs = -mReplaySpeed
End If

'' adjust to allow for overheads
'delayMillisecs = delayMillisecs - 2

If delayMillisecs > &H7FFFFFFF Then
    ' the delay is more than 28 days! We can't handle this so don't
    ' dispatch the tick
ElseIf delayMillisecs < 5 Then
    ' process immediately
    'debug.print FormatTimestamp(GetTimestamp, TimestampDateAndTimeISO8601) & "  Delay 0"
    fireTick pIndex
Else
    Set mTickTimer = CreateIntervalTimer(delayMillisecs, ExpiryTimeUnitMilliseconds, , , pIndex)
    mTickTimer.StartTimer
    'debug.print FormatTimestamp(GetTimestamp, TimestampDateAndTimeISO8601) & "  Delay " & delayMillisecs
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireProgress()
Const ProcName As String = "fireProgress"
On Error GoTo Err

If mTicksPlayed = 0 Then Exit Sub
Dim lTimestamp As Date
lTimestamp = mClock.Timestamp
'debug.print FormatTimestamp(GetTimestamp, TimestampDateAndTimeISO8601) & "  Notify progress: " & FormatTimestamp(lTimestamp, TimestampDateAndTimeISO8601)

If Not mSingleSource Then
    mTickfileReplayController.NotifyProgress lTimestamp, mTicksPlayed, 0
ElseIf mPlayers(0).Player Is Nothing Then
    mTickfileReplayController.NotifyProgress lTimestamp, mTicksPlayed, 100
Else
    mTickfileReplayController.NotifyProgress lTimestamp, mTicksPlayed, mPlayers(0).Player.ReplayProgressPercent
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireTick(ByVal pIndex As Long)
Const ProcName As String = "fireTick"
On Error GoTo Err

If mClockNeedsSetting Then mClock.SetTime ConvertDateUTCToLocal(mPlayers(pIndex).UTCTimestamp)

If Not mPlayers(pIndex).Player Is Nothing Then
    mPlayers(pIndex).Player.ProcessTick mPlayers(pIndex).Tick
    mPlayers(pIndex).UTCTimestamp = MaxDate
    mTicksPlayed = mTicksPlayed + 1
    
    mNumberOfTicksAwaited = mNumberOfTicksAwaited + 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getIndexToSchedule() As Long
Const ProcName As String = "getIndexToSchedule"
On Error GoTo Err

Dim lMinIndex As Long
lMinIndex = -1

Dim lMinTime As Date
lMinTime = MaxDate

Dim i As Long
For i = 0 To UBound(mPlayers)
    If Not mPlayers(i).Player Is Nothing Then
        If mPlayers(i).UTCTimestamp < lMinTime Then
            lMinTime = mPlayers(i).UTCTimestamp
            lMinIndex = i
        End If
    End If
Next

getIndexToSchedule = lMinIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getPlayerIndex(ByVal pPlayer As TickfilePlayer) As Long
getPlayerIndex = IIf(mSingleSource, 0, pPlayer.TickStream.StreamId)
End Function

Private Sub notifyFirstTickTime(ByVal pUTCTimestamp As Date)
Const ProcName As String = "notifyFirstTickTime"
On Error GoTo Err

Dim i As Long
For i = 0 To UBound(mPlayers)
    If Not mPlayers(i).Player Is Nothing Then mPlayers(i).Player.FirstTickTimeUTC = pUTCTimestamp
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub schedule()
Const ProcName As String = "schedule"
On Error GoTo Err

If Not mStarted Then Exit Sub
If mNumberOfTicksAwaited <> 0 Then Exit Sub
If Not mTickTimer Is Nothing Then Exit Sub

Dim lIndex As Long
lIndex = getIndexToSchedule

If lIndex >= 0 Then dispatchTick lIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub undispatch()
Const ProcName As String = "undispatch"
On Error GoTo Err

If mTickTimer Is Nothing Then Exit Sub

mTickTimer.StopTimer
Set mTickTimer = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub unProgress()
Const ProcName As String = "unProgress"
On Error GoTo Err

If Not mProgressTimerListItem Is Nothing Then
    mProgressTimerListItem.Cancel
    Set mProgressTimerListItem = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
