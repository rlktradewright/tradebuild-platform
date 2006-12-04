VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntervalTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Objects of this class generate events at specified intervals, which may be regular or random."
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'Fires when the IntervalTimer object's timer interval has expired.
Public Event TimerExpired()
Attribute TimerExpired.VB_Description = "Fires when the IntervalTimer object's timer interval has expired."

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

Public Enum ErrorCodes
    CantCreateTimer = vbObjectError + 512
    CantKillTimer
    InvalidPropertyValue
    TimerIntervalNotSet
    ExpiryTimeInvalid
    CallerCantSetTime
    CallInvalidInThisState
End Enum

Public Enum TimerStates
    Running = 1
    Stopped
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mFirstTimerIntervalMillisecs As Long
Private mTimerIntervalMillisecs As Long
Private mRepeat As Boolean
Private mTimerNumber As Long
Private mRandomIntervals As Boolean
Private mState As TimerStates
Private mFirstInterval As Boolean
Private mFirstExpiryNotified As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mState = Stopped
End Sub

Private Sub Class_Terminate()
Debug.Print "IntervalTimer terminated"
If mTimerNumber <> 0 Then EndTimer mTimerNumber
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

' Specifies the date and time at which the IntervalTimer object is to fire the first
' TimerExpired event.
Public Property Let FirstExpiryTime(ByVal value As Date)
Attribute FirstExpiryTime.VB_Description = "Specifies the date and time at which the IntervalTimer object is to fire the first TimerExpired event."
If value < Now Then
    err.Raise ErrorCodes.InvalidPropertyValue, _
                "IntervalTimer.Timer::(Let)FirstExpiryTime", _
                "Value is before present"
End If
mFirstTimerIntervalMillisecs = (value - Now) * 86400 * 1000
If mTimerNumber <> 0 Then StartTimer
End Property

' Specifies the interval in milliseconds after which the IntervalTimer object is to fire
' the first TimerExpired event.
Public Property Let FirstTimerIntervalMillisecs(newValue As Long)
Attribute FirstTimerIntervalMillisecs.VB_Description = "Specifies the interval in milliseconds after which the IntervalTimer object is to fire the first TimerExpired event."
If newValue < 0 Then
    err.Raise InvalidPropertyValue, _
                "IntervalTimer.Timer::(Let)FirstTimerIntervalMillisecs", _
                "Value cannot be less than zero"
End If
mFirstTimerIntervalMillisecs = newValue
If mTimerNumber <> 0 Then StartTimer
End Property

' Specifies the interval in seconds after which the IntervalTimer object is to fire
' the first TimerExpired event.
Public Property Let FirstTimerIntervalSecs(newValue As Long)
Attribute FirstTimerIntervalSecs.VB_Description = "Specifies the interval in seconds after which the IntervalTimer object is to fire the first TimerExpired event."
If newValue < 0 Then
    err.Raise InvalidPropertyValue, _
                "IntervalTimer.Timer::(Let)FirstTimerIntervalSecs", _
                "Value cannot be less than zero"
End If
mFirstTimerIntervalMillisecs = newValue * 1000
If mTimerNumber <> 0 Then StartTimer
End Property

' If set to True, indicates that TimerExpired events are to be fired at random intervals
' not exceeding the time specified in TimerIntervalMillsecs or TimerIntervalSecs.
Public Property Let RandomIntervals(ByVal value As Boolean)
Attribute RandomIntervals.VB_Description = "If set to True, indicates that TimerExpired events are to be fired at random intervals not exceeding the time specified in TimerIntervalMillsecs or TimerIntervalSecs."
mRandomIntervals = value
End Property

Public Property Get RandomIntervals() As Boolean
RandomIntervals = mRandomIntervals
End Property

' If set to True, indicates that TimerExpired events are to be fired regularly at intervals
' specified in TimerIntervalMillsecs or TimerIntervalSecs. Otherwise only a single
' TimerExpired event will occur.
Public Property Let RepeatNotifications(ByVal value As Boolean)
Attribute RepeatNotifications.VB_Description = "If set to True, indicates that TimerExpired events are to be fired regularly at intervals specified in TimerIntervalMillsecs or TimerIntervalSecs. Otherwise only a single TimerExpired event will occur."
mRepeat = value
End Property

Public Property Get RepeatNotifications() As Boolean
RepeatNotifications = mRepeat
End Property

' Indicates whether the IntervalTimer object is Running or Stopped.
Public Property Get State() As TimerStates
Attribute State.VB_Description = "Indicates whether the IntervalTimer object is Running or Stopped."
State = mState
End Property

' Specifies the interval in milliseconds at which the IntervalTimer object is to fire
' TimerExpired events.
Public Property Let TimerIntervalMillisecs(newValue As Long)
Attribute TimerIntervalMillisecs.VB_Description = "Specifies the interval in milliseconds at which the IntervalTimer object is to fire TimerExpired events."
If newValue <= 0 Then
    err.Raise InvalidPropertyValue, _
                "IntervalTimer.Timer::(Let)TimerIntervalMillisecs", _
                "Value must be greater than zero"
End If
mTimerIntervalMillisecs = newValue
If mTimerNumber <> 0 And Not mFirstInterval Then StartTimer
End Property

' Specifies the interval in seconds at which the IntervalTimer object is to fire
' TimerExpired events.
Public Property Let TimerIntervalSecs(newValue As Long)
Attribute TimerIntervalSecs.VB_Description = "Specifies the interval in seconds at which the IntervalTimer object is to fire TimerExpired events."
If newValue <= 0 Then
    err.Raise InvalidPropertyValue, _
                "IntervalTimer.Timer::(Let)TimerIntervalSecs", _
                "Value must be greater than zero"
End If
mTimerIntervalMillisecs = newValue * 1000
If mTimerNumber <> 0 And Not mFirstInterval Then StartTimer
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Notify()
Dim interval As Long

If Not mRepeat Then
    StopTimer
ElseIf mRandomIntervals Then
    StopTimer
    If mRepeat Then
        Randomize
        interval = Rnd * mTimerIntervalMillisecs + 1
        mTimerNumber = BeginTimer(interval, False, Me)
    End If
ElseIf mTimerIntervalMillisecs <> 0 And mFirstInterval Then
    mFirstInterval = False
    mTimerNumber = BeginTimer(mTimerIntervalMillisecs, True, Me)
End If
RaiseEvent TimerExpired
End Sub

' Tells the IntervalTimer object to start producing TimerExpired events.
Public Sub StartTimer()
Attribute StartTimer.VB_Description = "Tells the IntervalTimer object to start producing TimerExpired events."
Dim interval As Long
If mFirstTimerIntervalMillisecs = 0 And mTimerIntervalMillisecs = 0 Then
    err.Raise TimerIntervalNotSet, _
                "IntervalTimer.Timer::StartTimer", _
                "Timer interval has not been set"
End If

If mTimerNumber <> 0 Then StopTimer

If mFirstTimerIntervalMillisecs <> 0 Then
    mFirstInterval = True
    interval = mFirstTimerIntervalMillisecs
    If mRandomIntervals Then
        Randomize
        interval = Rnd * mFirstTimerIntervalMillisecs + 1
    End If
    mTimerNumber = BeginTimer(mFirstTimerIntervalMillisecs, False, Me)
ElseIf mRepeat Then
    If mRandomIntervals Then
        Randomize
        interval = Rnd * mTimerIntervalMillisecs + 1
        mTimerNumber = BeginTimer(interval, False, Me)
    Else
        mTimerNumber = BeginTimer(mTimerIntervalMillisecs, True, Me)
    End If
Else
    interval = mTimerIntervalMillisecs
    If mRandomIntervals Then
        Randomize
        interval = Rnd * mTimerIntervalMillisecs + 1
    End If
    mTimerNumber = BeginTimer(interval, False, Me)
End If
mState = TimerStates.Running
End Sub

' Tells the IntervalTimer object to stop producing TimerExpired events.
Public Sub StopTimer()
Attribute StopTimer.VB_Description = "Tells the IntervalTimer object to stop producing TimerExpired events."
mFirstInterval = False
mState = TimerStates.Stopped
If mTimerNumber = 0 Then Exit Sub
EndTimer mTimerNumber
mTimerNumber = 0
End Sub

'================================================================================
' Helper Functions
'================================================================================


