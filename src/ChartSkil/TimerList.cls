VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Maintains a list of TimerListEntry objects and raises events when they expire."
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

' Fires when a TimerListEntry object has expired. The Data argument is the
' Data property from the expired TimerListEntry object.
Event Expired(ByVal Data As Object)
Attribute Expired.VB_Description = "Fires when a TimerListEntry object has expired. The Data argument is the Data property from the expired TimerListEntry object."

' Fires when a TimerListEntry object has expired. The ExpiredEntry argument is the
' TimerListEntry object which has expired.
Event ExpiredTimerListEntry(ByVal ExpiredEntry As TimerListEntry)
Attribute ExpiredTimerListEntry.VB_Description = "Fires when a TimerListEntry object has expired. The ExpiredEntry argument is the TimerListEntry object which has expired."

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

Public Enum ExpiryTimeUnits
    DateTime
    RelativeSeconds
    RelativeMilliseconds
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1
Private mTimerListRoot As TimerListEntry

Private mTimerIntervalMillisecs As Long
Private mUseCallerSuppliedTime As Boolean
Private mState As TimerStates
Private mTime As Date   ' only relevant when mUseCallerSuppliedTime is true - used
                        ' to hold the most recently notified time

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTimer = New IntervalTimer
mTimer.RepeatNotifications = True
mTimer.TimerIntervalSecs = 1
mState = Stopped
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
CheckTimer GetTimestamp
End Sub

'================================================================================
' Properties
'================================================================================

' Indicates whether the TimerList is running.
Public Property Get State() As TimerStates
Attribute State.VB_Description = "Indicates whether the TimerList is running."
State = mState
End Property

' Specifies the frequency in milliseconds with which the TimerList should check
' for expiry of TimerListEntry objects.
Public Property Let TimerIntervalMillisecs(ByVal value As Long)
Attribute TimerIntervalMillisecs.VB_Description = "Specifies the frequency in milliseconds with which the TimerList should check for expiry of TimerListEntry objects."
If value <= 0 Then
    err.Raise InvalidPropertyValue, _
                "IntervalTimer.TimerList::(Let)TimerIntervalMillisecs", _
                "Value must be greater than zero"
End If
mTimerIntervalMillisecs = value
mTimer.TimerIntervalMillisecs = mTimerIntervalMillisecs
End Property

' Specifies the frequency in seconds with which the TimerList should check for
' expiry of TimerListEntry objects.
Public Property Let TimerIntervalSecs(ByVal value As Long)
Attribute TimerIntervalSecs.VB_Description = "Specifies the frequency in seconds with which the TimerList should check for expiry of TimerListEntry objects."
If value <= 0 Then
    err.Raise InvalidPropertyValue, _
                "IntervalTimer.TimerList::(Let)TimerIntervalSecs", _
                "Value must be greater than zero"
End If
mTimerIntervalMillisecs = value * 1000
mTimer.TimerIntervalSecs = value
End Property

' If true, indicates that the TimerList object is not to use an internal timer
' to determine when TimerListEntry objects have expired. Instead, time will be
' notified by the caller via the SetTime method.
Public Property Let UseCallerSuppliedTime(ByVal value As Boolean)
Attribute UseCallerSuppliedTime.VB_Description = "If true, indicates that the TimerList object is not to use an internal timer to determine when TimerListEntry objects have expired. Instead, time will be notified by the caller via the SetTime method."
mUseCallerSuppliedTime = value
If mUseCallerSuppliedTime Then
    mTimer.StopTimer
Else
    If mState = TimerStates.Running Then
        mTimer.StartTimer
    End If
End If
End Property

Public Property Get UseCallerSuppliedTime() As Boolean
UseCallerSuppliedTime = mUseCallerSuppliedTime
End Property

'================================================================================
' Methods
'================================================================================

' Creates a new TimerListEntry object.
Public Function AddTimerListEntry(ByVal ExpiryTime As Double, _
                        ByVal units As ExpiryTimeUnits, _
                        ByVal Data As Object) As TimerListEntry
Attribute AddTimerListEntry.VB_Description = "Creates a new TimerListEntry object."
Dim entry As TimerListEntry
Dim NextEntry As TimerListEntry
Dim prevEntry As TimerListEntry

Set AddTimerListEntry = New TimerListEntry
Set entry = AddTimerListEntry

Select Case units
Case ExpiryTimeUnits.DateTime

Case ExpiryTimeUnits.RelativeMilliseconds
    If mUseCallerSuppliedTime Then
        ExpiryTime = mTime + (ExpiryTime / 86400000)
    Else
        ExpiryTime = GetTimestamp + (ExpiryTime / 86400000)
    End If
Case ExpiryTimeUnits.RelativeSeconds
    If mUseCallerSuppliedTime Then
        ExpiryTime = mTime + (ExpiryTime / 86400)
    Else
        ExpiryTime = GetTimestamp + (ExpiryTime / 86400)
    End If
End Select

If ExpiryTime <= IIf(mUseCallerSuppliedTime, mTime, GetTimestamp) Then
    err.Raise ExpiryTimeInvalid, _
                "IntervalTimer.TimerList::AddTimerListEntry", _
                "Expiry time is before present"
End If

entry.ExpiryTime = ExpiryTime
entry.Data = Data

If mTimerListRoot Is Nothing Then
    Set mTimerListRoot = entry
    If Not mUseCallerSuppliedTime And _
        mState = TimerStates.Running And _
        mTimer.State = Stopped _
    Then mTimer.StartTimer
ElseIf mTimerListRoot.ExpiryTime >= entry.ExpiryTime Then
    entry.NextEntry = mTimerListRoot
    Set mTimerListRoot = entry
Else
    Set prevEntry = mTimerListRoot
    Set NextEntry = prevEntry.NextEntry
    
    Do While Not NextEntry Is Nothing
        If NextEntry.ExpiryTime >= entry.ExpiryTime Then
            entry.NextEntry = NextEntry
            Exit Do
        End If
        Set prevEntry = NextEntry
        Set NextEntry = NextEntry.NextEntry
    Loop
    prevEntry.NextEntry = entry
End If
 
End Function

' Removes the specified TimerListEntry object from the TimerList.
Public Sub RemoveTimerListEntry(ByVal theTimer As TimerListEntry)
Attribute RemoveTimerListEntry.VB_Description = "Removes the specified TimerListEntry object from the TimerList."
Dim entry As TimerListEntry
Dim NextEntry As TimerListEntry
Dim prevEntry As TimerListEntry

Set entry = theTimer

If mTimerListRoot Is Nothing Then Exit Sub

If mTimerListRoot Is entry Then
    Set mTimerListRoot = entry.NextEntry
    entry.NextEntry = Nothing
Else
    Set prevEntry = mTimerListRoot
    Set NextEntry = prevEntry.NextEntry
    
    Do While Not NextEntry Is Nothing
        If NextEntry Is entry Then
            prevEntry.NextEntry = NextEntry.NextEntry
            entry.NextEntry = Nothing
            Exit Do
        End If
        Set prevEntry = NextEntry
        Set NextEntry = NextEntry.NextEntry
    Loop
End If

If mTimerListRoot Is Nothing Then
    If Not mUseCallerSuppliedTime Then mTimer.StopTimer
End If
End Sub

' Notifies the TimerList object of the current time. The TimerList object uses
' this to determine if any TimerListEntry objects have expired. NB: this method
' may only be called if the UseCallerSuppliedTime property is True and StartTiming
' has been called.
Public Sub SetTime(ByVal theTime As Date)
Attribute SetTime.VB_Description = "Notifies the TimerList object of the current time. The TimerList object uses this to determine if any TimerListEntry objects have expired. NB: this method may only be called if the UseCallerSuppliedTime property is True and StartTiming has been called."
If Not mUseCallerSuppliedTime Then
    err.Raise CallerCantSetTime, _
                "IntervalTimer.TimerList::SetTime", _
                "Caller can't set time when UseCallerSuppliedTime is false"
End If
If Not mState = TimerStates.Running Then
    err.Raise CallInvalidInThisState, _
                "IntervalTimer.TimerList::SetTime", _
                "Can't set time in the Stopped state."
End If

mTime = theTime
CheckTimer theTime
End Sub

' Tells the TimerList object to start checking for expiry of TimerListEntry objects.
Public Sub StartTiming()
Attribute StartTiming.VB_Description = "Tells the TimerList object to start checking for expiry of TimerListEntry objects."
If Not mTimerListRoot Is Nothing Then
    If Not mUseCallerSuppliedTime Then mTimer.StartTimer
End If
mState = TimerStates.Running
End Sub

' Tells the TimerList object to stop checking for expiry of TimerListEntry objects.
Public Sub StopTiming()
Attribute StopTiming.VB_Description = "Tells the TimerList object to stop checking for expiry of TimerListEntry objects."
mTimer.StopTimer
mState = TimerStates.Stopped
End Sub


'================================================================================
' Helper Functions
'================================================================================


Private Sub CheckTimer(ByVal currentTime As Date)
Dim currEntry As TimerListEntry
Do While Not mTimerListRoot Is Nothing
    Set currEntry = mTimerListRoot  ' use a temp variable in case the caller
                                    ' removes this entry when processing the event
                                    ' raised in the following call
    If Not mTimerListRoot.CheckExpiry(currentTime) Then Exit Do
    RaiseEvent Expired(currEntry.Data)
    RaiseEvent ExpiredTimerListEntry(currEntry)
    If mTimerListRoot Is currEntry Then
        ' current entry hasn't been deleted during the event processing
        Set mTimerListRoot = currEntry.NextEntry
    End If
Loop
End Sub

Private Function GetTimestamp() As Date
GetTimestamp = CDbl(Int(Now)) + (CDbl(Timer) / 86400#)
End Function


