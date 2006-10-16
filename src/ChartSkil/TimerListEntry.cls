VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimerListEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Objects of this class are created by the AddTimerListEntry method of the TimerList class."
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

' Fires when this TimerListEntry object expires.
Event Expired()
Attribute Expired.VB_Description = "Fires when this TimerListEntry object expires."

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mExpiryTime As Date
Private mNext As TimerListEntry
Private mUserObject As Object

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

' A reference to any required object. This object reference is returned in the
' Data argument to the TimerList object's Expired event when this TimerListEntry
' object expires.
Public Property Get Data() As Object
Attribute Data.VB_Description = "A reference to any required object. This object reference is returned in the Data argument to the TimerList object's Expired event when this TimerListEntry object expires."
Set Data = mUserObject
End Property

Friend Property Let Data(ByVal value As Object)
Set mUserObject = value
End Property

' Specifies the time at which this TimerListEntry object will expire.
Public Property Get ExpiryTime() As Date
Attribute ExpiryTime.VB_Description = "Specifies the time at which this TimerListEntry object will expire."
ExpiryTime = mExpiryTime
End Property

Friend Property Let ExpiryTime(ByVal value As Date)
mExpiryTime = value
End Property

Friend Property Let NextEntry(ByVal value As TimerListEntry)
Set mNext = value
End Property

Friend Property Get NextEntry() As TimerListEntry
Set NextEntry = mNext
End Property

'================================================================================
' Methods
'================================================================================

Friend Function CheckExpiry(ByVal currentTime As Date) As Boolean
If currentTime >= mExpiryTime Then
    RaiseEvent Expired
    CheckExpiry = True
End If
End Function

'================================================================================
' Helper Functions
'================================================================================


