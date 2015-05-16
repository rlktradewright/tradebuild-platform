VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
' @see
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

Event SessionDateChanged( _
                ByRef ev As SessionEventData)

Event SessionStarted( _
                ByRef ev As SessionEventData)

Event SessionEnded( _
                ByRef ev As SessionEventData)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "Session"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

' All times are in this timezone
Private mTimezone                                   As TimeZone

Private mSessionStartTime                           As Date
Private mSessionEndTime                             As Date

Private mCurrentSessionTimes                        As SessionTimes

'Private mIsLinkable                                 As Boolean

'Private mClock                                      As Clock

Private mBuilderRef                                 As WeakReference

Private mTimestamp                                  As Date

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mTimezone = GetTimeZone("")
'Set mClock = CreateSimulatedClock(1, "")
mCurrentSessionTimes.EndTime = MinDate
mCurrentSessionTimes.StartTime = MinDate
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get CurrentSessionEndTime() As Date
Const ProcName As String = "CurrentSessionEndTime"
On Error GoTo Err

CurrentSessionEndTime = mCurrentSessionTimes.EndTime

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentSessionStartTime() As Date
Const ProcName As String = "CurrentSessionStartTime"
On Error GoTo Err

CurrentSessionStartTime = mCurrentSessionTimes.StartTime

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'Friend Property Let IsLinkable( _
'                ByVal Value As Boolean)
'mIsLinkable = Value
'End Property
                
'Public Property Get IsLinkable() As Boolean
'IsLinkable = mIsLinkable
'End Property

Friend Property Get SessionBuilder() As SessionBuilder
Set SessionBuilder = mBuilderRef.Target
End Property

Public Property Get SessionCurrentTime() As Date
Const ProcName As String = "SessionCurrentTime"
On Error GoTo Err

'SessionCurrentTime = mClock.Timestamp
SessionCurrentTime = mTimestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let SessionEndTime( _
                ByVal Value As Date)
Const ProcName As String = "SessionEndTime"
On Error GoTo Err

mSessionEndTime = gNormaliseTime(Value)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SessionEndTime() As Date
SessionEndTime = mSessionEndTime
End Property

Friend Property Let SessionStartTime( _
                ByVal Value As Date)
Const ProcName As String = "SessionStartTime"
On Error GoTo Err

mSessionStartTime = gNormaliseTime(Value)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property
                
Public Property Get SessionStartTime() As Date
SessionStartTime = mSessionStartTime
End Property

Friend Property Let TimeZone( _
                ByVal Value As TimeZone)
Const ProcName As String = "TimeZone"
On Error GoTo Err

If Value Is Nothing Then
    Set mTimezone = GetTimeZone("")
Else
    Set mTimezone = Value
End If

'Set mClock = CreateSimulatedClock(1, mTimezone.StandardName)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TimeZone() As TimeZone
Set TimeZone = mTimezone
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function GetSessionTimes( _
                            ByVal Timestamp As Date) As SessionTimes
Const ProcName As String = "GetSessionTimes"
On Error GoTo Err

GetSessionTimes = gCalcSessionTimes(Timestamp, mSessionStartTime, mSessionEndTime)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetOffsetSessionTimes( _
                ByVal Timestamp As Date, _
                ByVal offset As Long) As SessionTimes
Const ProcName As String = "GetOffsetSessionTimes"
On Error GoTo Err

GetOffsetSessionTimes = gCalcOffsetSessionTimes(Timestamp, offset, mSessionStartTime, mSessionEndTime)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Initialise(ByVal pSessionBuilder As SessionBuilder)
Set mBuilderRef = CreateWeakReference(pSessionBuilder)
End Sub

Public Function IsTimeInFirstSessionForWeek(ByVal pTimestamp As Date) As Boolean
Const ProcName As String = "IsTimeInFirstSessionForWeek"
On Error GoTo Err

Dim lSessionTimes As SessionTimes
lSessionTimes = GetSessionTimes(pTimestamp)
If TimeValue(lSessionTimes.StartTime) >= CDate("12:00") Then
    IsTimeInFirstSessionForWeek = (Weekday(lSessionTimes.StartTime, vbSunday) = 1)
Else
    IsTimeInFirstSessionForWeek = (Weekday(lSessionTimes.StartTime, vbMonday) = 1)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsTimeInSession(ByVal Timestamp As Date) As Boolean
IsTimeInSession = Timestamp >= mCurrentSessionTimes.StartTime And Timestamp < mCurrentSessionTimes.EndTime

'If Timestamp < mCurrentSessionTimes.StartTime Then
'ElseIf Timestamp >= mCurrentSessionTimes.EndTime Then
'Else
'    IsTimeInSession = True
'End If
End Function

''
' Links this <code>Session</code> object to another <code>Session</code>
' object so that both will have identical properties.
'
' This method succeeds only if this <code>Session</code> object has
' its <code>IsLinkable</code> property set to <code>True</code>, which can
' only be done using its <code>SessionBuilder</code> object's <code>SetLinkable</code>
' method.
'
' @param pSessionToLinkTo the <code>Session</code> object which this
'               <code>Session</code> object is to be linked to.
' @see SetLinkable, UnLinkSession
'
'@/
Public Sub LinkToSession( _
                ByVal pSessionToLinkTo As Session)
Const ProcName As String = "LinkSession"
On Error GoTo Err

'Assert IsLinkable, "Session is not linkable"

pSessionToLinkTo.SessionBuilder.LinkSession Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function SetSessionCurrentTime( _
                ByVal pTimestamp As Date) As SessionEventData
Const ProcName As String = "SetSessionCurrentTime"
On Error GoTo Err

Static sInitialised As Boolean
Static sSessionEndNotified As Boolean
Static sNextSessionTimes As SessionTimes
Static sPrevTimeNotified As Date

Dim ev As SessionEventData
Set ev.Source = Me
ev.changeType = SessionChangeNone

'mClock.SetTime pTimestamp
mTimestamp = pTimestamp

If Int(pTimestamp) > Int(sPrevTimeNotified) Then
    ev.changeType = SessionChangeDateChange
    ev.Timestamp = Int(pTimestamp)
    RaiseEvent SessionDateChanged(ev)
End If

If Not sInitialised Then
    sInitialised = True
    mCurrentSessionTimes = gCalcSessionTimes(pTimestamp, mSessionStartTime, mSessionEndTime)
    sNextSessionTimes = gCalcOffsetSessionTimes(pTimestamp, 1, mSessionStartTime, mSessionEndTime)
    ev.changeType = SessionChangeStart
    ev.Timestamp = mCurrentSessionTimes.StartTime
    RaiseEvent SessionStarted(ev)
End If

If pTimestamp >= mCurrentSessionTimes.EndTime And Not sSessionEndNotified Then
    sSessionEndNotified = True
    ev.changeType = SessionChangeEnd
    ev.Timestamp = mCurrentSessionTimes.EndTime
    RaiseEvent SessionEnded(ev)
    sSessionEndNotified = True
End If

If pTimestamp >= sNextSessionTimes.StartTime Then
    mCurrentSessionTimes = gCalcSessionTimes(pTimestamp, mSessionStartTime, mSessionEndTime)
    sNextSessionTimes = gCalcOffsetSessionTimes(pTimestamp, 1, mSessionStartTime, mSessionEndTime)
    sSessionEndNotified = False
    ev.changeType = SessionChangeStart
    ev.Timestamp = mCurrentSessionTimes.StartTime
    RaiseEvent SessionStarted(ev)
End If

SetSessionCurrentTime = ev

sPrevTimeNotified = pTimestamp

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

