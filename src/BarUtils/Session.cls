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

Event SessionStarted( _
                ByRef ev As SessionEventData)

Event SessionEnded( _
                ByRef ev As SessionEventData)

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName As String = "TimeframeUtils25"
Private Const ModuleName As String = "SessionBuilder"

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
Private mTimezone As TimeZone

Private mSessionStartTime As Date
Private mSessionEndTime As Date

Private mCurrentSessionTimes As SessionTimes

Private mIsLinkable As Boolean

Private mClock As Clock

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mClock = CreateSimulatedClock(1, "")
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
CurrentSessionEndTime = mCurrentSessionTimes.endTime
End Property

Public Property Get CurrentSessionStartTime() As Date
CurrentSessionStartTime = mCurrentSessionTimes.startTime
End Property

Friend Property Let IsLinkable( _
                ByVal value As Boolean)
mIsLinkable = value
End Property
                
Public Property Get IsLinkable() As Boolean
IsLinkable = mIsLinkable
End Property

Public Property Get SessionCurrentTime() As Date
Const ProcName As String = "SessionCurrentTime"
Dim failpoint As String
On Error GoTo Err

SessionCurrentTime = mClock.Timestamp

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Friend Property Let SessionEndTime( _
                ByVal value As Date)
Const ProcName As String = "SessionEndTime"
Dim failpoint As String
On Error GoTo Err

mSessionEndTime = gNormaliseTime(value)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get SessionEndTime() As Date
SessionEndTime = mSessionEndTime
End Property

Friend Property Let SessionStartTime( _
                ByVal value As Date)
Const ProcName As String = "SessionStartTime"
Dim failpoint As String
On Error GoTo Err

mSessionStartTime = gNormaliseTime(value)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property
                
Public Property Get SessionStartTime() As Date
SessionStartTime = mSessionStartTime
End Property

Friend Property Let TimeZone( _
                ByVal value As TimeZone)
Const ProcName As String = "TimeZone"
Dim failpoint As String
On Error GoTo Err

Set mTimezone = value
Set mClock = CreateSimulatedClock(1, mTimezone.StandardName)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get TimeZone() As TimeZone
Set TimeZone = mTimezone
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function BarEndTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod) As Date
Const ProcName As String = "BarEndTime"
Dim failpoint As String
On Error GoTo Err

BarEndTime = gBarEndTime(Timestamp, BarTimePeriod, mSessionStartTime, mSessionStartTime)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function BarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod) As Date
Const ProcName As String = "BarStartTime"
Dim failpoint As String
On Error GoTo Err

BarStartTime = gBarStartTime(Timestamp, BarTimePeriod, mSessionStartTime)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function GetSessionTimes( _
                            ByVal Timestamp As Date) As SessionTimes
Const ProcName As String = "GetSessionTimes"
Dim failpoint As String
On Error GoTo Err

GetSessionTimes = gCalcSessionTimes(Timestamp, mSessionStartTime, mSessionEndTime)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function GetOffsetSessionTimes( _
                ByVal Timestamp As Date, _
                ByVal offset As Long) As SessionTimes
Const ProcName As String = "GetOffsetSessionTimes"
Dim failpoint As String
On Error GoTo Err

GetOffsetSessionTimes = gCalcOffsetSessionTimes(Timestamp, offset, mSessionStartTime, mSessionEndTime)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function IsTimeInFirstSessionForWeek(ByVal pTimestamp As Date) As Boolean
Dim lSessionTimes As SessionTimes
lSessionTimes = GetSessionTimes(pTimestamp)
If TimeValue(lSessionTimes.startTime) >= CDate("12:00") Then
    IsTimeInFirstSessionForWeek = (weekday(lSessionTimes.startTime, vbSunday) = 1)
Else
    IsTimeInFirstSessionForWeek = (weekday(lSessionTimes.startTime, vbMonday) = 1)
End If
End Function

Public Function IsTimeInSession(ByVal Timestamp As Date) As Boolean

Const ProcName As String = "IsTimeInSession"
Dim failpoint As String
On Error GoTo Err

If Timestamp >= mCurrentSessionTimes.startTime And _
    Timestamp < mCurrentSessionTimes.endTime _
Then
    IsTimeInSession = True
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function NumberOfBarsInSession( _
                ByVal BarTimePeriod As TimePeriod) As Long
Const ProcName As String = "NumberOfBarsInSession"
Dim failpoint As String
On Error GoTo Err

NumberOfBarsInSession = gCalcNumberOfBarsInSession(BarTimePeriod, mSessionStartTime, mSessionEndTime)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function


Public Function OffsetBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long) As Date
Const ProcName As String = "OffsetBarStartTime"
Dim failpoint As String
On Error GoTo Err

OffsetBarStartTime = gCalcOffsetBarStartTime( _
                                    Timestamp, _
                                    BarTimePeriod, _
                                    offset, _
                                    mSessionStartTime, _
                                    mSessionEndTime)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Friend Function SetSessionCurrentTime( _
                ByVal pTimestamp As Date) As SessionEventData
Const ProcName As String = "SetSessionCurrentTime"
On Error GoTo Err

Static lInitialised As Boolean
Static lSessionEndNotified As Boolean
Static lNextSessionTimes As SessionTimes

Dim ev As SessionEventData

Set ev.Source = Me
ev.changeType = SessionChangeNone

If Not lInitialised Then
    lInitialised = True
    mCurrentSessionTimes = gCalcSessionTimes(pTimestamp, mSessionStartTime, mSessionEndTime)
    lNextSessionTimes = gCalcOffsetSessionTimes(pTimestamp, 1, mSessionStartTime, mSessionEndTime)
    ev.changeType = SessionChangeStart
    ev.Timestamp = mCurrentSessionTimes.startTime
    RaiseEvent SessionStarted(ev)
End If

mClock.SetTime pTimestamp

If pTimestamp >= mCurrentSessionTimes.endTime And Not lSessionEndNotified Then
    lSessionEndNotified = True
    ev.changeType = SessionChangeEnd
    ev.Timestamp = mCurrentSessionTimes.endTime
    RaiseEvent SessionEnded(ev)
    lSessionEndNotified = True
End If

'If pTimestamp >= lNextSessionTimes.StartTime Or _
'    pTimestamp < mReferenceTime _
'Then
'    mReferenceTime = pTimestamp
'    mCurrentSessionTimes = gCalcSessionTimes(pTimestamp, mSessionStartTime, mSessionEndTime)
'End If

If pTimestamp >= lNextSessionTimes.startTime Then
    mCurrentSessionTimes = gCalcSessionTimes(pTimestamp, mSessionStartTime, mSessionEndTime)
    lNextSessionTimes = gCalcOffsetSessionTimes(pTimestamp, 1, mSessionStartTime, mSessionEndTime)
    lSessionEndNotified = False
    ev.changeType = SessionChangeStart
    ev.Timestamp = mCurrentSessionTimes.startTime
    RaiseEvent SessionStarted(ev)
End If

SetSessionCurrentTime = ev

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


