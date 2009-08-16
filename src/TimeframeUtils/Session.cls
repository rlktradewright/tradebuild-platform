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
                ByRef ev As SessionEvent)

Event sessionEnded( _
                ByRef ev As SessionEvent)

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

Public Type SessionEvent
    source              As Object
End Type

'@================================================================================
' Member variables
'@================================================================================

' All times are in this timezone
Private mTimezone As TimeZone

Private mSessionStartTime As Date
Private mSessionEndTime As Date

Private mReferenceTime As Date

Private mCurrentSessionStartTime As Date
Private mCurrentSessionEndTime As Date

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
CurrentSessionEndTime = mCurrentSessionEndTime
End Property

Public Property Get CurrentSessionStartTime() As Date
CurrentSessionStartTime = mCurrentSessionStartTime
End Property

Friend Property Let IsLinkable( _
                ByVal value As Boolean)
mIsLinkable = value
End Property
                
Public Property Get IsLinkable() As Boolean
IsLinkable = mIsLinkable
End Property

Public Property Get SessionCurrentTime() As Date
SessionCurrentTime = mClock.Timestamp
End Property

Friend Property Let SessionEndTime( _
                ByVal value As Date)
mSessionEndTime = gNormaliseTime(value)
End Property

Public Property Get SessionEndTime() As Date
SessionEndTime = mSessionEndTime
End Property

Friend Property Let SessionStartTime( _
                ByVal value As Date)
mSessionStartTime = gNormaliseTime(value)
End Property
                
Public Property Get SessionStartTime() As Date
SessionStartTime = mSessionStartTime
End Property

Friend Property Let TimeZone( _
                ByVal value As TimeZone)
Set mTimezone = value
Set mClock = CreateSimulatedClock(1, mTimezone.standardName)
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
BarEndTime = gBarEndTime(Timestamp, BarTimePeriod, mSessionStartTime, mSessionStartTime)
End Function

Public Function BarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod) As Date
BarStartTime = gBarStartTime(Timestamp, BarTimePeriod, mSessionStartTime)
End Function

Public Function GetSessionTimes( _
                            ByVal Timestamp As Date) As SessionTimes
Dim st As SessionTimes
gCalcSessionTimes Timestamp, mSessionStartTime, mSessionEndTime, st.startTime, st.endTime
GetSessionTimes = st
End Function

Public Function GetOffsetSessionTimes( _
                ByVal Timestamp As Date, _
                ByVal offset As Long) As SessionTimes
Dim st As SessionTimes
gCalcOffsetSessionTimes Timestamp, _
                        offset, _
                        mSessionStartTime, _
                        mSessionEndTime, _
                        st.startTime, _
                        st.endTime
GetOffsetSessionTimes = st
End Function

Public Function IsTimeInSession(ByVal Timestamp As Date) As Boolean

If Timestamp >= mCurrentSessionStartTime And _
    Timestamp < mCurrentSessionEndTime _
Then
    IsTimeInSession = True
End If
End Function

Public Function NumberOfBarsInSession( _
                ByVal BarTimePeriod As TimePeriod) As Long
NumberOfBarsInSession = gCalcNumberOfBarsInSession(BarTimePeriod, mSessionStartTime, mSessionEndTime)
End Function


Public Function OffsetBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long) As Date
OffsetBarStartTime = gCalcOffsetBarStartTime( _
                                    Timestamp, _
                                    BarTimePeriod, _
                                    offset, _
                                    mSessionStartTime, _
                                    mSessionEndTime)
End Function

Friend Sub SetSessionCurrentTime( _
                ByVal Timestamp As Date)
Dim ev As SessionEvent
Static sessionStartNotified As Boolean
mClock.SetTime Timestamp
If Timestamp >= mCurrentSessionEndTime Then
    Set ev.source = Me
    RaiseEvent sessionEnded(ev)
    sessionStartNotified = False
End If

If Timestamp >= mCurrentSessionEndTime Or _
    Timestamp < mReferenceTime _
Then
    mReferenceTime = Timestamp
    gCalcSessionTimes Timestamp, _
                        mSessionStartTime, _
                        mSessionEndTime, _
                        mCurrentSessionStartTime, _
                        mCurrentSessionEndTime
End If

If Timestamp >= mCurrentSessionStartTime And _
    Not sessionStartNotified _
Then
    Set ev.source = Me
    RaiseEvent SessionStarted(ev)
    sessionStartNotified = True
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


