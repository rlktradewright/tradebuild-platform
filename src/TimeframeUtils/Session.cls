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

Public Property Get currentSessionEndTime() As Date
currentSessionEndTime = mCurrentSessionEndTime
End Property

Public Property Get currentSessionStartTime() As Date
currentSessionStartTime = mCurrentSessionStartTime
End Property

Friend Property Let isLinkable( _
                ByVal value As Boolean)
mIsLinkable = value
End Property
                
Public Property Get isLinkable() As Boolean
isLinkable = mIsLinkable
End Property

Public Property Get sessionCurrentTime() As Date
sessionCurrentTime = mClock.timestamp
End Property

Friend Property Let sessionEndTime( _
                ByVal value As Date)
mSessionEndTime = gNormaliseTime(value)
End Property

Public Property Get sessionEndTime() As Date
sessionEndTime = mSessionEndTime
End Property

Friend Property Let sessionStartTime( _
                ByVal value As Date)
mSessionStartTime = gNormaliseTime(value)
End Property
                
Public Property Get sessionStartTime() As Date
sessionStartTime = mSessionStartTime
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
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal units As TimePeriodUnits) As Date
BarEndTime = gBarEndTime(timestamp, barLength, units, mSessionStartTime, mSessionStartTime)
End Function

Public Function BarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal units As TimePeriodUnits) As Date
BarStartTime = gBarStartTime(timestamp, barLength, units, mSessionStartTime)
End Function

Public Function isTimeInSession(ByVal timestamp As Date) As Boolean

If timestamp >= mCurrentSessionStartTime And _
    timestamp < mCurrentSessionEndTime _
Then
    isTimeInSession = True
End If
End Function

Public Function NumberOfBarsInSession( _
                ByVal barLength As Long, _
                ByVal units As TimePeriodUnits) As Long
NumberOfBarsInSession = gCalcNumberOfBarsInSession(barLength, units, mSessionStartTime, mSessionEndTime)
End Function


Public Function OffsetBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal units As TimePeriodUnits, _
                ByVal offset As Long) As Date
OffsetBarStartTime = gCalcOffsetBarStartTime( _
                                    timestamp, _
                                    barLength, _
                                    units, _
                                    offset, _
                                    mSessionStartTime, _
                                    mSessionEndTime)
End Function

Public Sub SessionTimes( _
                            ByVal timestamp As Date, _
                            ByRef sessionStartTime As Date, _
                            ByRef sessionEndTime As Date)
gCalcSessionTimes timestamp, mSessionStartTime, mSessionEndTime, sessionStartTime, sessionEndTime
End Sub

Friend Sub setSessionCurrentTime( _
                ByVal timestamp As Date)
Dim ev As SessionEvent
Static sessionStartNotified As Boolean
mClock.SetTime timestamp
If timestamp >= mCurrentSessionEndTime Then
    Set ev.source = Me
    RaiseEvent sessionEnded(ev)
    sessionStartNotified = False
End If

If timestamp >= mCurrentSessionEndTime Or _
    timestamp < mReferenceTime _
Then
    mReferenceTime = timestamp
    gCalcSessionTimes timestamp, _
                        mSessionStartTime, _
                        mSessionEndTime, _
                        mCurrentSessionStartTime, _
                        mCurrentSessionEndTime
End If

If timestamp >= mCurrentSessionStartTime And _
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


