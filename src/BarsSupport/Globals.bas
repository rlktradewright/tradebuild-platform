Attribute VB_Name = "Globals"
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

'================================================================================
' Constants
'================================================================================

Public Const MaxDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023
Public Const MinDouble As Double = -(2 - 2 ^ -52) * 2 ^ 1023

Public Const DummyHigh As Double = MinDouble
Public Const DummyLow As Double = MaxDouble

Public Const DefaultStudyValueNameStr As String = "$DEFAULT"
Public Const MovingAverageStudyValueNameStr As String = "MA"

Public Const OneMicroSecond As Double = 1.15740740740741E-11
Public Const OneSecond As Double = 1 / 86400
Public Const OneMinute As Double = 1 / 1440
Public Const OneHour As Double = 1 / 24


'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' External function declarations
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mStudyLibraryManager      As New StudyLibraryManager

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

Public Property Get StudyLibraryManager() As StudyLibraryManager
Set StudyLibraryManager = mStudyLibraryManager
End Property

'================================================================================
' Methods
'================================================================================

Public Function gBarEndTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal timeUnits As TimePeriodUnits, _
                Optional ByVal sessionStartTime As Date) As Date
Dim startTime As Date
startTime = gBarStartTime( _
                timestamp, _
                barLength, _
                timeUnits, _
                sessionStartTime)
Select Case timeUnits
Case TimePeriodSecond
    gBarEndTime = startTime + (barLength / 86400) + OneMicroSecond
Case TimePeriodMinute
    gBarEndTime = startTime + (barLength / 1440) + OneMicroSecond
Case TimePeriodHour
    gBarEndTime = startTime + (barLength / 24) + OneMicroSecond
Case TimePeriodDay
    gBarEndTime = startTime + barLength
Case TimePeriodWeek
    gBarEndTime = startTime + 7 * barLength
Case TimePeriodMonth
    gBarEndTime = DateAdd("m", barLength, startTime)
'Case TimePeriodLunarMonth
'
Case TimePeriodYear
    gBarEndTime = DateAdd("yyyy", barLength, startTime)
End Select
End Function

Public Function gBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal timeUnits As TimePeriodUnits, _
                ByVal sessionStartTime As Date) As Date
' minutes from midnight to start of sesssion
Dim sessionOffset              As Long
Dim theDate As Long
Dim theTime As Double
Dim theTimeMins As Long
Dim theTimeSecs As Long

sessionOffset = Int(1440 * (sessionStartTime + OneMicroSecond - Int(sessionStartTime)))

theDate = Int(CDbl(timestamp))
' NB: don't use TimeValue to get the time, as VB rounds it to
' the nearest second
theTime = CDbl(timestamp + OneMicroSecond) - theDate

Select Case timeUnits
Case TimePeriodSecond
    theTimeSecs = Fix(theTime * 86400) ' seconds since midnight
    gBarStartTime = theDate + _
                ((barLength) * Int((theTimeSecs - sessionOffset * 60) / barLength) + _
                    sessionOffset * 60) / 86400
Case TimePeriodMinute
    theTimeMins = Fix(theTime * 1440) ' minutes since midnight
    gBarStartTime = theDate + _
                (barLength * Int((theTimeMins - sessionOffset) / barLength) + _
                    sessionOffset) / 1440
Case TimePeriodHour
    theTimeMins = Fix(theTime * 1440) ' minutes since midnight
    gBarStartTime = theDate + _
                (60 * barLength * Int((theTimeMins - sessionOffset) / (60 * barLength)) + _
                    sessionOffset) / 1440
Case TimePeriodDay
    If theTime >= sessionStartTime Then
        gBarStartTime = theDate
    Else
        gBarStartTime = theDate - 1
    End If
Case TimePeriodWeek
    gBarStartTime = theDate - DatePart("w", theDate, vbSunday) + 1
Case TimePeriodMonth
    gBarStartTime = DateSerial(Year(theDate), Month(theDate), 1)
'Case TimePeriodLunarMonth
'
Case TimePeriodYear
    gBarStartTime = DateSerial(Year(theDate), 1, 1)
End Select

End Function

Public Function gCalcBarLength( _
                ByVal length As Long, _
                ByVal units As TWUtilities.TimePeriodUnits) As Date
Select Case units
Case TimePeriodSecond
    gCalcBarLength = length * OneSecond
Case TimePeriodMinute
    gCalcBarLength = length * OneMinute
Case TimePeriodHour
    gCalcBarLength = length * OneHour
Case TimePeriodDay
    gCalcBarLength = length
End Select
End Function

Public Sub gCalcSessionTimes( _
                ByVal timestamp As Date, _
                ByVal startTime As Date, _
                ByVal endTime As Date, _
                ByRef sessionStartTime As Date, _
                ByRef sessionEndTime As Date)
Dim i As Long

i = -1
Do
    i = i + 1
Loop Until calcSessionTimesHelper(timestamp + i, _
                            startTime, _
                            endTime, _
                            sessionStartTime, _
                            sessionEndTime)
End Sub

Public Function gOffsetBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal timeUnits As TimePeriodUnits, _
                ByVal offset As Long, _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date) As Date
Dim sessStart As Date
Dim sessEnd As Date
Dim datumBarStart As Date
Dim proposedStart As Date
Dim remainingOffset As Long
Dim barsFromSessStart As Long
Dim i As Long

datumBarStart = gBarStartTime(timestamp, barLength, timeUnits, sessionStartTime)
' !!!!!!!!!!!!!!!!!!!!!!!!!!!
' calcSessionTimes needs to be amended
' to take account of holidays
gCalcSessionTimes datumBarStart, _
                    sessionStartTime, _
                    sessionEndTime, _
                    sessStart, _
                    sessEnd

If datumBarStart < sessStart Then
    ' specified timestamp was between sessions
    datumBarStart = sessStart
End If
    
remainingOffset = offset
proposedStart = datumBarStart
Do While remainingOffset > 0
    barsFromSessStart = (proposedStart - sessStart) / (barLength / 1440)
    If barsFromSessStart >= remainingOffset Then
        proposedStart = proposedStart - (remainingOffset * barLength) / 1440
        remainingOffset = 0
    Else
        remainingOffset = remainingOffset - barsFromSessStart
        Do
            i = i + 1
            gCalcSessionTimes proposedStart - i, _
                                sessionStartTime, _
                                sessionEndTime, _
                                sessStart, _
                                sessEnd
        Loop Until sessStart <= proposedStart
        proposedStart = sessEnd
    End If
Loop
gOffsetBarStartTime = proposedStart
                
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Function calcSessionTimesHelper(ByVal timestamp As Date, _
                            ByVal startTime As Date, _
                            ByVal endTime As Date, _
                            ByRef sessionStartTime As Date, _
                            ByRef sessionEndTime As Date) As Boolean
Dim referenceDate As Date
Dim referenceTime As Date
Dim weekday As Long

referenceDate = DateValue(timestamp)
referenceTime = TimeValue(timestamp)

If startTime < endTime Then
    ' session doesn't span midnight
    If referenceTime < endTime Then
        sessionStartTime = referenceDate + startTime
        sessionEndTime = referenceDate + endTime
    Else
        sessionStartTime = referenceDate + 1 + startTime
        sessionEndTime = referenceDate + 1 + endTime
    End If
ElseIf startTime > endTime Then
    ' session spans midnight
    If referenceTime >= endTime Then
        sessionStartTime = referenceDate + startTime
        sessionEndTime = referenceDate + 1 + endTime
    Else
        sessionStartTime = referenceDate - 1 + startTime
        sessionEndTime = referenceDate + endTime
    End If
Else
    ' this instrument trades 24hrs, or the contract service provider doesn't know
    ' the session start and end times
    sessionStartTime = referenceDate
    sessionEndTime = referenceDate + 1
End If

weekday = DatePart("w", sessionStartTime)
If startTime < endTime Then
    ' session doesn't span midnight
    If weekday <> vbSaturday And weekday <> vbSunday Then calcSessionTimesHelper = True
ElseIf startTime > endTime Then
    ' session DOES span midnight
    If weekday <> vbFriday And weekday <> vbSaturday Then calcSessionTimesHelper = True
Else
    ' 24-hour session or no session times known
    If weekday <> vbSaturday And weekday <> vbSunday Then calcSessionTimesHelper = True
End If
End Function



