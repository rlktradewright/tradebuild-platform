Attribute VB_Name = "Globals"
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Public Const MaxDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023
Public Const MinDouble As Double = -(2 - 2 ^ -52) * 2 ^ 1023

Public Const DummyHigh As Double = MinDouble
Public Const DummyLow As Double = MaxDouble

Public Const OneMicroSecond As Double = 1.15740740740741E-11
Public Const OneSecond As Double = 1 / 86400
Public Const OneMinute As Double = 1 / 1440
Public Const OneHour As Double = 1 / 24

Public Const TimePeriodNameSecond As String = "Second"
Public Const TimePeriodNameMinute As String = "Minute"
Public Const TimePeriodNameHour As String = "Hourly"
Public Const TimePeriodNameDay As String = "Daily"
Public Const TimePeriodNameWeek As String = "Weekly"
Public Const TimePeriodNameMonth As String = "Monthly"
Public Const TimePeriodNameYear As String = "Yearly"

Public Const TimePeriodNameSeconds As String = "Seconds"
Public Const TimePeriodNameMinutes As String = "Minutes"
Public Const TimePeriodNameHours As String = "Hours"
Public Const TimePeriodNameDays As String = "Days"
Public Const TimePeriodNameWeeks As String = "Weeks"
Public Const TimePeriodNameMonths As String = "Months"
Public Const TimePeriodNameYears As String = "Years"
Public Const TimePeriodNameVolumeIncrement As String = "Volume"
Public Const TimePeriodNameTickVolumeIncrement As String = "Tick Volume"
Public Const TimePeriodNameTickIncrement As String = "Ticks Movement"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Function gBarEndTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal units As TimePeriodUnits, _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date) As Date
Dim startTime As Date
startTime = gBarStartTime( _
                timestamp, _
                barLength, _
                units, _
                sessionStartTime)
Select Case units
Case TimePeriodSecond
    gBarEndTime = startTime + (barLength / 86400) - OneMicroSecond
Case TimePeriodMinute
    gBarEndTime = startTime + (barLength / 1440) - OneMicroSecond
Case TimePeriodHour
    gBarEndTime = startTime + (barLength / 24) - OneMicroSecond
Case TimePeriodDay
    gBarEndTime = gCalcWorkingDayDate(gCalcWorkingDayNumber(startTime) + barLength, startTime)
Case TimePeriodWeek
    gBarEndTime = startTime + 7 * barLength
Case TimePeriodMonth
    gBarEndTime = DateAdd("m", barLength, startTime)
Case TimePeriodYear
    gBarEndTime = DateAdd("yyyy", barLength, startTime)
Case TimePeriodVolume, _
        TimePeriodTickVolume, _
        TimePeriodTickMovement
    gBarEndTime = timestamp
End Select

Dim sessStart As Date
Dim sessEnd As Date

calcSessionTimesHelper startTime, sessionStartTime, sessionEndTime, sessStart, sessEnd

If startTime < sessStart And gBarEndTime > sessStart Then
    gBarEndTime = Int(gBarEndTime) + sessionStartTime
End If
End Function

Public Function gBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal units As TimePeriodUnits, _
                ByVal sessionStartTime As Date) As Date

' minutes from midnight to start of sesssion
Dim sessionOffset           As Long
Dim theDate                 As Long
Dim theTime                 As Double
Dim theTimeMins             As Long
Dim theTimeSecs             As Long

sessionOffset = Int(1440 * (sessionStartTime + OneMicroSecond - Int(sessionStartTime)))

theDate = Int(CDbl(timestamp))
' NB: don't use TimeValue to get the time, as VB rounds it to
' the nearest second
theTime = CDbl(timestamp + OneMicroSecond) - theDate

Select Case units
Case TimePeriodSecond
    theTimeSecs = Fix(theTime * 86400) ' seconds since midnight
    If theTimeSecs < sessionOffset * 60 Then
        theDate = theDate - 1
        theTimeSecs = theTimeSecs + 86400
    End If
    gBarStartTime = theDate + _
                (barLength * Int((theTimeSecs - sessionOffset * 60) / barLength) + _
                    sessionOffset * 60) / 86400
Case TimePeriodMinute
    theTimeMins = Fix(theTime * 1440) ' minutes since midnight
    If theTimeMins < sessionOffset Then
        theDate = theDate - 1
        theTimeMins = theTimeMins + 1440
    End If
    gBarStartTime = theDate + _
                (barLength * Int((theTimeMins - sessionOffset) / barLength) + _
                    sessionOffset) / 1440
Case TimePeriodHour
    theTimeMins = Fix(theTime * 1440) ' minutes since midnight
    If theTimeMins < sessionOffset Then
        theDate = theDate - 1
        theTimeMins = theTimeMins + 1440
    End If
    gBarStartTime = theDate + _
                (60 * barLength * Int((theTimeMins - sessionOffset) / (60 * barLength)) + _
                    sessionOffset) / 1440
Case TimePeriodDay
    Dim workingDayNum As Long
    If theTime < sessionStartTime Then
        theDate = theDate - 1
    End If
    
    If barLength = 1 Then
        gBarStartTime = theDate + sessionStartTime
    Else
        workingDayNum = gCalcWorkingDayNumber(theDate)
        
        gBarStartTime = gCalcWorkingDayDate(1 + barLength * Int((workingDayNum - 1) / barLength), _
                                        theDate) + sessionStartTime
    End If
Case TimePeriodWeek
    Dim weekNum As Long
    
    weekNum = DatePart("ww", theDate, vbMonday, vbFirstFullWeek)
    If weekNum >= 52 And Month(theDate) = 1 Then
        ' this must be part of the final week of the previous year
        theDate = DateAdd("yyyy", -1, theDate)
    End If
    gBarStartTime = gCalcWeekStartDate(1 + barLength * Int((weekNum - 1) / barLength), _
                                    theDate) + sessionStartTime

Case TimePeriodMonth
    Dim monthNum As Long
    
    monthNum = Month(theDate)
    gBarStartTime = gCalcMonthStartDate(1 + barLength * Int((monthNum - 1) / barLength), _
                                    theDate) + sessionStartTime
Case TimePeriodYear
    gBarStartTime = DateSerial(1900 + barLength * Int((Year(theDate) - 1900) / barLength), 1, 1)
Case TimePeriodVolume, _
        TimePeriodTickVolume, _
        TimePeriodTickMovement
    gBarStartTime = timestamp
End Select

End Function

Public Function gCalcBarLength( _
                ByVal length As Long, _
                ByVal units As TimePeriodUnits) As Date
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

Public Function gCalcMonthStartDate( _
                ByVal monthNumber As Long, _
                ByVal baseDate As Date) As Date
Dim yearStart As Date

yearStart = DateAdd("d", 1 - DatePart("y", baseDate), baseDate)
gCalcMonthStartDate = DateAdd("m", monthNumber - 1, yearStart)
End Function


Public Function gCalcNumberOfBarsInSession( _
                ByVal barLength As Long, _
                ByVal units As TimePeriodUnits, _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date) As Long

If sessionEndTime > sessionStartTime Then
    gCalcNumberOfBarsInSession = -Int(-(sessionEndTime - sessionStartTime) / gCalcBarLength(barLength, units))
Else
    gCalcNumberOfBarsInSession = -Int(-(1 + sessionEndTime - sessionStartTime) / gCalcBarLength(barLength, units))
End If
End Function

Public Function gCalcOffsetBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal units As TimePeriodUnits, _
                ByVal offset As Long, _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date) As Date
Dim sessStart As Date
Dim sessEnd As Date
Dim datumBarStart As Date
Dim proposedStart As Date
Dim remainingOffset As Long
Dim barsFromSessStart As Long
Dim barsToSessEnd As Long
Dim i As Long
Dim barLengthDays As Date
Dim numBarsInSession As Long

Select Case units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
    gCalcOffsetBarStartTime = calcOffsetDailyBarStartTime( _
                                                    timestamp, _
                                                    barLength, _
                                                    offset, _
                                                    sessionStartTime)
    Exit Function
Case TimePeriodWeek
    gCalcOffsetBarStartTime = calcOffsetWeeklyBarStartTime( _
                                                    timestamp, _
                                                    barLength, _
                                                    offset, _
                                                    sessionStartTime)
    Exit Function
Case TimePeriodMonth
    gCalcOffsetBarStartTime = calcOffsetMonthlyBarStartTime( _
                                                    timestamp, _
                                                    barLength, _
                                                    offset, _
                                                    sessionStartTime)
    Exit Function
Case TimePeriodYear
    gCalcOffsetBarStartTime = calcOffsetYearlyBarStartTime( _
                                                    timestamp, _
                                                    barLength, _
                                                    offset, _
                                                    sessionStartTime)
    Exit Function
End Select

barLengthDays = gCalcBarLength(barLength, units) + OneMicroSecond

numBarsInSession = gCalcNumberOfBarsInSession(barLength, units, sessionStartTime, sessionEndTime)

datumBarStart = gBarStartTime(timestamp, barLength, units, sessionStartTime)

gCalcSessionTimes timestamp, _
                    sessionStartTime, _
                    sessionEndTime, _
                    sessStart, _
                    sessEnd

If offset > 0 Then
    
    If datumBarStart < sessStart Then
        ' specified timestamp was between sessions
        datumBarStart = sessStart
        offset = offset - 1
    End If
    
    barsToSessEnd = Round((sessEnd - datumBarStart) / barLengthDays, 0)
    If barsToSessEnd >= offset Then
        ' all required bars are in this session
        proposedStart = datumBarStart + offset * barLengthDays
    Else
        remainingOffset = offset - barsToSessEnd
        proposedStart = datumBarStart + barsToSessEnd * barLengthDays
        Do While remainingOffset > 0
            gCalcSessionTimes proposedStart, _
                                sessionStartTime, _
                                sessionEndTime, _
                                sessStart, _
                                sessEnd
            If numBarsInSession >= remainingOffset Then
                proposedStart = sessStart + remainingOffset * barLengthDays
                remainingOffset = 0
            Else
                proposedStart = sessEnd
                remainingOffset = remainingOffset - numBarsInSession
            End If
        Loop
    End If
    If proposedStart >= sessEnd Then
        gCalcSessionTimes proposedStart, sessionStartTime, sessionEndTime, sessStart, sessEnd
        proposedStart = sessStart
    End If
Else
    offset = -offset

    If datumBarStart < sessStart Then
        ' specified timestamp was between sessions
        datumBarStart = sessStart
    End If
    
    proposedStart = sessStart
    barsFromSessStart = Round((datumBarStart - sessStart) / barLengthDays, 0)
    If barsFromSessStart >= offset Then
        ' all required bars are in this session
        'proposedStart = datumBarStart - offset * barLengthDays
        proposedStart = sessStart + (barsFromSessStart - offset) * barLengthDays
    Else
        remainingOffset = offset - barsFromSessStart
        proposedStart = sessStart
        Do While remainingOffset > 0
            
            ' find the previous session start - may be a weekend in the
            ' way
            i = 0
            Do
                i = i + 1
                gCalcSessionTimes proposedStart - i, _
                                    sessionStartTime, _
                                    sessionEndTime, _
                                    sessStart, _
                                    sessEnd
            Loop Until sessStart < proposedStart
            
            If numBarsInSession >= remainingOffset Then
                proposedStart = sessStart + (numBarsInSession - remainingOffset) * barLengthDays
                remainingOffset = 0
            Else
                proposedStart = sessStart
                remainingOffset = remainingOffset - numBarsInSession
            End If
        Loop
    End If
End If

gCalcOffsetBarStartTime = gBarStartTime(proposedStart, _
                                        barLength, _
                                        units, _
                                        sessionStartTime)


End Function

Public Sub gCalcOffsetSessionTimes( _
                ByVal timestamp As Date, _
                ByVal offset As Long, _
                ByVal startTime As Date, _
                ByVal endTime As Date, _
                ByRef sessionStartTime As Date, _
                ByRef sessionEndTime As Date)
Dim datumSessionStart As Date
Dim datumSessionEnd As Date
Dim targetWorkingDayNum As Long

gCalcSessionTimes timestamp, startTime, endTime, datumSessionStart, datumSessionEnd

targetWorkingDayNum = gCalcWorkingDayNumber(datumSessionStart) + offset

gCalcSessionTimes gCalcWorkingDayDate(targetWorkingDayNum, datumSessionStart), _
                    startTime, _
                    endTime, _
                    sessionStartTime, _
                    sessionEndTime
                
End Sub

' !!!!!!!!!!!!!!!!!!!!!!!!!!!
' calcSessionTimes needs to be amended
' to take account of holidays
Public Sub gCalcSessionTimes( _
                ByVal timestamp As Date, _
                ByVal startTime As Date, _
                ByVal endTime As Date, _
                ByRef sessionStartTime As Date, _
                ByRef sessionEndTime As Date)
Dim weekday As Long

calcSessionTimesHelper timestamp, _
                        startTime, _
                        endTime, _
                        sessionStartTime, _
                        sessionEndTime

weekday = DatePart("w", sessionStartTime)
If startTime > endTime Then
    ' session DOES span midnight
    If weekday = vbFriday Then
        sessionStartTime = sessionStartTime + 2
        sessionEndTime = sessionEndTime + 2
    ElseIf weekday = vbSaturday Then
        sessionStartTime = sessionStartTime + 1
        sessionEndTime = sessionEndTime + 1
    End If
Else
    ' session doesn't span midnight or 24-hour session or no session times known
    If weekday = vbSaturday Then
        sessionStartTime = sessionStartTime + 2
        sessionEndTime = sessionEndTime + 2
    ElseIf weekday = vbSunday Then
        sessionStartTime = sessionStartTime + 1
        sessionEndTime = sessionEndTime + 1
    End If
End If

End Sub

Public Function gCalcWeekStartDate( _
                ByVal weekNumber As Long, _
                ByVal baseDate As Date) As Date
Dim yearStart As Date
Dim week1Date As Date
Dim dow1 As Long    ' day of week of 1st jan of base year

yearStart = DateAdd("d", 1 - DatePart("y", baseDate), baseDate)

dow1 = DatePart("w", yearStart, vbMonday)

If dow1 = 1 Then
    week1Date = yearStart
Else
    week1Date = DateAdd("d", 8 - dow1, yearStart)
End If

gCalcWeekStartDate = DateAdd("ww", weekNumber - 1, week1Date)
End Function


Public Function gCalcWorkingDayDate( _
                ByVal dayNumber As Long, _
                ByVal baseDate As Date) As Date
Dim yearStart As Date
Dim yearEnd As Date
Dim doy As Long

Dim wd1 As Long     ' weekdays in first week (excluding weekend)
Dim we1 As Long     ' weekend days at start of first week

Dim dow1 As Long    ' day of week of 1st jan of base year

' number of whole weeks after the first week
Dim numWholeWeeks As Long

yearStart = DateAdd("d", 1 - DatePart("y", baseDate), baseDate)

Do While dayNumber < 0
    yearEnd = yearStart - 1
    yearStart = DateAdd("yyyy", -1, yearStart)
    dayNumber = dayNumber + gCalcWorkingDayNumber(yearEnd) + 1
Loop

dow1 = DatePart("w", yearStart, vbMonday)

If dow1 = 7 Then
    ' Sunday
    wd1 = 0
    we1 = 1
ElseIf dow1 = 6 Then
    ' Saturday
    wd1 = 0
    we1 = 2
Else
    wd1 = 5 - dow1 + 1
    we1 = 2
End If

If dayNumber <= wd1 Then
    doy = dayNumber
ElseIf dayNumber - wd1 <= 5 Then
    doy = we1 + dayNumber
Else
    numWholeWeeks = Int((dayNumber - wd1) / 5) - 1
    doy = wd1 + we1 + IIf(numWholeWeeks > 0, 7 * numWholeWeeks + 5, 5) + IIf(((dayNumber - wd1) Mod 5) > 0, ((dayNumber - wd1) Mod 5) + 2, 0)
End If

gCalcWorkingDayDate = DateAdd("d", doy - 1, yearStart)
End Function



Public Function gCalcWorkingDayNumber( _
                ByVal pDate As Date) As Long
Dim doy As Long     ' day of year
Dim woy As Long     ' week of year
Dim wd1 As Long     ' weekdays in first week (excluding weekend)
'Dim we1 As Long     ' weekend days at start of first week
Dim wdN As Long     ' weekdays in last week

Dim dow1 As Long    ' day of week of 1st jan
Dim dow As Long     ' day of week of supplied date

doy = DatePart("y", pDate, vbMonday)
woy = DatePart("ww", pDate, vbMonday)
dow = DatePart("w", pDate, vbMonday)
dow1 = DatePart("w", pDate - doy + 1, vbMonday)

If dow1 = 7 Then
    ' Sunday
    wd1 = 0
'    we1 = 1
ElseIf dow1 = 6 Then
    ' Saturday
    wd1 = 0
'    we1 = 2
Else
    wd1 = 5 - dow1 + 1
'    we1 = 0
End If

If dow = 7 Or dow = 6 Then
    wdN = 5
Else
    wdN = dow
End If

gCalcWorkingDayNumber = wd1 + 5 * (woy - 2) + wdN

End Function

Public Function gNormaliseTime( _
            ByVal timestamp As Date) As Date
gNormaliseTime = timestamp - Int(timestamp)
End Function

Public Function gTimePeriodUnitsFromString( _
                timeUnits As String) As TimePeriodUnits

Select Case UCase$(timeUnits)
Case UCase$(TimePeriodNameSecond), UCase$(TimePeriodNameSeconds), "SEC", "SECS", "S"
    gTimePeriodUnitsFromString = TimePeriodSecond
Case UCase$(TimePeriodNameMinute), UCase$(TimePeriodNameMinutes), "MIN", "MINS", "M"
    gTimePeriodUnitsFromString = TimePeriodMinute
Case UCase$(TimePeriodNameHour), UCase$(TimePeriodNameHours), "HR", "HRS", "H"
    gTimePeriodUnitsFromString = TimePeriodHour
Case UCase$(TimePeriodNameDay), UCase$(TimePeriodNameDays), "D", "DY", "DYS"
    gTimePeriodUnitsFromString = TimePeriodDay
Case UCase$(TimePeriodNameWeek), UCase$(TimePeriodNameWeeks), "W", "WK", "WKS"
    gTimePeriodUnitsFromString = TimePeriodWeek
Case UCase$(TimePeriodNameMonth), UCase$(TimePeriodNameMonths), "MTH", "MNTH", "MTHS", "MNTHS", "MM"
    gTimePeriodUnitsFromString = TimePeriodMonth
Case UCase$(TimePeriodNameYear), UCase$(TimePeriodNameYears), "YR", "YRS", "Y", "YY", "YS"
    gTimePeriodUnitsFromString = TimePeriodYear
Case UCase$(TimePeriodNameVolumeIncrement), "VOL", "V"
    gTimePeriodUnitsFromString = TimePeriodVolume
Case UCase$(TimePeriodNameTickVolumeIncrement), "TICKVOL", "TICK VOL", "TICKVOLUME", "TV"
    gTimePeriodUnitsFromString = TimePeriodTickVolume
Case UCase$(TimePeriodNameTickIncrement), "TICK", "TICKS", "TCK", "TCKS", "TM", "TICKSMOVEMENT", "TICKMOVEMENT"
    gTimePeriodUnitsFromString = TimePeriodTickMovement
Case Else
    gTimePeriodUnitsFromString = TimePeriodNone
End Select
End Function

Public Function gTimePeriodUnitsToString( _
                timeUnits As TimePeriodUnits) As String

Select Case timeUnits
Case TimePeriodSecond
    gTimePeriodUnitsToString = TimePeriodNameSeconds
Case TimePeriodMinute
    gTimePeriodUnitsToString = TimePeriodNameMinutes
Case TimePeriodHour
    gTimePeriodUnitsToString = TimePeriodNameHours
Case TimePeriodDay
    gTimePeriodUnitsToString = TimePeriodNameDays
Case TimePeriodWeek
    gTimePeriodUnitsToString = TimePeriodNameWeeks
Case TimePeriodMonth
    gTimePeriodUnitsToString = TimePeriodNameMonths
Case TimePeriodYear
    gTimePeriodUnitsToString = TimePeriodNameYears
Case TimePeriodVolume
    gTimePeriodUnitsToString = TimePeriodNameVolumeIncrement
Case TimePeriodTickVolume
    gTimePeriodUnitsToString = TimePeriodNameTickVolumeIncrement
Case TimePeriodTickMovement
    gTimePeriodUnitsToString = TimePeriodNameTickIncrement
End Select
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcOffsetDailyBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal offset As Long, _
                ByVal sessionStartTime As Date) As Date
Dim datumBarStart As Date
Dim targetWorkingDayNum As Long

datumBarStart = gBarStartTime(timestamp, barLength, TimePeriodDay, sessionStartTime)

targetWorkingDayNum = gCalcWorkingDayNumber(datumBarStart) + offset * barLength

calcOffsetDailyBarStartTime = gCalcWorkingDayDate(targetWorkingDayNum, datumBarStart)
End Function

Private Function calcOffsetMonthlyBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal offset As Long, _
                ByVal sessionStartTime As Date) As Date
Dim datumBarStart As Date

datumBarStart = gBarStartTime(timestamp, barLength, TimePeriodMonth, sessionStartTime)

calcOffsetMonthlyBarStartTime = DateAdd("m", offset * barLength, datumBarStart)
End Function

Private Function calcOffsetWeeklyBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal offset As Long, _
                ByVal sessionStartTime As Date) As Date
Dim datumBarStart As Date
Dim datumWeekNumber As Long
Dim yearStart As Date
Dim yearEnd As Date
Dim yearEndWeekNumber As Long
Dim proposedWeekNumber As Long

datumBarStart = gBarStartTime(timestamp, barLength, TimePeriodWeek, sessionStartTime)
datumWeekNumber = DatePart("ww", datumBarStart, vbMonday, vbFirstFullWeek)

yearStart = DateAdd("d", 1 - DatePart("y", datumBarStart), datumBarStart)
yearEnd = DateAdd("yyyy", 1, yearStart - 1)
yearEndWeekNumber = DatePart("ww", yearEnd, vbMonday, vbFirstFullWeek)

proposedWeekNumber = datumWeekNumber + offset * barLength

Do While proposedWeekNumber < 1 Or proposedWeekNumber > yearEndWeekNumber
    If proposedWeekNumber < 1 Then
        offset = offset + Int(datumWeekNumber / barLength) + 1
        yearEnd = yearStart - 1
        yearStart = DateAdd("yyyy", -1, yearStart)
        yearEndWeekNumber = DatePart("ww", yearEnd, vbMonday, vbFirstFullWeek)
        datumBarStart = gBarStartTime(yearEnd, barLength, TimePeriodWeek, sessionStartTime)
        datumWeekNumber = DatePart("ww", datumBarStart, vbMonday, vbFirstFullWeek)
        
        proposedWeekNumber = datumWeekNumber + offset * barLength
        
    ElseIf proposedWeekNumber > yearEndWeekNumber Then
        offset = offset - Int(yearEndWeekNumber - datumWeekNumber) / barLength - 1
        yearStart = yearEnd + 1
        yearEnd = DateAdd("yyyy", 1, yearEnd)
        yearEndWeekNumber = DatePart("ww", yearEnd, vbMonday, vbFirstFullWeek)
        'datumBarStart = gCalcWeekStartDate(1, yearStart)
        'datumWeekNumber = DatePart("ww", datumBarStart, vbMonday, vbFirstFullWeek)
        datumWeekNumber = 1
        
        proposedWeekNumber = datumWeekNumber + offset * barLength
        
    End If
    
Loop

calcOffsetWeeklyBarStartTime = gCalcWeekStartDate(proposedWeekNumber, yearStart)
End Function


Private Function calcOffsetYearlyBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal offset As Long, _
                ByVal sessionStartTime As Date) As Date
Dim datumBarStart As Date

datumBarStart = gBarStartTime(timestamp, barLength, TimePeriodYear, sessionStartTime)

calcOffsetYearlyBarStartTime = DateAdd("yyyy", offset * barLength, datumBarStart)
End Function

Private Sub calcSessionTimesHelper( _
                ByVal timestamp As Date, _
                ByVal startTime As Date, _
                ByVal endTime As Date, _
                ByRef sessionStartTime As Date, _
                ByRef sessionEndTime As Date)
Dim referenceDate As Date
Dim referenceTime As Date

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
    ' this instrument trades 24hrs, or the caller doesn't know
    ' the session start and end times
    sessionStartTime = referenceDate
    sessionEndTime = referenceDate + 1
End If

End Sub

