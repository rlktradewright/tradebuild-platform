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

Public Const ProjectName                        As String = "TimeframeUtils26"
Private Const ModuleName                As String = "Globals"

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

Public Const TimePeriodShortNameSeconds As String = "s"
Public Const TimePeriodShortNameMinutes As String = "m"
Public Const TimePeriodShortNameHours As String = "h"
Public Const TimePeriodShortNameDays As String = "D"
Public Const TimePeriodShortNameWeeks As String = "W"
Public Const TimePeriodShortNameMonths As String = "M"
Public Const TimePeriodShortNameYears As String = "Y"
Public Const TimePeriodShortNameVolumeIncrement As String = "V"
Public Const TimePeriodShortNameTickVolumeIncrement As String = "TV"
Public Const TimePeriodShortNameTickIncrement As String = "T"

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

Public Property Get gErrorLogger() As Logger
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("error")
Set gErrorLogger = lLogger
End Property

Public Property Get gLogger() As Logger
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("log")
Set gLogger = lLogger
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gBarEndTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Date
Dim startTime As Date
Const ProcName As String = "gBarEndTime"
Dim failpoint As String
On Error GoTo Err

startTime = gBarStartTime( _
                Timestamp, _
                BarTimePeriod, _
                SessionStartTime)
Select Case BarTimePeriod.Units
Case TimePeriodSecond
    gBarEndTime = startTime + (BarTimePeriod.Length / 86400) - OneMicroSecond
Case TimePeriodMinute
    gBarEndTime = startTime + (BarTimePeriod.Length / 1440) - OneMicroSecond
Case TimePeriodHour
    gBarEndTime = startTime + (BarTimePeriod.Length / 24) - OneMicroSecond
Case TimePeriodDay
    gBarEndTime = gCalcWorkingDayDate(gCalcWorkingDayNumber(startTime) + BarTimePeriod.Length, startTime)
Case TimePeriodWeek
    gBarEndTime = startTime + 7 * BarTimePeriod.Length
Case TimePeriodMonth
    gBarEndTime = DateAdd("m", BarTimePeriod.Length, startTime)
Case TimePeriodYear
    gBarEndTime = DateAdd("yyyy", BarTimePeriod.Length, startTime)
Case TimePeriodVolume, _
        TimePeriodTickVolume, _
        TimePeriodTickMovement
    gBarEndTime = Timestamp
End Select

Dim sessStart As Date
Dim sessEnd As Date

calcSessionTimesHelper startTime, SessionStartTime, SessionEndTime, sessStart, sessEnd

If startTime < sessStart And gBarEndTime > sessStart Then
    gBarEndTime = Int(gBarEndTime) + SessionStartTime
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function gBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date) As Date

' minutes from midnight to start of sesssion
Dim sessionOffset           As Long
Dim theDate                 As Long
Dim theTime                 As Double
Dim theTimeMins             As Long
Dim theTimeSecs             As Long

Const ProcName As String = "gBarStartTime"
Dim failpoint As String
On Error GoTo Err

sessionOffset = Int(1440 * (SessionStartTime + OneMicroSecond - Int(SessionStartTime)))

theDate = Int(CDbl(Timestamp))
' NB: don't use TimeValue to get the time, as VB rounds it to
' the nearest second
theTime = CDbl(Timestamp + OneMicroSecond) - theDate

Select Case BarTimePeriod.Units
Case TimePeriodSecond
    theTimeSecs = Fix(theTime * 86400) ' seconds since midnight
    If theTimeSecs < sessionOffset * 60 Then
        theDate = theDate - 1
        theTimeSecs = theTimeSecs + 86400
    End If
    gBarStartTime = theDate + _
                (BarTimePeriod.Length * Int((theTimeSecs - sessionOffset * 60) / BarTimePeriod.Length) + _
                    sessionOffset * 60) / 86400
Case TimePeriodMinute
    theTimeMins = Fix(theTime * 1440) ' minutes since midnight
    If theTimeMins < sessionOffset Then
        theDate = theDate - 1
        theTimeMins = theTimeMins + 1440
    End If
    gBarStartTime = theDate + _
                (BarTimePeriod.Length * Int((theTimeMins - sessionOffset) / BarTimePeriod.Length) + _
                    sessionOffset) / 1440
Case TimePeriodHour
    theTimeMins = Fix(theTime * 1440) ' minutes since midnight
    If theTimeMins < sessionOffset Then
        theDate = theDate - 1
        theTimeMins = theTimeMins + 1440
    End If
    gBarStartTime = theDate + _
                (60 * BarTimePeriod.Length * Int((theTimeMins - sessionOffset) / (60 * BarTimePeriod.Length)) + _
                    sessionOffset) / 1440
Case TimePeriodDay
    Dim workingDayNum As Long
    If theTime < SessionStartTime Then
        theDate = theDate - 1
    End If
    
    If BarTimePeriod.Length = 1 Then
        gBarStartTime = theDate + SessionStartTime
    Else
        workingDayNum = gCalcWorkingDayNumber(theDate)
        
        gBarStartTime = gCalcWorkingDayDate(1 + BarTimePeriod.Length * Int((workingDayNum - 1) / BarTimePeriod.Length), _
                                        theDate) + SessionStartTime
    End If
Case TimePeriodWeek
    Dim weekNum As Long
    
    weekNum = DatePart("ww", theDate, vbMonday, vbFirstFullWeek)
    If weekNum >= 52 And Month(theDate) = 1 Then
        ' this must be part of the final week of the previous year
        theDate = DateAdd("yyyy", -1, theDate)
    End If
    gBarStartTime = gCalcWeekStartDate(1 + BarTimePeriod.Length * Int((weekNum - 1) / BarTimePeriod.Length), _
                                    theDate) + SessionStartTime

Case TimePeriodMonth
    Dim monthNum As Long
    
    monthNum = Month(theDate)
    gBarStartTime = gCalcMonthStartDate(1 + BarTimePeriod.Length * Int((monthNum - 1) / BarTimePeriod.Length), _
                                    theDate) + SessionStartTime
Case TimePeriodYear
    gBarStartTime = DateSerial(1900 + BarTimePeriod.Length * Int((Year(theDate) - 1900) / BarTimePeriod.Length), 1, 1)
Case TimePeriodVolume, _
        TimePeriodTickVolume, _
        TimePeriodTickMovement
    gBarStartTime = Timestamp
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint

End Function

Public Function gCalcBarLength( _
                ByVal BarTimePeriod As TimePeriod) As Date
Const ProcName As String = "gCalcBarLength"
Dim failpoint As String
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodSecond
    gCalcBarLength = BarTimePeriod.Length * OneSecond
Case TimePeriodMinute
    gCalcBarLength = BarTimePeriod.Length * OneMinute
Case TimePeriodHour
    gCalcBarLength = BarTimePeriod.Length * OneHour
Case TimePeriodDay
    gCalcBarLength = BarTimePeriod.Length
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function gCalcMonthStartDate( _
                ByVal monthNumber As Long, _
                ByVal baseDate As Date) As Date
Dim yearStart As Date

Const ProcName As String = "gCalcMonthStartDate"
Dim failpoint As String
On Error GoTo Err

yearStart = DateAdd("d", 1 - DatePart("y", baseDate), baseDate)
gCalcMonthStartDate = DateAdd("m", monthNumber - 1, yearStart)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function


Public Function gCalcNumberOfBarsInSession( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Long

Const ProcName As String = "gCalcNumberOfBarsInSession"
Dim failpoint As String
On Error GoTo Err

If SessionEndTime > SessionStartTime Then
    gCalcNumberOfBarsInSession = -Int(-(SessionEndTime - SessionStartTime) / gCalcBarLength(BarTimePeriod))
Else
    gCalcNumberOfBarsInSession = -Int(-(1 + SessionEndTime - SessionStartTime) / gCalcBarLength(BarTimePeriod))
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function gCalcOffsetBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Date
Dim sessStart As Date
Dim sessEnd As Date
Dim datumBarStart As Date
Dim proposedStart As Date
Dim remainingOffset As Long
Dim BarsFromSessStart As Long
Dim BarsToSessEnd As Long
Dim i As Long
Dim BarLengthDays As Date
Dim numBarsInSession As Long

Const ProcName As String = "gCalcOffsetBarStartTime"
Dim failpoint As String
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
    gCalcOffsetBarStartTime = calcOffsetDailyBarStartTime( _
                                                    Timestamp, _
                                                    BarTimePeriod.Length, _
                                                    offset, _
                                                    SessionStartTime)
    Exit Function
Case TimePeriodWeek
    gCalcOffsetBarStartTime = calcOffsetWeeklyBarStartTime( _
                                                    Timestamp, _
                                                    BarTimePeriod.Length, _
                                                    offset, _
                                                    SessionStartTime)
    Exit Function
Case TimePeriodMonth
    gCalcOffsetBarStartTime = calcOffsetMonthlyBarStartTime( _
                                                    Timestamp, _
                                                    BarTimePeriod.Length, _
                                                    offset, _
                                                    SessionStartTime)
    Exit Function
Case TimePeriodYear
    gCalcOffsetBarStartTime = calcOffsetYearlyBarStartTime( _
                                                    Timestamp, _
                                                    BarTimePeriod.Length, _
                                                    offset, _
                                                    SessionStartTime)
    Exit Function
End Select

BarLengthDays = gCalcBarLength(BarTimePeriod) + OneMicroSecond

numBarsInSession = gCalcNumberOfBarsInSession(BarTimePeriod, SessionStartTime, SessionEndTime)

datumBarStart = gBarStartTime(Timestamp, BarTimePeriod, SessionStartTime)

gCalcSessionTimes Timestamp, _
                    SessionStartTime, _
                    SessionEndTime, _
                    sessStart, _
                    sessEnd

If offset > 0 Then
    
    If datumBarStart < sessStart Then
        ' specified Timestamp was between sessions
        datumBarStart = sessStart
        offset = offset - 1
    End If
    
    BarsToSessEnd = Round((sessEnd - datumBarStart) / BarLengthDays, 0)
    If BarsToSessEnd >= offset Then
        ' all required Bars are in this session
        proposedStart = datumBarStart + offset * BarLengthDays
    Else
        remainingOffset = offset - BarsToSessEnd
        proposedStart = datumBarStart + BarsToSessEnd * BarLengthDays
        Do While remainingOffset > 0
            gCalcSessionTimes proposedStart, _
                                SessionStartTime, _
                                SessionEndTime, _
                                sessStart, _
                                sessEnd
            If numBarsInSession >= remainingOffset Then
                proposedStart = sessStart + remainingOffset * BarLengthDays
                remainingOffset = 0
            Else
                proposedStart = sessEnd
                remainingOffset = remainingOffset - numBarsInSession
            End If
        Loop
    End If
    If proposedStart >= sessEnd Then
        gCalcSessionTimes proposedStart, SessionStartTime, SessionEndTime, sessStart, sessEnd
        proposedStart = sessStart
    End If
Else
    offset = -offset

    If datumBarStart < sessStart Then
        ' specified Timestamp was between sessions
        datumBarStart = sessStart
    End If
    
    proposedStart = sessStart
    BarsFromSessStart = Round((datumBarStart - sessStart) / BarLengthDays, 0)
    If BarsFromSessStart >= offset Then
        ' all required Bars are in this session
        'proposedStart = datumBarStart - offset * BarLengthDays
        proposedStart = sessStart + (BarsFromSessStart - offset) * BarLengthDays
    Else
        remainingOffset = offset - BarsFromSessStart
        proposedStart = sessStart
        Do While remainingOffset > 0
            
            ' find the previous session start - may be a weekend in the
            ' way
            i = 0
            Do
                i = i + 1
                gCalcSessionTimes proposedStart - i, _
                                    SessionStartTime, _
                                    SessionEndTime, _
                                    sessStart, _
                                    sessEnd
            Loop Until sessStart < proposedStart
            
            If numBarsInSession >= remainingOffset Then
                proposedStart = sessStart + (numBarsInSession - remainingOffset) * BarLengthDays
                remainingOffset = 0
            Else
                proposedStart = sessStart
                remainingOffset = remainingOffset - numBarsInSession
            End If
        Loop
    End If
End If

gCalcOffsetBarStartTime = gBarStartTime(proposedStart, _
                                        BarTimePeriod, _
                                        SessionStartTime)


Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint

End Function

Public Sub gCalcOffsetSessionTimes( _
                ByVal Timestamp As Date, _
                ByVal offset As Long, _
                ByVal startTime As Date, _
                ByVal endTime As Date, _
                ByRef SessionStartTime As Date, _
                ByRef SessionEndTime As Date)
Dim datumSessionStart As Date
Dim datumSessionEnd As Date
Dim targetWorkingDayNum As Long

Const ProcName As String = "gCalcOffsetSessionTimes"
Dim failpoint As String
On Error GoTo Err

gCalcSessionTimes Timestamp, startTime, endTime, datumSessionStart, datumSessionEnd

targetWorkingDayNum = gCalcWorkingDayNumber(datumSessionStart) + offset

gCalcSessionTimes gCalcWorkingDayDate(targetWorkingDayNum, datumSessionStart), _
                    startTime, _
                    endTime, _
                    SessionStartTime, _
                    SessionEndTime

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
                
End Sub

' !!!!!!!!!!!!!!!!!!!!!!!!!!!
' calcSessionTimes needs to be amended
' to take acCount of holidays
Public Sub gCalcSessionTimes( _
                ByVal Timestamp As Date, _
                ByVal startTime As Date, _
                ByVal endTime As Date, _
                ByRef SessionStartTime As Date, _
                ByRef SessionEndTime As Date)
Dim weekday As Long

Const ProcName As String = "gCalcSessionTimes"
Dim failpoint As String
On Error GoTo Err

calcSessionTimesHelper Timestamp, _
                        startTime, _
                        endTime, _
                        SessionStartTime, _
                        SessionEndTime

weekday = DatePart("w", SessionStartTime)
If startTime > endTime Then
    ' session DOES span midnight
    If weekday = vbFriday Then
        SessionStartTime = SessionStartTime + 2
        SessionEndTime = SessionEndTime + 2
    ElseIf weekday = vbSaturday Then
        SessionStartTime = SessionStartTime + 1
        SessionEndTime = SessionEndTime + 1
    End If
Else
    ' session doesn't span midnight or 24-hour session or no session times known
    If weekday = vbSaturday Then
        SessionStartTime = SessionStartTime + 2
        SessionEndTime = SessionEndTime + 2
    ElseIf weekday = vbSunday Then
        SessionStartTime = SessionStartTime + 1
        SessionEndTime = SessionEndTime + 1
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint

End Sub

Public Function gCalcWeekStartDate( _
                ByVal weekNumber As Long, _
                ByVal baseDate As Date) As Date
Dim yearStart As Date
Dim week1Date As Date
Dim dow1 As Long    ' day of week of 1st jan of base year

Const ProcName As String = "gCalcWeekStartDate"
Dim failpoint As String
On Error GoTo Err

yearStart = DateAdd("d", 1 - DatePart("y", baseDate), baseDate)

dow1 = DatePart("w", yearStart, vbMonday)

If dow1 = 1 Then
    week1Date = yearStart
Else
    week1Date = DateAdd("d", 8 - dow1, yearStart)
End If

gCalcWeekStartDate = DateAdd("ww", weekNumber - 1, week1Date)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
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

Const ProcName As String = "gCalcWorkingDayDate"
Dim failpoint As String
On Error GoTo Err

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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function



Public Function gCalcWorkingDayNumber( _
                ByVal pDate As Date) As Long
Dim doy As Long     ' day of year
Dim woy As Long     ' week of year
Dim wd1 As Long     ' weekdays in first week (excluding weekend)
'Dim we1 As Long     ' weekend days at start of first week
Dim wdN As Long     ' weekdays in last week

Dim dow1 As Long    ' day of week of 1st jan
Dim dow As Long     ' day of week of sUpplied date

Const ProcName As String = "gCalcWorkingDayNumber"
Dim failpoint As String
On Error GoTo Err

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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint

End Function

Public Function gNormaliseTime( _
            ByVal Timestamp As Date) As Date
gNormaliseTime = Timestamp - Int(Timestamp)
End Function

Public Function gTimePeriodUnitsFromString( _
                timeUnits As String) As TimePeriodUnits

Const ProcName As String = "gTimePeriodUnitsFromString"
Dim failpoint As String
On Error GoTo Err

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
Case UCase$(TimePeriodNameTickIncrement), "TICK", "TICKS", "TCK", "TCKS", "T", "TM", "TICKSMOVEMENT", "TICKMOVEMENT"
    gTimePeriodUnitsFromString = TimePeriodTickMovement
Case Else
    gTimePeriodUnitsFromString = TimePeriodNone
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
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

Public Function gTimePeriodUnitsToShortString( _
                timeUnits As TimePeriodUnits) As String

Select Case timeUnits
Case TimePeriodSecond
    gTimePeriodUnitsToShortString = TimePeriodShortNameSeconds
Case TimePeriodMinute
    gTimePeriodUnitsToShortString = TimePeriodShortNameMinutes
Case TimePeriodHour
    gTimePeriodUnitsToShortString = TimePeriodShortNameHours
Case TimePeriodDay
    gTimePeriodUnitsToShortString = TimePeriodShortNameDays
Case TimePeriodWeek
    gTimePeriodUnitsToShortString = TimePeriodShortNameWeeks
Case TimePeriodMonth
    gTimePeriodUnitsToShortString = TimePeriodShortNameMonths
Case TimePeriodYear
    gTimePeriodUnitsToShortString = TimePeriodShortNameYears
Case TimePeriodVolume
    gTimePeriodUnitsToShortString = TimePeriodShortNameVolumeIncrement
Case TimePeriodTickVolume
    gTimePeriodUnitsToShortString = TimePeriodShortNameTickVolumeIncrement
Case TimePeriodTickMovement
    gTimePeriodUnitsToShortString = TimePeriodShortNameTickIncrement
End Select
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcOffsetDailyBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarLength As Long, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date) As Date
Dim datumBarStart As Date
Dim targetWorkingDayNum As Long

Const ProcName As String = "calcOffsetDailyBarStartTime"
Dim failpoint As String
On Error GoTo Err

datumBarStart = gBarStartTime(Timestamp, gGetTimePeriod(BarLength, TimePeriodDay), SessionStartTime)

targetWorkingDayNum = gCalcWorkingDayNumber(datumBarStart) + offset * BarLength

calcOffsetDailyBarStartTime = gCalcWorkingDayDate(targetWorkingDayNum, datumBarStart)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Private Function calcOffsetMonthlyBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarLength As Long, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date) As Date
Dim datumBarStart As Date

Const ProcName As String = "calcOffsetMonthlyBarStartTime"
Dim failpoint As String
On Error GoTo Err

datumBarStart = gBarStartTime(Timestamp, gGetTimePeriod(BarLength, TimePeriodMonth), SessionStartTime)

calcOffsetMonthlyBarStartTime = DateAdd("m", offset * BarLength, datumBarStart)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Private Function calcOffsetWeeklyBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarLength As Long, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date) As Date
Dim datumBarStart As Date
Dim datumWeekNumber As Long
Dim yearStart As Date
Dim yearEnd As Date
Dim yearEndWeekNumber As Long
Dim proposedWeekNumber As Long

Const ProcName As String = "calcOffsetWeeklyBarStartTime"
Dim failpoint As String
On Error GoTo Err

datumBarStart = gBarStartTime(Timestamp, gGetTimePeriod(BarLength, TimePeriodWeek), SessionStartTime)
datumWeekNumber = DatePart("ww", datumBarStart, vbMonday, vbFirstFullWeek)

yearStart = DateAdd("d", 1 - DatePart("y", datumBarStart), datumBarStart)
yearEnd = DateAdd("yyyy", 1, yearStart - 1)
yearEndWeekNumber = DatePart("ww", yearEnd, vbMonday, vbFirstFullWeek)

proposedWeekNumber = datumWeekNumber + offset * BarLength

Do While proposedWeekNumber < 1 Or proposedWeekNumber > yearEndWeekNumber
    If proposedWeekNumber < 1 Then
        offset = offset + Int(datumWeekNumber / BarLength) + 1
        yearEnd = yearStart - 1
        yearStart = DateAdd("yyyy", -1, yearStart)
        yearEndWeekNumber = DatePart("ww", yearEnd, vbMonday, vbFirstFullWeek)
        datumBarStart = gBarStartTime(yearEnd, gGetTimePeriod(BarLength, TimePeriodWeek), SessionStartTime)
        datumWeekNumber = DatePart("ww", datumBarStart, vbMonday, vbFirstFullWeek)
        
        proposedWeekNumber = datumWeekNumber + offset * BarLength
        
    ElseIf proposedWeekNumber > yearEndWeekNumber Then
        offset = offset - Int(yearEndWeekNumber - datumWeekNumber) / BarLength - 1
        yearStart = yearEnd + 1
        yearEnd = DateAdd("yyyy", 1, yearEnd)
        yearEndWeekNumber = DatePart("ww", yearEnd, vbMonday, vbFirstFullWeek)
        'datumBarStart = gCalcWeekStartDate(1, yearStart)
        'datumWeekNumber = DatePart("ww", datumBarStart, vbMonday, vbFirstFullWeek)
        datumWeekNumber = 1
        
        proposedWeekNumber = datumWeekNumber + offset * BarLength
        
    End If
    
Loop

calcOffsetWeeklyBarStartTime = gCalcWeekStartDate(proposedWeekNumber, yearStart)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function


Private Function calcOffsetYearlyBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarLength As Long, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date) As Date
Dim datumBarStart As Date

Const ProcName As String = "calcOffsetYearlyBarStartTime"
Dim failpoint As String
On Error GoTo Err

datumBarStart = gBarStartTime(Timestamp, gGetTimePeriod(BarLength, TimePeriodYear), SessionStartTime)

calcOffsetYearlyBarStartTime = DateAdd("yyyy", offset * BarLength, datumBarStart)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Private Sub calcSessionTimesHelper( _
                ByVal Timestamp As Date, _
                ByVal startTime As Date, _
                ByVal endTime As Date, _
                ByRef SessionStartTime As Date, _
                ByRef SessionEndTime As Date)
Dim referenceDate As Date
Dim referenceTime As Date

Const ProcName As String = "calcSessionTimesHelper"
Dim failpoint As String
On Error GoTo Err

referenceDate = DateValue(Timestamp)
referenceTime = TimeValue(Timestamp)

If startTime < endTime Then
    ' session doesn't span midnight
    If referenceTime < endTime Then
        SessionStartTime = referenceDate + startTime
        SessionEndTime = referenceDate + endTime
    Else
        SessionStartTime = referenceDate + 1 + startTime
        SessionEndTime = referenceDate + 1 + endTime
    End If
ElseIf startTime > endTime Then
    ' session spans midnight
    If referenceTime >= endTime Then
        SessionStartTime = referenceDate + startTime
        SessionEndTime = referenceDate + 1 + endTime
    Else
        SessionStartTime = referenceDate - 1 + startTime
        SessionEndTime = referenceDate + endTime
    End If
Else
    ' this instrument trades 24hrs, or the caller doesn't know
    ' the session start and end times
    SessionStartTime = referenceDate
    SessionEndTime = referenceDate + 1
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint

End Sub

