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

Public Const ProjectName                        As String = "BarUtils27"
Private Const ModuleName                        As String = "Globals"

Public Const OneDay As Currency = 8640000       ' in centisecond units
Public Const MidDay As Currency = 4320000       ' in centisecond units

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

Public Property Get gLogger() As FormattingLogger
Static lLogger As FormattingLogger
If lLogger Is Nothing Then Set lLogger = CreateFormattingLogger("barutils.log", ProjectName)
Set gLogger = lLogger
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gBarEndTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Currency
Const ProcName As String = "gBarEndTime"
On Error GoTo Err

Dim startTimeSecs As Currency
startTimeSecs = gBarStartTime(Timestamp, BarTimePeriod, SessionStartTime)

Select Case BarTimePeriod.Units
Case TimePeriodSecond, TimePeriodMinute, TimePeriodHour
    gBarEndTime = startTimeSecs + gCalcBarLengthSeconds(BarTimePeriod)
Case TimePeriodDay
    gBarEndTime = gDateToCentiSeconds(GetOffsetSessionTimes(gCentiSecondstoDate(startTimeSecs), BarTimePeriod.Length, SessionStartTime, SessionEndTime).StartTime)
Case TimePeriodWeek
    gBarEndTime = gDateToCentiSeconds(gCentiSecondstoDate(startTimeSecs) + 7 * BarTimePeriod.Length)
Case TimePeriodMonth
    gBarEndTime = gDateToCentiSeconds(DateAdd("m", BarTimePeriod.Length, gCentiSecondstoDate(startTimeSecs)))
Case TimePeriodYear
    gBarEndTime = gDateToCentiSeconds(DateAdd("yyyy", BarTimePeriod.Length, gCentiSecondstoDate(startTimeSecs)))
Case TimePeriodVolume, _
        TimePeriodTickVolume, _
        TimePeriodTickMovement, _
        TimePeriodNone
    gBarEndTime = gDateToCentiSeconds(Timestamp)
End Select

Select Case BarTimePeriod.Units
    Case TimePeriodSecond, _
            TimePeriodMinute, _
            TimePeriodHour
        
        ' adjust if bar is at end of session but does not fit exactly into session
        
        Dim lNextSessionStartTimeSecs  As Currency
        lNextSessionStartTimeSecs = gDateToCentiSeconds(GetSessionTimesIgnoringWeekend(gCentiSecondstoDate(startTimeSecs + OneDay), SessionStartTime, SessionEndTime).StartTime)
        If gBarEndTime > lNextSessionStartTimeSecs Then gBarEndTime = lNextSessionStartTimeSecs
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date) As Currency
Const ProcName As String = "gBarStartTime"
On Error GoTo Err

AssertArgument SessionStartTime < 1#, "Session start time must be a time value only"

' seconds from midnight to start of sesssion
Dim sessionStartSecs As Currency
sessionStartSecs = gDateToCentiSeconds(SessionStartTime)

Dim dayNum As Long: dayNum = Int(CDbl(Timestamp))

Dim dayNumSecs As Currency
dayNumSecs = gDateToCentiSeconds(dayNum)

' NB: don't use TimeValue to get the time, as VB rounds it to
' the nearest second
Dim timeSecs As Currency
timeSecs = gDateToCentiSeconds(Timestamp) - dayNumSecs  ' seconds since midnight

Select Case BarTimePeriod.Units
Case TimePeriodSecond, TimePeriodMinute, TimePeriodHour
    Dim barLengthSecs As Currency: barLengthSecs = gCalcBarLengthSeconds(BarTimePeriod)
    
    If timeSecs < sessionStartSecs Then
        dayNumSecs = dayNumSecs - OneDay
        timeSecs = timeSecs + OneDay
    End If
    gBarStartTime = dayNumSecs + _
                    sessionStartSecs + _
                    barLengthSecs * Int((timeSecs - sessionStartSecs) / barLengthSecs)
Case TimePeriodDay
    If BarTimePeriod.Length = 1 Then
        Dim lSessionEndTime As Date
        If sessionStartSecs >= MidDay Then
            lSessionEndTime = gCentiSecondstoDate(sessionStartSecs - 100)
        Else
            lSessionEndTime = CDbl(1#)
        End If
        
        Dim lSessionTimes As SessionTimes
        gBarStartTime = gDateToCentiSeconds(GetSessionTimes(Timestamp, SessionStartTime, lSessionEndTime).StartTime)
        
    Else
        Dim workingDayNum As Long
        workingDayNum = WorkingDayNumber(dayNum)

        gBarStartTime = gDateToCentiSeconds(WorkingDayDate(1 + BarTimePeriod.Length * Int((workingDayNum - 1) / BarTimePeriod.Length), _
                                            dayNum)) + sessionStartSecs
        If timeSecs < sessionStartSecs Then gBarStartTime = gBarStartTime - OneDay
    End If
Case TimePeriodWeek
    Dim weekNum As Long
    If sessionStartSecs >= MidDay Then
        weekNum = DatePart("ww", dayNum + 1, vbMonday, vbFirstFullWeek)
    Else
        weekNum = DatePart("ww", dayNum, vbMonday, vbFirstFullWeek)
    End If
    
    If weekNum >= 52 And Month(dayNum) = 1 Then
        ' this must be part of the final week of the previous year
        dayNum = DateAdd("yyyy", -1, dayNum)
    End If
    gBarStartTime = gDateToCentiSeconds(WeekStartDateFromWeekNumber(1 + BarTimePeriod.Length * Int((weekNum - 1) / BarTimePeriod.Length), _
                                        dayNum) + SessionStartTime)
    If sessionStartSecs >= MidDay Then gBarStartTime = gBarStartTime - OneDay
Case TimePeriodMonth
    Dim monthNum As Long
    
    monthNum = Month(dayNum)
    gBarStartTime = gDateToCentiSeconds(MonthStartDateFromMonthNumber(1 + BarTimePeriod.Length * Int((monthNum - 1) / BarTimePeriod.Length), _
                                        dayNum) + SessionStartTime)
Case TimePeriodYear
    gBarStartTime = gDateToCentiSeconds(CDate(DateSerial(1900 + BarTimePeriod.Length * Int((Year(Timestamp) - 1900) / BarTimePeriod.Length), 1, 1)))
Case TimePeriodVolume, _
        TimePeriodTickVolume, _
        TimePeriodTickMovement, _
        TimePeriodNone
    gBarStartTime = gDateToCentiSeconds(Timestamp)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCalcBarLengthSeconds( _
                ByVal BarTimePeriod As TimePeriod) As Currency
Const ProcName As String = "gCalcBarLengthSeconds"
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodSecond
    gCalcBarLengthSeconds = BarTimePeriod.Length
Case TimePeriodMinute
    gCalcBarLengthSeconds = BarTimePeriod.Length * 60
Case TimePeriodHour
    gCalcBarLengthSeconds = BarTimePeriod.Length * 3600
Case TimePeriodDay
    gCalcBarLengthSeconds = BarTimePeriod.Length * 1440
Case Else
    AssertArgument False, "Invalid BarTimePeriod"
End Select

gCalcBarLengthSeconds = gCalcBarLengthSeconds * 100 ' convert to centiseconds

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCalcNumberOfBarsInSession( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Long
Const ProcName As String = "gCalcNumberOfBarsInSession"
On Error GoTo Err

If SessionEndTime > SessionStartTime Then
    gCalcNumberOfBarsInSession = Int((gDateToCentiSeconds(SessionEndTime) - gDateToCentiSeconds(SessionStartTime)) / gCalcBarLengthSeconds(BarTimePeriod))
Else
    gCalcNumberOfBarsInSession = Int((OneDay + gDateToCentiSeconds(SessionEndTime) - gDateToCentiSeconds(SessionStartTime)) / gCalcBarLengthSeconds(BarTimePeriod))
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCalcOffsetBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Date
Const ProcName As String = "gCalcOffsetBarStartTime"
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
    gCalcOffsetBarStartTime = calcOffsetDailyBarStartTime( _
                                                    Timestamp, _
                                                    BarTimePeriod, _
                                                    offset, _
                                                    SessionStartTime)
    Exit Function
Case TimePeriodWeek
    gCalcOffsetBarStartTime = calcOffsetWeeklyBarStartTime( _
                                                    Timestamp, _
                                                    BarTimePeriod, _
                                                    offset, _
                                                    SessionStartTime)
    Exit Function
Case TimePeriodMonth
    gCalcOffsetBarStartTime = calcOffsetMonthlyBarStartTime( _
                                                    Timestamp, _
                                                    BarTimePeriod, _
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

Dim barLengthSecs As Currency
barLengthSecs = gCalcBarLengthSeconds(BarTimePeriod)

Dim numBarsInSession As Long
numBarsInSession = gCalcNumberOfBarsInSession(BarTimePeriod, SessionStartTime, SessionEndTime)

Dim datumBarStartSecs As Currency
datumBarStartSecs = gBarStartTime(Timestamp, BarTimePeriod, SessionStartTime)

Dim lSessTimes As SessionTimes
lSessTimes = GetSessionTimes(Timestamp, _
                    SessionStartTime, _
                    SessionEndTime)

Dim proposedStartSecs As Currency
Dim remainingOffset As Long
If offset > 0 Then
    If datumBarStartSecs > gDateToCentiSeconds(lSessTimes.EndTime) Then
        ' specified Timestamp was between sessions
        lSessTimes = GetOffsetSessionTimes(gCentiSecondstoDate(datumBarStartSecs), 1, SessionStartTime, SessionEndTime)
        datumBarStartSecs = gDateToCentiSeconds(lSessTimes.StartTime)
        offset = offset - 1
    End If
    
    Dim BarsToSessEnd As Long
    BarsToSessEnd = Int((gDateToCentiSeconds(lSessTimes.EndTime) - datumBarStartSecs) / barLengthSecs)
    If BarsToSessEnd > offset Then
        ' all required Bars are in this session
        proposedStartSecs = datumBarStartSecs + offset * barLengthSecs
    Else
        remainingOffset = offset - BarsToSessEnd
        proposedStartSecs = gDateToCentiSeconds(lSessTimes.EndTime)
        Do While remainingOffset >= 0
            lSessTimes = GetOffsetSessionTimes(gCentiSecondstoDate(proposedStartSecs), 1, SessionStartTime, SessionEndTime)
            
            If numBarsInSession > remainingOffset Then
                proposedStartSecs = gDateToCentiSeconds(lSessTimes.StartTime) + remainingOffset * barLengthSecs
                remainingOffset = -1
            Else
                proposedStartSecs = gDateToCentiSeconds(lSessTimes.EndTime)
                remainingOffset = remainingOffset - numBarsInSession
            End If
        Loop
    End If
Else
    offset = -offset

    If datumBarStartSecs >= gDateToCentiSeconds(lSessTimes.EndTime) Then
        ' specified Timestamp was between sessions
        datumBarStartSecs = gBarEndTime(lSessTimes.EndTime, BarTimePeriod, SessionStartTime, SessionEndTime)
    End If
    
    proposedStartSecs = gDateToCentiSeconds(lSessTimes.StartTime)
    
    Dim BarsFromSessStart As Long
    BarsFromSessStart = Int((datumBarStartSecs - gDateToCentiSeconds(lSessTimes.StartTime)) / barLengthSecs)
    If BarsFromSessStart >= offset Then
        ' all required Bars are in this session
        proposedStartSecs = gDateToCentiSeconds(lSessTimes.StartTime) + (BarsFromSessStart - offset) * barLengthSecs
    Else
        remainingOffset = offset - BarsFromSessStart
        Do While remainingOffset > 0
            lSessTimes = GetSessionTimes(gCentiSecondstoDate(proposedStartSecs - OneDay), SessionStartTime, SessionEndTime)
            
            If numBarsInSession >= remainingOffset Then
                proposedStartSecs = gDateToCentiSeconds(lSessTimes.EndTime) - remainingOffset * barLengthSecs
                remainingOffset = 0
            Else
                proposedStartSecs = gDateToCentiSeconds(lSessTimes.StartTime)
                remainingOffset = remainingOffset - numBarsInSession
            End If
        Loop
    End If
End If

gCalcOffsetBarStartTime = gCentiSecondstoDate(proposedStartSecs)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

' returns the date in units of 100th of a second. Expresses as a Currency data type,
' this gives a resolution of one microsecond.
Public Function gDateToCentiSeconds(ByVal pDate As Date) As Currency
gDateToCentiSeconds = CCur(CDbl(pDate) * 8640000)
End Function

Public Function gCentiSecondstoDate(ByVal pSeconds As Currency) As Date
gCentiSecondstoDate = CDate(CDbl(pSeconds) / 8640000)
End Function

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Function gMaxNumberOfBarsInTimespan( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pStartTime As Date, _
                ByVal pEndTime As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Long
Const ProcName As String = "gMaxNumberOfBarsInTimespan"
On Error GoTo Err

AssertArgument pStartTime <> 0, "pStartTime must be supplied"
AssertArgument pBarTimePeriod.Length <> 0, "pBarTimePeriod.Length is 0"
Select Case pBarTimePeriod.Units
    Case TimePeriodNone, TimePeriodTickMovement, TimePeriodTickVolume, TimePeriodVolume
        AssertArgument False, "Must be a fixed time period"
End Select
        
If pEndTime = 0 Then pEndTime = Now

Dim lStartTime As Date
lStartTime = gCentiSecondstoDate(gBarStartTime(pStartTime, pBarTimePeriod, pSessionStartTime))

Dim lEndTime As Date
lEndTime = gCentiSecondstoDate(gBarEndTime(pEndTime, pBarTimePeriod, pSessionStartTime, pSessionEndTime))

Select Case pBarTimePeriod.Units
    Case TimePeriodSecond
        gMaxNumberOfBarsInTimespan = Int((86400# * (lEndTime - lStartTime) + pBarTimePeriod.Length - 1) / pBarTimePeriod.Length)
    Case TimePeriodMinute
        gMaxNumberOfBarsInTimespan = Int((1440# * (lEndTime - lStartTime) + pBarTimePeriod.Length - 1) / pBarTimePeriod.Length)
    Case TimePeriodHour
        gMaxNumberOfBarsInTimespan = Int((24# * (lEndTime - lStartTime) + pBarTimePeriod.Length - 1) / pBarTimePeriod.Length)
    Case TimePeriodDay
        gMaxNumberOfBarsInTimespan = Int(((lEndTime - lStartTime) + pBarTimePeriod.Length - 1) / pBarTimePeriod.Length)
    Case TimePeriodWeek
        gMaxNumberOfBarsInTimespan = Int(((lEndTime - lStartTime) / 7 + pBarTimePeriod.Length - 1) / pBarTimePeriod.Length)
    Case TimePeriodMonth
        gMaxNumberOfBarsInTimespan = Int(((Year(lEndTime) - Year(lStartTime)) * 12 + Month(lEndTime) - Month(lStartTime) + pBarTimePeriod.Length - 1) / pBarTimePeriod.Length)
    Case TimePeriodYear
        gMaxNumberOfBarsInTimespan = Int((Year(lEndTime) - Year(lStartTime) + pBarTimePeriod.Length - 1) / pBarTimePeriod.Length)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gNormaliseSessionTime( _
            ByVal Timestamp As Date) As Date
If CDbl(Timestamp) = 1# Then
    gNormaliseSessionTime = 1#
Else
    gNormaliseSessionTime = CDate(Round(86400# * (CDbl(Timestamp) - CDbl(Int(Timestamp)))) / 86400#)
End If
End Function

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcOffsetDailyBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date) As Date
Const ProcName As String = "calcOffsetDailyBarStartTime"
On Error GoTo Err

calcOffsetDailyBarStartTime = GetOffsetSessionTimes(Timestamp, offset, SessionStartTime, 0#).StartTime

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcOffsetMonthlyBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date) As Date
Const ProcName As String = "calcOffsetMonthlyBarStartTime"
On Error GoTo Err

Dim datumBarStart As Date
datumBarStart = gCentiSecondstoDate(gBarStartTime(Timestamp, BarTimePeriod, SessionStartTime))

calcOffsetMonthlyBarStartTime = DateAdd("m", offset * BarTimePeriod.Length, datumBarStart)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcOffsetWeeklyBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date) As Date
Const ProcName As String = "calcOffsetWeeklyBarStartTime"
On Error GoTo Err

Dim datumBarStart As Date
datumBarStart = gCentiSecondstoDate(gBarStartTime(Timestamp, BarTimePeriod, SessionStartTime))

Dim datumWeekNumber As Long
datumWeekNumber = DatePart("ww", datumBarStart, vbMonday, vbFirstFullWeek)

Dim yearStart As Date
yearStart = DateAdd("d", 1 - DatePart("y", datumBarStart), datumBarStart)

Dim yearEnd As Date
yearEnd = DateAdd("yyyy", 1, yearStart - 1)

Dim yearEndWeekNumber As Long
yearEndWeekNumber = DatePart("ww", yearEnd, vbMonday, vbFirstFullWeek)

Dim proposedWeekNumber As Long
proposedWeekNumber = datumWeekNumber + offset * BarTimePeriod.Length

Do While proposedWeekNumber < 1 Or proposedWeekNumber > yearEndWeekNumber
    If proposedWeekNumber < 1 Then
        offset = offset + Int(datumWeekNumber / BarTimePeriod.Length) + 1
        yearEnd = yearStart - 1
        yearStart = DateAdd("yyyy", -1, yearStart)
        yearEndWeekNumber = DatePart("ww", yearEnd, vbMonday, vbFirstFullWeek)
        datumBarStart = gCentiSecondstoDate(gBarStartTime(yearEnd, GetTimePeriod(BarTimePeriod.Length, TimePeriodWeek), SessionStartTime))
        datumWeekNumber = DatePart("ww", datumBarStart, vbMonday, vbFirstFullWeek)
        
        proposedWeekNumber = datumWeekNumber + offset * BarTimePeriod.Length
        
    ElseIf proposedWeekNumber > yearEndWeekNumber Then
        offset = offset - Int(yearEndWeekNumber - datumWeekNumber) / BarTimePeriod.Length - 1
        yearStart = yearEnd + 1
        yearEnd = DateAdd("yyyy", 1, yearEnd)
        yearEndWeekNumber = DatePart("ww", yearEnd, vbMonday, vbFirstFullWeek)
        'datumBarStart = gCalcWeekStartDate(1, yearStart)
        'datumWeekNumber = DatePart("ww", datumBarStart, vbMonday, vbFirstFullWeek)
        datumWeekNumber = 1
        
        proposedWeekNumber = datumWeekNumber + offset * BarTimePeriod.Length
        
    End If
    
Loop

calcOffsetWeeklyBarStartTime = WeekStartDateFromWeekNumber(proposedWeekNumber, yearStart)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


Private Function calcOffsetYearlyBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal barLength As Long, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date) As Date
Const ProcName As String = "calcOffsetYearlyBarStartTime"
On Error GoTo Err

Dim datumBarStart As Date
datumBarStart = gCentiSecondstoDate(gBarStartTime(Timestamp, GetTimePeriod(barLength, TimePeriodYear), SessionStartTime))

calcOffsetYearlyBarStartTime = DateAdd("yyyy", offset * barLength, datumBarStart)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

