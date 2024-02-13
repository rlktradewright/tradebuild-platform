Attribute VB_Name = "BarsGlobals"
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
Private Const ModuleName                        As String = "BarsGlobals"

Public Const OneDayCentiSecs                    As Currency = 8640000

Public Const MaxWorkingDaysPeryear              As Long = 262
Public Const WorkingDaysPerWeek                 As Long = 5
Public Const MinWorkingDaysPerMonth             As Long = 20
Public Const MinWorkingDaysPerYear              As Long = 260

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
startTimeSecs = gBarStartTime(Timestamp, BarTimePeriod, SessionStartTime, SessionEndTime)

Select Case BarTimePeriod.Units
Case TimePeriodSecond, TimePeriodMinute, TimePeriodHour
    gBarEndTime = startTimeSecs + gCalcBarLengthSeconds(BarTimePeriod)
Case TimePeriodDay
    gBarEndTime = gDateToCentiSeconds(GetOffsetSessionTimes(gCentiSecondsToDate(startTimeSecs), BarTimePeriod.Length, SessionStartTime, SessionEndTime).StartTime)
Case TimePeriodWeek
    gBarEndTime = gDateToCentiSeconds(gCentiSecondsToDate(startTimeSecs) + 7 * BarTimePeriod.Length)
Case TimePeriodMonth
    Dim lMonth As Long
    lMonth = Month(gCentiSecondsToDate(startTimeSecs))
    If SessionStartTime > SessionEndTime Then
        lMonth = lMonth + 1
    End If
    Dim lBarEnd As Date
    lBarEnd = Int(MonthStartDateFromMonthNumber(lMonth + BarTimePeriod.Length, gCentiSecondsToDate(startTimeSecs)))
    If SessionStartTime > SessionEndTime Then lBarEnd = lBarEnd - 1
    lBarEnd = lBarEnd + SessionStartTime
    gBarEndTime = gDateToCentiSeconds(lBarEnd)
Case TimePeriodYear
    gBarEndTime = gDateToCentiSeconds(DateAdd("yyyy", BarTimePeriod.Length, gCentiSecondsToDate(startTimeSecs)))
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
        lNextSessionStartTimeSecs = gDateToCentiSeconds(GetSessionTimesIgnoringWeekend(gCentiSecondsToDate(startTimeSecs + OneDayCentiSecs), SessionStartTime, SessionEndTime).StartTime)
        If gBarEndTime > lNextSessionStartTimeSecs Then gBarEndTime = lNextSessionStartTimeSecs
End Select

Exit Function

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Currency
Const ProcName As String = "gBarStartTime"
On Error GoTo Err

AssertArgument SessionStartTime < 1#, "Session start time must be a time value only"
AssertArgument SessionEndTime <= 1#, "Session end time must be a time value only"

' seconds from midnight to start of sesssion
Dim sessionStartSecs As Currency
sessionStartSecs = gDateToCentiSeconds(SessionStartTime)

If SessionEndTime = 0 Then SessionEndTime = 1#
Dim sessionSpansMidnight As Boolean
sessionSpansMidnight = (SessionStartTime > SessionEndTime)

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
        dayNumSecs = dayNumSecs - OneDayCentiSecs
        timeSecs = timeSecs + OneDayCentiSecs
    End If
    gBarStartTime = dayNumSecs + _
                    sessionStartSecs + _
                    barLengthSecs * Int((timeSecs - sessionStartSecs) / barLengthSecs)
Case TimePeriodDay
    If BarTimePeriod.Length = 1 Then
        Dim lSessionTimes As SessionTimes
        gBarStartTime = gDateToCentiSeconds(GetSessionTimes(Timestamp, SessionStartTime, SessionEndTime).StartTime)
        
    Else
        Dim workingDayNum As Long
        workingDayNum = WorkingDayNumber(dayNum)

        gBarStartTime = gDateToCentiSeconds(WorkingDayDate(1 + BarTimePeriod.Length * Int((workingDayNum - 1) / BarTimePeriod.Length), _
                                            dayNum)) + sessionStartSecs
        If timeSecs < sessionStartSecs Then gBarStartTime = gBarStartTime - OneDayCentiSecs
    End If
Case TimePeriodWeek
    Dim weekNum As Long
    If sessionSpansMidnight And timeSecs >= sessionStartSecs Then
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
    If sessionSpansMidnight Then gBarStartTime = gBarStartTime - OneDayCentiSecs
Case TimePeriodMonth
    If sessionSpansMidnight And timeSecs >= sessionStartSecs Then dayNum = dayNum + 1
    Dim monthNum As Long: monthNum = Month(dayNum)
    gBarStartTime = gDateToCentiSeconds(MonthStartDateFromMonthNumber( _
                                            1 + BarTimePeriod.Length * Int((monthNum - 1) / BarTimePeriod.Length), _
                                            dayNum) + _
                                        SessionStartTime)
    If sessionSpansMidnight Then gBarStartTime = gBarStartTime - OneDayCentiSecs
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
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
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
    gCalcBarLengthSeconds = BarTimePeriod.Length * 86400
Case Else
    AssertArgument False, "Invalid BarTimePeriod"
End Select

gCalcBarLengthSeconds = gCalcBarLengthSeconds * 100 ' convert to centiseconds

Exit Function

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCalcNumberOfBarsInSession( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Long
Const ProcName As String = "gCalcNumberOfBarsInSession"
On Error GoTo Err

SessionStartTime = SessionStartTime - Int(SessionStartTime)
SessionEndTime = SessionEndTime - Int(SessionEndTime)
If SessionEndTime > SessionStartTime Then
    gCalcNumberOfBarsInSession = -Int(-(gDateToCentiSeconds(SessionEndTime) - gDateToCentiSeconds(SessionStartTime)) / gCalcBarLengthSeconds(BarTimePeriod))
Else
    gCalcNumberOfBarsInSession = -Int(-(OneDayCentiSecs + gDateToCentiSeconds(SessionEndTime) - gDateToCentiSeconds(SessionStartTime)) / gCalcBarLengthSeconds(BarTimePeriod))
End If

Exit Function

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
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
                                                    SessionStartTime, _
                                                    SessionEndTime)
    Exit Function
Case TimePeriodWeek
    gCalcOffsetBarStartTime = calcOffsetWeeklyBarStartTime( _
                                                    Timestamp, _
                                                    BarTimePeriod, _
                                                    offset, _
                                                    SessionStartTime, _
                                                    SessionEndTime)
    Exit Function
Case TimePeriodMonth
    gCalcOffsetBarStartTime = calcOffsetMonthlyBarStartTime( _
                                                    Timestamp, _
                                                    BarTimePeriod, _
                                                    offset, _
                                                    SessionStartTime, _
                                                    SessionEndTime)
    Exit Function
Case TimePeriodYear
    gCalcOffsetBarStartTime = calcOffsetYearlyBarStartTime( _
                                                    Timestamp, _
                                                    BarTimePeriod.Length, _
                                                    offset, _
                                                    SessionStartTime, _
                                                    SessionEndTime)
    Exit Function
Case Else
    AssertArgument False, "Invalid BarTimePeriod"
End Select

Dim barLengthSecs As Currency
barLengthSecs = gCalcBarLengthSeconds(BarTimePeriod)

Dim numBarsInSession As Long
numBarsInSession = gCalcNumberOfBarsInSession(BarTimePeriod, SessionStartTime, SessionEndTime)

Dim datumBarStartSecs As Currency
datumBarStartSecs = gBarStartTime(Timestamp, BarTimePeriod, SessionStartTime, SessionEndTime)

Dim lSessTimes As SessionTimes
lSessTimes = GetSessionTimes(Timestamp, _
                    SessionStartTime, _
                    SessionEndTime)

Dim proposedStartSecs As Currency
Dim remainingOffset As Long
If offset > 0 Then
    If datumBarStartSecs > gDateToCentiSeconds(lSessTimes.EndTime) Then
        ' specified Timestamp was between sessions
        lSessTimes = GetOffsetSessionTimes(gCentiSecondsToDate(datumBarStartSecs), 1, SessionStartTime, SessionEndTime)
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
            lSessTimes = GetOffsetSessionTimes(gCentiSecondsToDate(proposedStartSecs), 1, SessionStartTime, SessionEndTime)
            
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

    If datumBarStartSecs > gDateToCentiSeconds(lSessTimes.EndTime) Then
        ' specified Timestamp was between sessions
        datumBarStartSecs = gDateToCentiSeconds(lSessTimes.EndTime)
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
            lSessTimes = GetSessionTimes(gCentiSecondsToDate(proposedStartSecs - OneDayCentiSecs), SessionStartTime, SessionEndTime)
            
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

gCalcOffsetBarStartTime = gCentiSecondsToDate(proposedStartSecs)

Exit Function

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function

' returns the date in units of 100th of a second. Expressed as a Currency data type,
' this gives a resolution of one microsecond.
Public Function gDateToCentiSeconds(ByVal pDate As Date) As Currency
gDateToCentiSeconds = CCur(CDbl(pDate) * 8640000)
End Function

Public Function gCentiSecondsToDate(ByVal pSeconds As Currency) As Date
gCentiSecondsToDate = CDate(CDbl(pSeconds) / 8640000)
End Function

Public Sub gGetTimespanData( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByRef pFromTime As Date, _
                ByRef pToTime As Date, _
                ByRef pFromSessionTimes As SessionTimes, _
                ByRef pToSessionTimes As SessionTimes, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date)
Const ProcName As String = "gGetTimespanData"
On Error GoTo Err

AssertArgument pFromTime <> 0, "pFromTime must be supplied"
AssertArgument pBarTimePeriod.Length <> 0, "pBarTimePeriod.Length is 0"
        
If pToTime = 0 Then pToTime = Now

Dim lStartTime As Date
lStartTime = gCentiSecondsToDate(gBarStartTime(pFromTime, pBarTimePeriod, pSessionStartTime, pSessionEndTime))

pFromSessionTimes = GetSessionTimes(lStartTime, pSessionStartTime, pSessionEndTime)
If lStartTime > pFromSessionTimes.EndTime Then
    pFromSessionTimes = GetOffsetSessionTimes(lStartTime, 1, pSessionStartTime, pSessionEndTime)
    lStartTime = pFromSessionTimes.StartTime
End If

Dim lEndTime As Date
lEndTime = gCentiSecondsToDate(gBarStartTime(pToTime, pBarTimePeriod, pSessionStartTime, pSessionEndTime))
If lEndTime < pToTime Then
    lEndTime = gCentiSecondsToDate(gBarEndTime(pToTime, pBarTimePeriod, pSessionStartTime, pSessionEndTime))
End If

pToSessionTimes = GetSessionTimes(lEndTime, pSessionStartTime, pSessionEndTime)
If lEndTime > pToSessionTimes.EndTime Then
    lEndTime = pToSessionTimes.EndTime
End If

If lEndTime > pToTime Then lEndTime = pToTime

pFromTime = lStartTime
pToTime = lEndTime

Exit Sub

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Sub

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

Public Function gMaxNumberOfBarsInTimespanNormalized( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pStartTime As Date, _
                ByVal pEndTime As Date, _
                ByRef pStartSessionTimes As SessionTimes, _
                ByRef pEndSessionTimes As SessionTimes) As Long
Const ProcName As String = "gMaxNumberOfBarsInTimespanNormalized"
On Error GoTo Err

Select Case pBarTimePeriod.Units
    Case TimePeriodNone, TimePeriodTickMovement, TimePeriodTickVolume, TimePeriodVolume
        AssertArgument False, "Must be a fixed time period"
End Select
AssertArgument pStartTime >= pStartSessionTimes.StartTime, _
                "pStartTime is not in pStartSessionTimes"
AssertArgument pEndTime <= pEndSessionTimes.EndTime, _
                "pEndTime is not in pEndSessionTimes"

Dim lNumberOfBars As Long
If pStartSessionTimes.StartTime = pEndSessionTimes.StartTime Then
    lNumberOfBars = calcNumberOfBarsInTimespan(pBarTimePeriod, pStartTime, pEndTime)
Else
    Dim lStartWorkingDate As Date
    lStartWorkingDate = (pStartSessionTimes.StartTime + pStartSessionTimes.EndTime) / 2#
    Dim lStartWorkingDayNumber As Long
    lStartWorkingDayNumber = WorkingDayNumber(lStartWorkingDate)

    Dim lEndWorkingDate As Date
    lEndWorkingDate = (pEndSessionTimes.StartTime + pEndSessionTimes.EndTime) / 2#
    Dim lEndWorkingDayNumber As Long
    lEndWorkingDayNumber = WorkingDayNumber(lEndWorkingDate)

    Dim lNumberOfWorkingDaysInInterval As Long
    If Year(lStartWorkingDate) = Year(lEndWorkingDate) Then
        lNumberOfWorkingDaysInInterval = lEndWorkingDayNumber - lStartWorkingDayNumber + 1
    Else
        lNumberOfWorkingDaysInInterval = _
                        MaxWorkingDaysPeryear - lStartWorkingDayNumber + _
                        lEndWorkingDayNumber + 1 + _
                        (Year(lEndWorkingDate) - Year(lStartWorkingDate) - 1) * MaxWorkingDaysPeryear
    End If
        
    If pBarTimePeriod.Units = TimePeriodWeek Then
        lNumberOfBars = -Int(-lNumberOfWorkingDaysInInterval / WorkingDaysPerWeek)
    ElseIf pBarTimePeriod.Units = TimePeriodMonth Then
        lNumberOfBars = -Int(-lNumberOfWorkingDaysInInterval / MinWorkingDaysPerMonth)
    ElseIf pBarTimePeriod.Units = TimePeriodYear Then
        lNumberOfBars = -Int(-lNumberOfWorkingDaysInInterval / MinWorkingDaysPerYear)
    Else
        lNumberOfBars = calcNumberOfBarsInTimespan(pBarTimePeriod, pStartTime, pStartSessionTimes.EndTime)
        lNumberOfBars = lNumberOfBars + _
                        calcNumberOfBarsInTimespan(pBarTimePeriod, pEndSessionTimes.StartTime, pEndTime)
    
        lNumberOfBars = lNumberOfBars + _
                        (lNumberOfWorkingDaysInInterval - 2) * _
                            gCalcNumberOfBarsInSession(pBarTimePeriod, pStartSessionTimes.StartTime, pStartSessionTimes.EndTime)
    End If
End If

gMaxNumberOfBarsInTimespanNormalized = lNumberOfBars

Exit Function

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gNormaliseSessionTime( _
            ByVal Timestamp As Date) As Date
If CDbl(Timestamp) = 1# Then
    gNormaliseSessionTime = 1#
Else
    gNormaliseSessionTime = CDate(Round(86400# * (CDbl(Timestamp) - CDbl(Int(Timestamp)))) / 86400#)
End If
End Function

Public Function gNormaliseTimestamp( _
                ByVal pTimestamp As Date, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
Select Case pTimePeriod.Units
Case TimePeriodDay, TimePeriodWeek, TimePeriodMonth, TimePeriodYear
    If pSessionStartTime > pSessionEndTime And _
        pTimestamp >= pSessionStartTime _
    Then
        gNormaliseTimestamp = Int(pTimestamp) + 1
    Else
        gNormaliseTimestamp = Int(pTimestamp)
    End If
Case Else
    gNormaliseTimestamp = pTimestamp
End Select
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

Private Function calcNumberOfBarsInTimespan( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStartTime As Date, _
                ByVal pEndTime As Date) As Long
Const ProcName As String = "calcNumberOfBarsInTimespan"
On Error GoTo Err

AssertArgument (Int(pStartTime) = Int(pEndTime)) Or _
                (pEndTime <= pStartTime + 1), _
                "Invalid timespan for this function"

Dim lBarLengthCentiSecs As Currency

Select Case pTimePeriod.Units
    Case TimePeriodSecond
        lBarLengthCentiSecs = pTimePeriod.Length * 100
    Case TimePeriodMinute
        lBarLengthCentiSecs = pTimePeriod.Length * 60 * 100
    Case TimePeriodHour
        lBarLengthCentiSecs = pTimePeriod.Length * 3600 * 100
    Case TimePeriodDay
        lBarLengthCentiSecs = pTimePeriod.Length * OneDayCentiSecs
    Case TimePeriodWeek
        lBarLengthCentiSecs = pTimePeriod.Length * OneDayCentiSecs * WorkingDaysPerWeek
    Case TimePeriodMonth
        lBarLengthCentiSecs = pTimePeriod.Length * OneDayCentiSecs * MinWorkingDaysPerMonth
    Case TimePeriodYear
        lBarLengthCentiSecs = pTimePeriod.Length * OneDayCentiSecs * MinWorkingDaysPerYear
End Select

calcNumberOfBarsInTimespan = Int((gDateToCentiSeconds(pEndTime - pStartTime) + lBarLengthCentiSecs - 1) / lBarLengthCentiSecs)

Exit Function

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcOffsetDailyBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Date
Const ProcName As String = "calcOffsetDailyBarStartTime"
On Error GoTo Err

calcOffsetDailyBarStartTime = GetOffsetSessionTimes(Timestamp, offset, SessionStartTime, SessionEndTime).StartTime

Exit Function

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcOffsetMonthlyBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Date
Const ProcName As String = "calcOffsetMonthlyBarStartTime"
On Error GoTo Err

Dim datumBarStart As Date
datumBarStart = gCentiSecondsToDate(gBarStartTime(Timestamp, BarTimePeriod, SessionStartTime, SessionEndTime))

calcOffsetMonthlyBarStartTime = DateAdd("m", offset * BarTimePeriod.Length, datumBarStart)

Exit Function

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcOffsetWeeklyBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Date
Const ProcName As String = "calcOffsetWeeklyBarStartTime"
On Error GoTo Err

Dim datumBarStart As Date
datumBarStart = gCentiSecondsToDate(gBarStartTime(Timestamp, BarTimePeriod, SessionStartTime, SessionEndTime))

Dim datumWeekNumber As Long
datumWeekNumber = DatePart("ww", datumBarStart, vbMonday, vbFirstFullWeek)

Dim proposedWeekNumber As Long
proposedWeekNumber = datumWeekNumber + offset * BarTimePeriod.Length

calcOffsetWeeklyBarStartTime = WeekStartDateFromWeekNumber(proposedWeekNumber, datumBarStart)

Exit Function

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function


Private Function calcOffsetYearlyBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal barLength As Long, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Date
Const ProcName As String = "calcOffsetYearlyBarStartTime"
On Error GoTo Err

Dim datumBarStart As Date
datumBarStart = gCentiSecondsToDate(gBarStartTime(Timestamp, GetTimePeriod(barLength, TimePeriodYear), SessionStartTime, SessionEndTime))

calcOffsetYearlyBarStartTime = DateAdd("yyyy", offset * barLength, datumBarStart)

Exit Function

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function

