Attribute VB_Name = "GBarUtils"
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

Private Const ModuleName                        As String = "GBarUtils"

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

'@================================================================================
' Methods
'@================================================================================

Public Function BarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                Optional ByVal SessionStartTime As Date, _
                Optional ByVal SessionEndTime As Date) As Date
Const ProcName As String = "BarStartTime"
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodVolume, _
        TimePeriodTickVolume, _
        TimePeriodTickMovement, _
        TimePeriodNone
    BarStartTime = Timestamp
Case Else
    BarStartTime = GBarUtils.centiSecondsToDate(GBarUtils.calcBarStartTime(Timestamp, _
                                    BarTimePeriod, _
                                    GBarUtils.NormaliseSessionTime(SessionStartTime), _
                                    GBarUtils.NormaliseSessionTime(SessionEndTime)))
End Select

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function BarEndTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                Optional ByVal SessionStartTime As Date, _
                Optional ByVal SessionEndTime As Date) As Date
Const ProcName As String = "BarEndTime"
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodVolume, _
        TimePeriodTickVolume, _
        TimePeriodTickMovement, _
        TimePeriodNone
    BarEndTime = Timestamp
Case Else
    BarEndTime = GBarUtils.centiSecondsToDate(GBarUtils.calcBarEndTime(Timestamp, _
                                    BarTimePeriod, _
                                    GBarUtils.NormaliseSessionTime(SessionStartTime), _
                                    GBarUtils.NormaliseSessionTime(SessionEndTime)))
End Select

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBar( _
                ByVal Timestamp As Date, _
                ByVal OpenValue As Double, _
                ByVal HighValue As Double, _
                ByVal LowValue As Double, _
                ByVal CloseValue As Double, _
                Optional ByVal Volume As BoxedDecimal, _
                Optional ByVal TickVolume As Long, _
                Optional ByVal OpenInterest As Long) As Bar
Const ProcName As String = "CreateBar"
On Error GoTo Err

Set CreateBar = New Bar
CreateBar.Initialise Timestamp, _
                    OpenValue, _
                    HighValue, _
                    LowValue, _
                    CloseValue, _
                    Volume, _
                    TickVolume, _
                    OpenInterest

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarsBuilder( _
                ByVal pBarTimePeriod As TimePeriod, _
                Optional ByVal pSession As Session, _
                Optional ByVal pTickSize As Double, _
                Optional ByVal pNumberOfBarsToCache As Long, _
                Optional ByVal pNormaliseDailyTimestamps As Boolean, _
                Optional ByVal pSave As Boolean = True) As BarsBuilder
Const ProcName As String = "CreateBarsBuilder"
On Error GoTo Err

AssertArgument Not pBarTimePeriod Is Nothing, "pBarTimePeriod is Nothing"
AssertArgument Not pSession Is Nothing Or pBarTimePeriod.Length = 0, "pSession is Nothing"

Dim lBarsBuilder As New BarsBuilder
lBarsBuilder.Initialise pBarTimePeriod, pSession, pTickSize, pNumberOfBarsToCache, pNormaliseDailyTimestamps, pSave
Set CreateBarsBuilder = lBarsBuilder

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarsBuilderFuture( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pSessionFuture As IFuture, _
                Optional ByVal pTickSize As Double, _
                Optional ByVal pNumberOfBarsToCache As Long, _
                Optional ByVal pNormaliseDailyTimestamps As Boolean, _
                Optional ByVal pSave As Boolean = True) As IFuture
Const ProcName As String = "CreateBarsBuilderFuture"
On Error GoTo Err

AssertArgument Not pBarTimePeriod Is Nothing, "pBarTimePeriod is Nothing"
AssertArgument Not pSessionFuture Is Nothing, "pSessionFuture is Nothing"

Dim lFutureBuilder As New BarsBuilderFutureBuilder
lFutureBuilder.Initialise pBarTimePeriod, pSessionFuture, pTickSize, pNumberOfBarsToCache, pNormaliseDailyTimestamps, pSave
Set CreateBarsBuilderFuture = lFutureBuilder.Future

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarsBuilderWithInitialBars( _
                ByVal pBars As Bars, _
                ByVal pSession As Session, _
                Optional ByVal pTickSize As Double) As BarsBuilder
Const ProcName As String = "CreateBarsBuilderWithInitialBars"
On Error GoTo Err

AssertArgument Not pBars Is Nothing, "pBars Is Nothing"
AssertArgument Not pSession Is Nothing, "pSession is Nothing"
    
Dim lBarsBuilder As New BarsBuilder
lBarsBuilder.InitialiseWithInitialBars pBars, pSession, pTickSize
Set CreateBarsBuilderWithInitialBars = lBarsBuilder

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarsBuilderWithInitialBarsFuture( _
                ByVal pBarsFuture As IFuture, _
                ByVal pSession As Session, _
                Optional ByVal pTickSize As Double) As IFuture
Const ProcName As String = "CreateBarsBuilderWithInitialBarsFuture"
On Error GoTo Err

AssertArgument Not pBarsFuture Is Nothing, "pBarsFuture Is Nothing"
AssertArgument Not pSession Is Nothing, "pSession is Nothing"
    
Dim lFutureBuilder As New BarsBuilderFutureBuilder
lFutureBuilder.InitialiseWithInitialBars pBarsFuture, pSession, pTickSize
Set CreateBarsBuilderWithInitialBarsFuture = lFutureBuilder.Future

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateValueCache( _
                ByVal CyclicSize As Long, _
                ByVal ValueName As String) As ValueCache
Const ProcName As String = "CreateValueCache"
On Error GoTo Err

Set CreateValueCache = New ValueCache
CreateValueCache.Initialise CyclicSize, ValueName

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateVolumeParser( _
                ByVal pSession As Session) As VolumeParser
Const ProcName As String = "CreateVolumeParser"
On Error GoTo Err

Set CreateVolumeParser = New VolumeParser
CreateVolumeParser.Initialise pSession

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub GetTimespanData( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByRef pFromTime As Date, _
                ByRef pToTime As Date, _
                ByRef pFromSessionTimes As SessionTimes, _
                ByRef pToSessionTimes As SessionTimes, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date)
Const ProcName As String = "GetTimespanData"
On Error GoTo Err

AssertArgument pFromTime <> 0, "pFromTime must be supplied"
AssertArgument pBarTimePeriod.Length <> 0, "pBarTimePeriod.Length is 0"
        
If pToTime = 0 Then pToTime = Now

Dim lStartTime As Date
lStartTime = GBarUtils.centiSecondsToDate(GBarUtils.calcBarStartTime(pFromTime, pBarTimePeriod, pSessionStartTime, pSessionEndTime))

pFromSessionTimes = GetSessionTimes(lStartTime, pSessionStartTime, pSessionEndTime)
If lStartTime > pFromSessionTimes.EndTime Then
    pFromSessionTimes = GetOffsetSessionTimes(lStartTime, 1, pSessionStartTime, pSessionEndTime)
    lStartTime = pFromSessionTimes.StartTime
End If

Dim lEndTime As Date
lEndTime = GBarUtils.centiSecondsToDate(GBarUtils.calcBarStartTime(pToTime, pBarTimePeriod, pSessionStartTime, pSessionEndTime))
If lEndTime < pToTime Then
    lEndTime = GBarUtils.centiSecondsToDate(GBarUtils.calcBarEndTime(pToTime, pBarTimePeriod, pSessionStartTime, pSessionEndTime))
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
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function MaxNumberOfBarsInTimespan( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pStartTime As Date, _
                Optional ByVal pEndTime As Date, _
                Optional ByVal pSessionStartTime As Date, _
                Optional ByVal pSessionEndTime As Date) As Long
Const ProcName As String = "MaxNumberOfBarsInTimespan"
On Error GoTo Err

Dim lFromSessionTimes As SessionTimes
Dim lToSessionTimes As SessionTimes
GBarUtils.GetTimespanData pBarTimePeriod, _
                pStartTime, _
                pEndTime, _
                lFromSessionTimes, _
                lToSessionTimes, _
                pSessionStartTime, _
                pSessionEndTime

MaxNumberOfBarsInTimespan = GBarUtils.MaxNumberOfBarsInTimespanNormalized( _
                                                        pBarTimePeriod, _
                                                        pStartTime, _
                                                        pEndTime, _
                                                        lFromSessionTimes, _
                                                        lToSessionTimes)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function MaxNumberOfBarsInTimespanNormalized( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pStartTime As Date, _
                ByVal pEndTime As Date, _
                ByRef pStartSessionTimes As SessionTimes, _
                ByRef pEndSessionTimes As SessionTimes) As Long
Const ProcName As String = "MaxNumberOfBarsInTimespanNormalized"
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
                            GBarUtils.calcNumberOfBarsInSession(pBarTimePeriod, pStartSessionTimes.StartTime, pStartSessionTimes.EndTime)
    End If
End If

MaxNumberOfBarsInTimespanNormalized = lNumberOfBars

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function NormaliseSessionTime( _
            ByVal Timestamp As Date) As Date
If CDbl(Timestamp) = 1# Then
    NormaliseSessionTime = 1#
Else
    NormaliseSessionTime = CDate(Round(86400# * (CDbl(Timestamp) - CDbl(Int(Timestamp)))) / 86400#)
End If
End Function

Public Function NormaliseTimestamp( _
                ByVal pTimestamp As Date, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
Select Case pTimePeriod.Units
Case TimePeriodDay, TimePeriodWeek, TimePeriodMonth, TimePeriodYear
    If pSessionStartTime > pSessionEndTime And _
        pTimestamp >= pSessionStartTime _
    Then
        NormaliseTimestamp = Int(pTimestamp) + 1
    Else
        NormaliseTimestamp = Int(pTimestamp)
    End If
Case Else
    NormaliseTimestamp = pTimestamp
End Select
End Function

Public Function NumberOfBarsInSession( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Long
Const ProcName As String = "NumberOfBarsInSession"
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case Else
    AssertArgument False, "Can't calculate number of Bars in session for this time unit"
End Select

NumberOfBarsInSession = GBarUtils.calcNumberOfBarsInSession( _
                                                BarTimePeriod, _
                                                GBarUtils.NormaliseSessionTime(SessionStartTime), _
                                                GBarUtils.NormaliseSessionTime(SessionEndTime))

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OffsetBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                Optional ByVal SessionStartTime As Date, _
                Optional ByVal SessionEndTime As Date) As Date
Const ProcName As String = "OffsetBarStartTime"
On Error GoTo Err

AssertArgument BarTimePeriod.Units <> TimePeriodNone, "Invalid time Units argument"

OffsetBarStartTime = GBarUtils.calcOffsetBarStartTime( _
                                Timestamp, _
                                BarTimePeriod, _
                                offset, _
                                GBarUtils.NormaliseSessionTime(SessionStartTime), _
                                GBarUtils.NormaliseSessionTime(SessionEndTime))

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcBarEndTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Currency
Const ProcName As String = "calcBarEndTime"
On Error GoTo Err

Dim startTimeSecs As Currency
startTimeSecs = GBarUtils.calcBarStartTime(Timestamp, BarTimePeriod, SessionStartTime, SessionEndTime)

Select Case BarTimePeriod.Units
Case TimePeriodSecond, TimePeriodMinute, TimePeriodHour
    calcBarEndTime = startTimeSecs + GBarUtils.calcBarLengthSeconds(BarTimePeriod)
Case TimePeriodDay
    calcBarEndTime = GBarUtils.dateToCentiSeconds(GetOffsetSessionTimes(GBarUtils.centiSecondsToDate(startTimeSecs), BarTimePeriod.Length, SessionStartTime, SessionEndTime).StartTime)
Case TimePeriodWeek
    calcBarEndTime = GBarUtils.dateToCentiSeconds(GBarUtils.centiSecondsToDate(startTimeSecs) + 7 * BarTimePeriod.Length)
Case TimePeriodMonth
    Dim lMonth As Long
    lMonth = Month(GBarUtils.centiSecondsToDate(startTimeSecs))
    If SessionStartTime > SessionEndTime Then
        lMonth = lMonth + 1
    End If
    Dim lBarEnd As Date
    lBarEnd = Int(MonthStartDateFromMonthNumber(lMonth + BarTimePeriod.Length, GBarUtils.centiSecondsToDate(startTimeSecs)))
    If SessionStartTime > SessionEndTime Then lBarEnd = lBarEnd - 1
    lBarEnd = lBarEnd + SessionStartTime
    calcBarEndTime = GBarUtils.dateToCentiSeconds(lBarEnd)
Case TimePeriodYear
    calcBarEndTime = GBarUtils.dateToCentiSeconds(DateAdd("yyyy", BarTimePeriod.Length, GBarUtils.centiSecondsToDate(startTimeSecs)))
Case TimePeriodVolume, _
        TimePeriodTickVolume, _
        TimePeriodTickMovement, _
        TimePeriodNone
    calcBarEndTime = GBarUtils.dateToCentiSeconds(Timestamp)
End Select

Select Case BarTimePeriod.Units
    Case TimePeriodSecond, _
            TimePeriodMinute, _
            TimePeriodHour
        
        ' adjust if bar is at end of session but does not fit exactly into session
        
        Dim lNextSessionStartTimeSecs As Currency
        lNextSessionStartTimeSecs = GBarUtils.dateToCentiSeconds(GetSessionTimesIgnoringWeekend(GBarUtils.centiSecondsToDate(startTimeSecs + OneDayCentiSecs), SessionStartTime, SessionEndTime).StartTime)
        If calcBarEndTime > lNextSessionStartTimeSecs Then calcBarEndTime = lNextSessionStartTimeSecs
End Select

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcBarLengthSeconds( _
                ByVal BarTimePeriod As TimePeriod) As Currency
Const ProcName As String = "calcBarLengthSeconds"
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodSecond
    calcBarLengthSeconds = BarTimePeriod.Length
Case TimePeriodMinute
    calcBarLengthSeconds = BarTimePeriod.Length * 60
Case TimePeriodHour
    calcBarLengthSeconds = BarTimePeriod.Length * 3600
Case TimePeriodDay
    calcBarLengthSeconds = BarTimePeriod.Length * 86400
Case Else
    AssertArgument False, "Invalid BarTimePeriod"
End Select

calcBarLengthSeconds = calcBarLengthSeconds * 100 ' convert to centiseconds

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Currency
Const ProcName As String = "calcBarStartTime"
On Error GoTo Err

AssertArgument SessionStartTime < 1#, "Session start time must be a time value only"
AssertArgument SessionEndTime <= 1#, "Session end time must be a time value only"

' seconds from midnight to start of sesssion
Dim sessionStartSecs As Currency
sessionStartSecs = GBarUtils.dateToCentiSeconds(SessionStartTime)

If SessionEndTime = 0 Then SessionEndTime = 1#
Dim sessionSpansMidnight As Boolean
sessionSpansMidnight = (SessionStartTime > SessionEndTime)

Dim dayNum As Long: dayNum = Int(CDbl(Timestamp))

Dim dayNumSecs As Currency
dayNumSecs = GBarUtils.dateToCentiSeconds(dayNum)

' NB: don't use TimeValue to get the time, as VB rounds it to
' the nearest second
Dim timeSecs As Currency
timeSecs = GBarUtils.dateToCentiSeconds(Timestamp) - dayNumSecs  ' seconds since midnight

Select Case BarTimePeriod.Units
Case TimePeriodSecond, TimePeriodMinute, TimePeriodHour
    Dim barLengthSecs As Currency: barLengthSecs = GBarUtils.calcBarLengthSeconds(BarTimePeriod)
    
    If timeSecs < sessionStartSecs Then
        dayNumSecs = dayNumSecs - OneDayCentiSecs
        timeSecs = timeSecs + OneDayCentiSecs
    End If
    calcBarStartTime = dayNumSecs + _
                    sessionStartSecs + _
                    barLengthSecs * Int((timeSecs - sessionStartSecs) / barLengthSecs)
Case TimePeriodDay
    If BarTimePeriod.Length = 1 Then
        Dim lSessionTimes As SessionTimes
        #If SingleDll Then
        calcBarStartTime = GBarUtils.dateToCentiSeconds(GSessionUtils.GetSessionTimes(Timestamp, SessionStartTime, SessionEndTime).StartTime)
        #Else
        calcBarStartTime = GBarUtils.dateToCentiSeconds(SessionUtils27.GetSessionTimes(Timestamp, SessionStartTime, SessionEndTime).StartTime)
        #End If
    Else
        Dim workingDayNum As Long
        workingDayNum = WorkingDayNumber(dayNum)

        calcBarStartTime = GBarUtils.dateToCentiSeconds(WorkingDayDate(1 + BarTimePeriod.Length * Int((workingDayNum - 1) / BarTimePeriod.Length), _
                                            dayNum)) + sessionStartSecs
        If timeSecs < sessionStartSecs Then calcBarStartTime = calcBarStartTime - OneDayCentiSecs
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
    calcBarStartTime = GBarUtils.dateToCentiSeconds(WeekStartDateFromWeekNumber(1 + BarTimePeriod.Length * Int((weekNum - 1) / BarTimePeriod.Length), _
                                        dayNum) + SessionStartTime)
    If sessionSpansMidnight Then calcBarStartTime = calcBarStartTime - OneDayCentiSecs
Case TimePeriodMonth
    If sessionSpansMidnight And timeSecs >= sessionStartSecs Then dayNum = dayNum + 1
    Dim monthNum As Long: monthNum = Month(dayNum)
    calcBarStartTime = GBarUtils.dateToCentiSeconds(MonthStartDateFromMonthNumber( _
                                            1 + BarTimePeriod.Length * Int((monthNum - 1) / BarTimePeriod.Length), _
                                            dayNum) + _
                                        SessionStartTime)
    If sessionSpansMidnight Then calcBarStartTime = calcBarStartTime - OneDayCentiSecs
Case TimePeriodYear
    calcBarStartTime = GBarUtils.dateToCentiSeconds(CDate(DateSerial(1900 + BarTimePeriod.Length * Int((Year(Timestamp) - 1900) / BarTimePeriod.Length), 1, 1)))
Case TimePeriodVolume, _
        TimePeriodTickVolume, _
        TimePeriodTickMovement, _
        TimePeriodNone
    calcBarStartTime = GBarUtils.dateToCentiSeconds(Timestamp)
End Select

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcNumberOfBarsInSession( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Long
Const ProcName As String = "calcNumberOfBarsInSession"
On Error GoTo Err

SessionStartTime = SessionStartTime - Int(SessionStartTime)
SessionEndTime = SessionEndTime - Int(SessionEndTime)
If SessionEndTime > SessionStartTime Then
    calcNumberOfBarsInSession = -Int(-(GBarUtils.dateToCentiSeconds(SessionEndTime) - GBarUtils.dateToCentiSeconds(SessionStartTime)) / GBarUtils.calcBarLengthSeconds(BarTimePeriod))
Else
    calcNumberOfBarsInSession = -Int(-(OneDayCentiSecs + GBarUtils.dateToCentiSeconds(SessionEndTime) - GBarUtils.dateToCentiSeconds(SessionStartTime)) / GBarUtils.calcBarLengthSeconds(BarTimePeriod))
End If

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

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

calcNumberOfBarsInTimespan = Int((GBarUtils.dateToCentiSeconds(pEndTime - pStartTime) + lBarLengthCentiSecs - 1) / lBarLengthCentiSecs)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcOffsetBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Date
Const ProcName As String = "calcOffsetBarStartTime"
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
    calcOffsetBarStartTime = calcOffsetDailyBarStartTime( _
                                                    Timestamp, _
                                                    BarTimePeriod, _
                                                    offset, _
                                                    SessionStartTime, _
                                                    SessionEndTime)
    Exit Function
Case TimePeriodWeek
    calcOffsetBarStartTime = calcOffsetWeeklyBarStartTime( _
                                                    Timestamp, _
                                                    BarTimePeriod, _
                                                    offset, _
                                                    SessionStartTime, _
                                                    SessionEndTime)
    Exit Function
Case TimePeriodMonth
    calcOffsetBarStartTime = calcOffsetMonthlyBarStartTime( _
                                                    Timestamp, _
                                                    BarTimePeriod, _
                                                    offset, _
                                                    SessionStartTime, _
                                                    SessionEndTime)
    Exit Function
Case TimePeriodYear
    calcOffsetBarStartTime = calcOffsetYearlyBarStartTime( _
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
barLengthSecs = GBarUtils.calcBarLengthSeconds(BarTimePeriod)

Dim numBarsInSession As Long
numBarsInSession = GBarUtils.calcNumberOfBarsInSession(BarTimePeriod, SessionStartTime, SessionEndTime)

Dim datumBarStartSecs As Currency
datumBarStartSecs = GBarUtils.calcBarStartTime(Timestamp, BarTimePeriod, SessionStartTime, SessionEndTime)

Dim lSessTimes As SessionTimes
lSessTimes = GetSessionTimes(Timestamp, _
                    SessionStartTime, _
                    SessionEndTime)

Dim proposedStartSecs As Currency
Dim remainingOffset As Long
If offset > 0 Then
    If datumBarStartSecs > GBarUtils.dateToCentiSeconds(lSessTimes.EndTime) Then
        ' specified Timestamp was between sessions
        lSessTimes = GetOffsetSessionTimes(GBarUtils.centiSecondsToDate(datumBarStartSecs), 1, SessionStartTime, SessionEndTime)
        datumBarStartSecs = GBarUtils.dateToCentiSeconds(lSessTimes.StartTime)
        offset = offset - 1
    End If
    
    Dim BarsToSessEnd As Long
    BarsToSessEnd = Int((GBarUtils.dateToCentiSeconds(lSessTimes.EndTime) - datumBarStartSecs) / barLengthSecs)
    If BarsToSessEnd > offset Then
        ' all required Bars are in this session
        proposedStartSecs = datumBarStartSecs + offset * barLengthSecs
    Else
        remainingOffset = offset - BarsToSessEnd
        proposedStartSecs = GBarUtils.dateToCentiSeconds(lSessTimes.EndTime)
        Do While remainingOffset >= 0
            lSessTimes = GetOffsetSessionTimes(GBarUtils.centiSecondsToDate(proposedStartSecs), 1, SessionStartTime, SessionEndTime)
            
            If numBarsInSession > remainingOffset Then
                proposedStartSecs = GBarUtils.dateToCentiSeconds(lSessTimes.StartTime) + remainingOffset * barLengthSecs
                remainingOffset = -1
            Else
                proposedStartSecs = GBarUtils.dateToCentiSeconds(lSessTimes.EndTime)
                remainingOffset = remainingOffset - numBarsInSession
            End If
        Loop
    End If
Else
    offset = -offset

    If datumBarStartSecs > GBarUtils.dateToCentiSeconds(lSessTimes.EndTime) Then
        ' specified Timestamp was between sessions
        datumBarStartSecs = GBarUtils.dateToCentiSeconds(lSessTimes.EndTime)
    End If
    
    proposedStartSecs = GBarUtils.dateToCentiSeconds(lSessTimes.StartTime)
    
    Dim BarsFromSessStart As Long
    BarsFromSessStart = Int((datumBarStartSecs - GBarUtils.dateToCentiSeconds(lSessTimes.StartTime)) / barLengthSecs)
    If BarsFromSessStart >= offset Then
        ' all required Bars are in this session
        proposedStartSecs = GBarUtils.dateToCentiSeconds(lSessTimes.StartTime) + (BarsFromSessStart - offset) * barLengthSecs
    Else
        remainingOffset = offset - BarsFromSessStart
        Do While remainingOffset > 0
            lSessTimes = GetSessionTimes(GBarUtils.centiSecondsToDate(proposedStartSecs - OneDayCentiSecs), SessionStartTime, SessionEndTime)
            
            If numBarsInSession >= remainingOffset Then
                proposedStartSecs = GBarUtils.dateToCentiSeconds(lSessTimes.EndTime) - remainingOffset * barLengthSecs
                remainingOffset = 0
            Else
                proposedStartSecs = GBarUtils.dateToCentiSeconds(lSessTimes.StartTime)
                remainingOffset = remainingOffset - numBarsInSession
            End If
        Loop
    End If
End If

calcOffsetBarStartTime = GBarUtils.centiSecondsToDate(proposedStartSecs)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
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
GBars.HandleUnexpectedError ProcName, ModuleName
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
datumBarStart = GBarUtils.centiSecondsToDate(GBarUtils.calcBarStartTime(Timestamp, BarTimePeriod, SessionStartTime, SessionEndTime))

calcOffsetMonthlyBarStartTime = DateAdd("m", offset * BarTimePeriod.Length, datumBarStart)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
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
datumBarStart = GBarUtils.centiSecondsToDate(GBarUtils.calcBarStartTime(Timestamp, BarTimePeriod, SessionStartTime, SessionEndTime))

Dim datumWeekNumber As Long
datumWeekNumber = DatePart("ww", datumBarStart, vbMonday, vbFirstFullWeek)

Dim proposedWeekNumber As Long
proposedWeekNumber = datumWeekNumber + offset * BarTimePeriod.Length

calcOffsetWeeklyBarStartTime = WeekStartDateFromWeekNumber(proposedWeekNumber, datumBarStart)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
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
datumBarStart = GBarUtils.centiSecondsToDate(GBarUtils.calcBarStartTime(Timestamp, GetTimePeriod(barLength, TimePeriodYear), SessionStartTime, SessionEndTime))

calcOffsetYearlyBarStartTime = DateAdd("yyyy", offset * barLength, datumBarStart)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function centiSecondsToDate(ByVal pSeconds As Currency) As Date
centiSecondsToDate = CDate(CDbl(pSeconds) / 8640000)
End Function

' returns the date in units of 100th of a second. Expressed as a Currency data type,
' this gives a resolution of one microsecond.
Private Function dateToCentiSeconds(ByVal pDate As Date) As Currency
dateToCentiSeconds = CCur(CDbl(pDate) * 8640000)
End Function



