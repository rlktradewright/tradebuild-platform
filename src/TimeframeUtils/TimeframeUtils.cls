VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeframeUtils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Const ProjectName               As String = "TimeframeUtils26"
Private Const ModuleName                As String = "GlobalsClass"

'@================================================================================
' Enums
'@================================================================================

Public Enum BarChangeTypes
    BarPriceChanged
    BarVolumeChanged
    BarOpenInterestChanged
    BarTickVolumeChanged    ' only fired when TickVolume is specifically incremented
                            ' BarPriceChanged implicitly includes BarTickVolumeChanged
End Enum

Public Enum BarTypes
    BarTypeTrade = 0
    BarTypeBid = 1
    BarTypeAsk = 2
End Enum

Public Enum SessionChangeTypes
    SessionChangeNone
    SessionChangeEnd
    SessionChangeStart
End Enum

''
' These values identify different Units for measuring elapsed time.
'
' @param TimePeriodNone
'   The unit is not specified.
' @param TimePeriodSecond
'   The unit is in seconds.
' @param TimePeriodMinute
'   The unit is in minutes
' @param TimePeriodHour
'   The unit is in hours
' @param TimePeriodDay
'   The unit is in days.
' @param TimePeriodWeek
'   The unit is in weeks.
' @param TimePeriodMonth
'   The unit is in months.
' @param TimePeriodYear
'   The unit is in years.
' @param TimePeriodTickMovement
'   The unit is a number of price Ticks. The corresponding time is the
'   period during which price moves in a range of the specified number of Ticks.
' @param TimePeriodTickVolume
'   The unit is a trade Tick. The corresponding time is the
'   period during which accumulated Tick Volume increases by the specified number.
' @param TimePeriodVolume
'   The unit is a contract or share traded. The corresponding time is the
'   period during which accumulated Volume increases by the specified number.
'@/
Public Enum TimePeriodUnits
    TimePeriodNone = 0
    TimePeriodSecond = 1
    TimePeriodMinute = 2
    TimePeriodHour = 3
    TimePeriodDay = 4
    TimePeriodWeek = 5
    TimePeriodMonth = 6
'    TimePeriodLunarMonth = 7
    TimePeriodYear = 8
    TimePeriodTickMovement = 9     ' for constant momentum, point & figure, etc
    TimePeriodTickVolume = 10
    TimePeriodVolume = 11
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type SessionEvent
    source              As Object
    Timestamp           As Date
    changeType          As SessionChangeTypes
End Type

Public Type SessionTimes
    startTime       As Date
    endTime         As Date
End Type

Public Type SValue
    value           As Variant
    BarNumber       As Long
    Timestamp       As Date
    BarStartTime    As Date
End Type

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

Public Function BarEndTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                Optional ByVal SessionStartTime As Date, _
                Optional ByVal SessionEndTime As Date) As Date
Const ProcName As String = "BarEndTime"
Dim failpoint As String
On Error GoTo Err

BarEndTime = gBarEndTime(Timestamp, _
                            BarTimePeriod, _
                            gNormaliseTime(SessionStartTime), _
                            gNormaliseTime(SessionEndTime))

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function BarLength( _
                ByVal BarTimePeriod As TimePeriod) As Date
Const ProcName As String = "BarLength"
Dim failpoint As String
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Can't calculate Bar length for this time unit"
End Select
BarLength = gCalcBarLength(BarTimePeriod)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function BarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                Optional ByVal SessionStartTime As Date) As Date
Const ProcName As String = "BarStartTime"
Dim failpoint As String
On Error GoTo Err

BarStartTime = gBarStartTime(Timestamp, _
                            BarTimePeriod, _
                            gNormaliseTime(SessionStartTime))

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function CreateBar( _
                ByVal Timestamp As Date, _
                ByVal OpenValue As Double, _
                ByVal HighValue As Double, _
                ByVal LowValue As Double, _
                ByVal CloseValue As Double, _
                Optional ByVal Volume As Long, _
                Optional ByVal TickVolume As Long, _
                Optional ByVal OpenInterest As Long) As Bar
Const ProcName As String = "CreateBar"
Dim failpoint As String
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function CreateBarsBuilder( _
                ByVal BarTimePeriod As TimePeriod, _
                Optional ByVal pSession As Session, _
                Optional ByVal TickSize As Double, _
                Optional ByVal numberOfBarsToCache As Long) As BarsBuilder
Set CreateBarsBuilder = New BarsBuilder
CreateBarsBuilder.Initialise BarTimePeriod, _
                            pSession, TickSize, numberOfBarsToCache
End Function

Public Function GetTimePeriod( _
                ByVal Length As Long, _
                ByVal Units As TimePeriodUnits) As TimePeriod
Const ProcName As String = "GetTimePeriod"
Dim failpoint As String
On Error GoTo Err

Set GetTimePeriod = gGetTimePeriod(Length, Units)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function CreateValueCache( _
                ByVal CyclicSize As Long, _
                ByVal ValueName As String) As ValueCache
Const ProcName As String = "CreateValueCache"
Dim failpoint As String
On Error GoTo Err

Set CreateValueCache = New ValueCache
CreateValueCache.Initialise CyclicSize, ValueName

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function CreateVolumeParser( _
                ByVal pSession As Session) As VolumeParser
Const ProcName As String = "CreateVolumeParser"
Dim failpoint As String
On Error GoTo Err

Set CreateVolumeParser = New VolumeParser
CreateVolumeParser.Initialise pSession

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function GetSessionTimes( _
                ByVal Timestamp As Date, _
                Optional ByVal startTime As Date, _
                Optional ByVal endTime As Date) As SessionTimes
Dim st As SessionTimes
Const ProcName As String = "GetSessionTimes"
Dim failpoint As String
On Error GoTo Err

gCalcSessionTimes Timestamp, _
                            gNormaliseTime(startTime), _
                            gNormaliseTime(endTime), _
                            st.startTime, _
                            st.endTime
GetSessionTimes = st

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function GetOffsetSessionTimes( _
                ByVal Timestamp As Date, _
                ByVal offset As Long, _
                Optional ByVal startTime As Date, _
                Optional ByVal endTime As Date) As SessionTimes
Dim st As SessionTimes
Const ProcName As String = "GetOffsetSessionTimes"
Dim failpoint As String
On Error GoTo Err

gCalcOffsetSessionTimes Timestamp, _
                            offset, _
                            gNormaliseTime(startTime), _
                            gNormaliseTime(endTime), _
                            st.startTime, _
                            st.endTime
GetOffsetSessionTimes = st

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function MonthStartDate( _
                ByVal monthNumber As Long, _
                ByVal baseDate As Date) As Date
Const ProcName As String = "MonthStartDate"
Dim failpoint As String
On Error GoTo Err

MonthStartDate = gCalcMonthStartDate(monthNumber, baseDate)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function


Public Function NumberOfBarsInSession( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Long
Const ProcName As String = "NumberOfBarsInSession"
Dim failpoint As String
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Can't calculate number of Bars in session for this time unit"
End Select

NumberOfBarsInSession = gCalcNumberOfBarsInSession( _
                                                BarTimePeriod, _
                                                gNormaliseTime(SessionStartTime), _
                                                gNormaliseTime(SessionEndTime))

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function OffsetBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                Optional ByVal SessionStartTime As Date, _
                Optional ByVal SessionEndTime As Date) As Date
Const ProcName As String = "OffsetBarStartTime"
Dim failpoint As String
On Error GoTo Err

If BarTimePeriod.Units = TimePeriodNone Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid time Units argument"
End If
OffsetBarStartTime = gCalcOffsetBarStartTime( _
                                                Timestamp, _
                                                BarTimePeriod, _
                                                offset, _
                                                gNormaliseTime(SessionStartTime), _
                                                gNormaliseTime(SessionEndTime))

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function TimePeriodFromString( _
                ByVal value As String) As TimePeriod
Dim i As Long
Dim Length As Long

Const ProcName As String = "TimePeriodFromString"
Dim failpoint As String
On Error GoTo Err

value = Trim$(value)

i = 1
Do While IsNumeric(Mid$(value, i, 1))
    i = i + 1
Loop

If i = 1 Then
    Length = 1
Else
    Length = CLng(Left$(value, i - 1))
End If
    
Set TimePeriodFromString = gGetTimePeriod(Length, _
                                        TimePeriodUnitsFromString(Trim$(Right$(value, Len(value) - i + 1))))

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function TimePeriodUnitsFromString( _
                timeUnits As String) As TimePeriodUnits
Const ProcName As String = "TimePeriodUnitsFromString"
Dim failpoint As String
On Error GoTo Err

TimePeriodUnitsFromString = gTimePeriodUnitsFromString(timeUnits)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function TimePeriodUnitsToString( _
                timeUnits As TimePeriodUnits) As String
Const ProcName As String = "TimePeriodUnitsToString"
Dim failpoint As String
On Error GoTo Err

TimePeriodUnitsToString = gTimePeriodUnitsToString(timeUnits)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function TimePeriodUnitsToShortString( _
                timeUnits As TimePeriodUnits) As String
Const ProcName As String = "TimePeriodUnitsToShortString"
Dim failpoint As String
On Error GoTo Err

TimePeriodUnitsToShortString = gTimePeriodUnitsToShortString(timeUnits)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function WeekStartDate( _
                ByVal weekNumber As Long, _
                ByVal baseDate As Date) As Date
Const ProcName As String = "WeekStartDate"
Dim failpoint As String
On Error GoTo Err

WeekStartDate = gCalcWeekStartDate(weekNumber, baseDate)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function WorkingDayDate( _
                ByVal dayNumber As Long, _
                ByVal baseDate As Date) As Date
Const ProcName As String = "WorkingDayDate"
Dim failpoint As String
On Error GoTo Err

WorkingDayDate = gCalcWorkingDayDate(dayNumber, baseDate)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function WorkingDayNumber( _
                ByVal pDate As Date) As Long
Const ProcName As String = "WorkingDayNumber"
Dim failpoint As String
On Error GoTo Err

WorkingDayNumber = gCalcWorkingDayNumber(pDate)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function


'@================================================================================
' Helper Functions
'@================================================================================


