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

Public Enum BarTypes
        BarTypeTrade = 0
        BarTypeBid = 1
        BarTypeAsk = 2
End Enum

''
' These values identify different units for measuring elapsed time.
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
'   The unit is a number of price ticks. The corresponding time is the
'   period during which price moves in a range of the specified number of ticks.
' @param TimePeriodTickVolume
'   The unit is a trade tick. The corresponding time is the
'   period during which accumulated tick volume increases by the specified number.
' @param TimePeriodVolume
'   The unit is a contract or share traded. The corresponding time is the
'   period during which accumulated volume increases by the specified number.
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

'Public Type Bar
'    timestamp       As Date
'    openPrice       As Double
'    highPrice       As Double
'    lowPrice        As Double
'    closePrice      As Double
'    volume          As Long
'    tickVolume      As Long
'    openInterest    As Long
'End Type

Public Type SValue
    value           As Variant
    barNumber       As Long
    timestamp       As Date
    BarStartTime    As Date
End Type

Public Type SessionTimes
    startTime       As Date
    endTime         As Date
End Type

'Public Type TimePeriod
'    length          As Long
'    units           As TimePeriodUnits
'End Type

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
                ByVal timestamp As Date, _
                ByVal barTimePeriod As TimePeriod, _
                Optional ByVal sessionStartTime As Date, _
                Optional ByVal sessionEndTime As Date) As Date
BarEndTime = gBarEndTime(timestamp, _
                            barTimePeriod, _
                            gNormaliseTime(sessionStartTime), _
                            gNormaliseTime(sessionEndTime))
End Function

Public Function barLength( _
                ByVal barTimePeriod As TimePeriod) As Date
Select Case barTimePeriod.units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "barLength", _
            "Can't calculate bar length for this time unit"
End Select
barLength = gCalcBarLength(barTimePeriod)
End Function

Public Function BarStartTime( _
                ByVal timestamp As Date, _
                ByVal barTimePeriod As TimePeriod, _
                Optional ByVal sessionStartTime As Date) As Date
BarStartTime = gBarStartTime(timestamp, _
                            barTimePeriod, _
                            gNormaliseTime(sessionStartTime))
End Function

Public Function CreateBar( _
                ByVal timestamp As Date, _
                ByVal openValue As Double, _
                ByVal highValue As Double, _
                ByVal lowValue As Double, _
                ByVal closeValue As Double, _
                Optional ByVal volume As Long, _
                Optional ByVal tickVolume As Long, _
                Optional ByVal openInterest As Long) As Bar
Set CreateBar = New Bar
CreateBar.initialise timestamp, _
                    openValue, _
                    highValue, _
                    lowValue, _
                    closeValue, _
                    volume, _
                    tickVolume, _
                    openInterest
End Function

Public Function CreateBarsBuilder( _
                ByVal barTimePeriod As TimePeriod, _
                Optional ByVal pSession As Session, _
                Optional ByVal tickSize As Double, _
                Optional ByVal numberOfBarsToCache As Long) As BarsBuilder
Set CreateBarsBuilder = New BarsBuilder
CreateBarsBuilder.initialise barTimePeriod, _
                            pSession, tickSize, numberOfBarsToCache
End Function

Public Function GetTimePeriod( _
                ByVal length As Long, _
                ByVal units As TimePeriodUnits) As TimePeriod

Dim failpoint As Long
On Error GoTo Err

Set GetTimePeriod = gGetTimePeriod(length, units)

Exit Function

Err:
GetLogger("").Log LogLevelSevere, "Error at: " & ProjectName & "." & ModuleName & ":" & "GetTimePeriod" & "." & failpoint & _
                            IIf(Err.source <> "", vbCrLf & Err.source, "") & vbCrLf & _
                            Err.Description
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "GetTimePeriod" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description


End Function

Public Function CreateValueCache( _
                ByVal cyclicSize As Long, _
                ByVal valueName As String) As ValueCache
Set CreateValueCache = New ValueCache
CreateValueCache.initialise cyclicSize, valueName
End Function

Public Function CreateVolumeParser( _
                ByVal pSession As Session) As VolumeParser
Set CreateVolumeParser = New VolumeParser
CreateVolumeParser.initialise pSession
End Function

Public Function GetSessionTimes( _
                ByVal timestamp As Date, _
                Optional ByVal startTime As Date, _
                Optional ByVal endTime As Date) As SessionTimes
Dim st As SessionTimes
gCalcSessionTimes timestamp, _
                            gNormaliseTime(startTime), _
                            gNormaliseTime(endTime), _
                            st.startTime, _
                            st.endTime
GetSessionTimes = st
End Function

Public Function GetOffsetSessionTimes( _
                ByVal timestamp As Date, _
                ByVal offset As Long, _
                Optional ByVal startTime As Date, _
                Optional ByVal endTime As Date) As SessionTimes
Dim st As SessionTimes
gCalcOffsetSessionTimes timestamp, _
                            offset, _
                            gNormaliseTime(startTime), _
                            gNormaliseTime(endTime), _
                            st.startTime, _
                            st.endTime
GetOffsetSessionTimes = st
End Function

Public Function MonthStartDate( _
                ByVal monthNumber As Long, _
                ByVal baseDate As Date) As Date
MonthStartDate = gCalcMonthStartDate(monthNumber, baseDate)
End Function


Public Function NumberOfBarsInSession( _
                ByVal barTimePeriod As TimePeriod, _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date) As Long
Select Case barTimePeriod.units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "NumberOfBarsInSession", _
            "Can't calculate number of bars in session for this time unit"
End Select

NumberOfBarsInSession = gCalcNumberOfBarsInSession( _
                                                barTimePeriod, _
                                                gNormaliseTime(sessionStartTime), _
                                                gNormaliseTime(sessionEndTime))
End Function

Public Function OffsetBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                Optional ByVal sessionStartTime As Date, _
                Optional ByVal sessionEndTime As Date) As Date
If barTimePeriod.units = TimePeriodNone Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "OffsetBarStartTime", _
            "Invalid time units argument"
End If
OffsetBarStartTime = gCalcOffsetBarStartTime( _
                                                timestamp, _
                                                barTimePeriod, _
                                                offset, _
                                                gNormaliseTime(sessionStartTime), _
                                                gNormaliseTime(sessionEndTime))
End Function

Public Function TimePeriodFromString( _
                ByVal value As String) As TimePeriod
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

value = Trim$(value)

i = 1
Do While IsNumeric(Mid$(value, i, 1))
    i = i + 1
Loop

Set TimePeriodFromString = gGetTimePeriod(IIf(i = 1, 1, CLng(Left$(value, i - 1))), _
                                        TimePeriodUnitsFromString(Trim$(Right$(value, Len(value) - i + 1))))
Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "TimePeriodFromString" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description


End Function

Public Function TimePeriodUnitsFromString( _
                timeUnits As String) As TimePeriodUnits
TimePeriodUnitsFromString = gTimePeriodUnitsFromString(timeUnits)
End Function

Public Function TimePeriodUnitsToString( _
                timeUnits As TimePeriodUnits) As String
TimePeriodUnitsToString = gTimePeriodUnitsToString(timeUnits)
End Function

Public Function TimePeriodUnitsToShortString( _
                timeUnits As TimePeriodUnits) As String
TimePeriodUnitsToShortString = gTimePeriodUnitsToShortString(timeUnits)
End Function

Public Function WeekStartDate( _
                ByVal weekNumber As Long, _
                ByVal baseDate As Date) As Date
WeekStartDate = gCalcWeekStartDate(weekNumber, baseDate)
End Function

Public Function WorkingDayDate( _
                ByVal dayNumber As Long, _
                ByVal baseDate As Date) As Date
WorkingDayDate = gCalcWorkingDayDate(dayNumber, baseDate)
End Function

Public Function WorkingDayNumber( _
                ByVal pDate As Date) As Long
WorkingDayNumber = gCalcWorkingDayNumber(pDate)
End Function


'@================================================================================
' Helper Functions
'@================================================================================


