VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalsClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Const ProjectName               As String = "TimeframeUtils26"
Private Const ModuleName                As String = "GlobalsClass"

Private Const TimePeriodNameSecond As String = "Second"
Private Const TimePeriodNameMinute As String = "Minute"
Private Const TimePeriodNameHour As String = "Hourly"
Private Const TimePeriodNameDay As String = "Daily"
Private Const TimePeriodNameWeek As String = "Weekly"
Private Const TimePeriodNameMonth As String = "Monthly"
Private Const TimePeriodNameYear As String = "Yearly"

Private Const TimePeriodNameSeconds As String = "Seconds"
Private Const TimePeriodNameMinutes As String = "Minutes"
Private Const TimePeriodNameHours As String = "Hours"
Private Const TimePeriodNameDays As String = "Days"
Private Const TimePeriodNameWeeks As String = "Weeks"
Private Const TimePeriodNameMonths As String = "Months"
Private Const TimePeriodNameYears As String = "Years"
Private Const TimePeriodNameVolumeIncrement As String = "Volume"
Private Const TimePeriodNameTickVolumeIncrement As String = "Tick Volume"
Private Const TimePeriodNameTickIncrement As String = "Ticks Movement"

'================================================================================
' Enums
'================================================================================

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

'================================================================================
' Types
'================================================================================

Public Type TimePeriod
    length          As Long
    units           As TimePeriodUnits
End Type

'================================================================================
' Member variables
'================================================================================

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

'================================================================================
' Methods
'================================================================================

Public Function BarEndTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal units As TimePeriodUnits, _
                Optional ByVal sessionStartTime As Date, _
                Optional ByVal sessionEndTime As Date) As Date
BarEndTime = gBarEndTime(timestamp, _
                            barLength, _
                            units, _
                            gNormaliseTime(sessionStartTime), _
                            gNormaliseTime(sessionEndTime))
End Function

Public Function barLength( _
                ByVal length As Long, _
                ByVal units As TimePeriodUnits) As Date
Select Case units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "StudiesSupp" & "." & "GlobalsClass" & ":" & "BarLength", _
            "Can't calculate bar length for this time unit"
End Select
barLength = gCalcBarLength(length, units)
End Function

Public Function BarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal units As TimePeriodUnits, _
                Optional ByVal sessionStartTime As Date) As Date
BarStartTime = gBarStartTime(timestamp, _
                            barLength, _
                            units, _
                            gNormaliseTime(sessionStartTime))
End Function

Public Function MonthStartDate( _
                ByVal monthNumber As Long, _
                ByVal baseDate As Date) As Date
MonthStartDate = gCalcMonthStartDate(monthNumber, baseDate)
End Function


Public Function NumberOfBarsInSession( _
                ByVal barLength As Long, _
                ByVal units As TimePeriodUnits, _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date) As Long
Select Case units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "BarsSupport" & "." & "GlobalsClass" & ":" & "NumberOfBarsInSession", _
            "Can't calculate number of bars in session for this time unit"
End Select

NumberOfBarsInSession = gCalcNumberOfBarsInSession( _
                                                barLength, _
                                                units, _
                                                gNormaliseTime(sessionStartTime), _
                                                gNormaliseTime(sessionEndTime))
End Function

Public Function OffsetBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal units As TimePeriodUnits, _
                ByVal offset As Long, _
                Optional ByVal sessionStartTime As Date, _
                Optional ByVal sessionEndTime As Date) As Date
If units = TimePeriodNone Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "BarsSupport" & "." & "GlobalsClass" & ":" & "OffsetBarStartTime", _
            "Invalid time units argument"
End If
OffsetBarStartTime = gCalcOffsetBarStartTime( _
                                                timestamp, _
                                                barLength, _
                                                units, _
                                                offset, _
                                                gNormaliseTime(sessionStartTime), _
                                                gNormaliseTime(sessionEndTime))
End Function

Public Sub SessionTimes( _
                ByVal timestamp As Date, _
                ByRef sessionStartTime As Date, _
                ByRef sessionEndTime As Date, _
                Optional ByVal startTime As Date, _
                Optional ByVal endTime As Date)
gCalcSessionTimes timestamp, _
                            gNormaliseTime(startTime), _
                            gNormaliseTime(endTime), _
                            sessionStartTime, _
                            sessionEndTime
End Sub

Public Function TimePeriodFromString( _
                ByVal value As String) As TimePeriod
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

i = 1
Do While IsNumeric(Mid$(value, i, 1))
    i = i + 1
Loop

If i = 1 Then
    TimePeriodFromString.length = 1
Else
    TimePeriodFromString.length = CLng(Left$(value, i - 1))
End If

TimePeriodFromString.units = TimePeriodUnitsFromString( _
                                        Trim$(Right$(value, Len(value) - i + 1)))

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "TimePeriodFromString" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description


End Function

Public Function TimePeriodToString( _
                tp As TimePeriod) As String
Dim tokens(1) As String

Dim failpoint As Long
On Error GoTo Err

tokens(0) = tp.length

Select Case tp.units
Case TimePeriodSecond
    If tp.length = 1 Then
        tokens(1) = TimePeriodNameSecond
    Else
        tokens(1) = TimePeriodNameSeconds
    End If
Case TimePeriodMinute
    If tp.length = 1 Then
        tokens(1) = TimePeriodNameMinute
    Else
        tokens(1) = TimePeriodNameMinutes
    End If
Case TimePeriodHour
    If tp.length = 1 Then
        TimePeriodToString = TimePeriodNameHour
        Exit Function
    Else
        tokens(1) = TimePeriodNameHours
    End If
Case TimePeriodDay
    If tp.length = 1 Then
        TimePeriodToString = TimePeriodNameDay
        Exit Function
    Else
        tokens(1) = TimePeriodNameDays
    End If
Case TimePeriodWeek
    If tp.length = 1 Then
        TimePeriodToString = TimePeriodNameWeek
        Exit Function
    Else
        tokens(1) = TimePeriodNameWeeks
    End If
Case TimePeriodMonth
    If tp.length = 1 Then
        TimePeriodToString = TimePeriodNameMonth
        Exit Function
    Else
        tokens(1) = TimePeriodNameMonths
    End If
Case TimePeriodYear
    If tp.length = 1 Then
        TimePeriodToString = TimePeriodNameYear
        Exit Function
    Else
        tokens(1) = TimePeriodNameYears
    End If
Case TimePeriodVolume
    tokens(1) = TimePeriodNameVolumeIncrement
Case TimePeriodTickVolume
    tokens(1) = TimePeriodNameTickVolumeIncrement
Case TimePeriodTickMovement
    tokens(1) = TimePeriodNameTickIncrement
End Select

TimePeriodToString = Join(tokens, " ")

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "TimePeriodToString" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Function

Public Function TimePeriodUnitsFromString( _
                timeUnits As String) As TimePeriodUnits

Select Case UCase$(timeUnits)
Case UCase$(TimePeriodNameSecond), UCase$(TimePeriodNameSeconds), "SEC", "SECS", "S"
    TimePeriodUnitsFromString = TimePeriodSecond
Case UCase$(TimePeriodNameMinute), UCase$(TimePeriodNameMinutes), "MIN", "MINS", "M"
    TimePeriodUnitsFromString = TimePeriodMinute
Case UCase$(TimePeriodNameHour), UCase$(TimePeriodNameHours), "HR", "HRS", "H"
    TimePeriodUnitsFromString = TimePeriodHour
Case UCase$(TimePeriodNameDay), UCase$(TimePeriodNameDays), "D", "DY", "DYS"
    TimePeriodUnitsFromString = TimePeriodDay
Case UCase$(TimePeriodNameWeek), UCase$(TimePeriodNameWeeks), "W", "WK", "WKS"
    TimePeriodUnitsFromString = TimePeriodWeek
Case UCase$(TimePeriodNameMonth), UCase$(TimePeriodNameMonths), "MTH", "MNTH", "MTHS", "MNTHS", "MM"
    TimePeriodUnitsFromString = TimePeriodMonth
Case UCase$(TimePeriodNameYear), UCase$(TimePeriodNameYears), "YR", "YRS", "Y", "YY", "YS"
    TimePeriodUnitsFromString = TimePeriodYear
Case UCase$(TimePeriodNameVolumeIncrement), "VOL", "V"
    TimePeriodUnitsFromString = TimePeriodVolume
Case UCase$(TimePeriodNameTickVolumeIncrement), "TICKVOL", "TICK VOL", "TICKVOLUME", "TV"
    TimePeriodUnitsFromString = TimePeriodTickVolume
Case UCase$(TimePeriodNameTickIncrement), "TICK", "TICKS", "TCK", "TCKS", "TM", "TICKSMOVEMENT", "TICKMOVEMENT"
    TimePeriodUnitsFromString = TimePeriodTickMovement
Case Else
    TimePeriodUnitsFromString = TimePeriodNone
End Select
End Function

Public Function TimePeriodUnitsToString( _
                timeUnits As TimePeriodUnits) As String

Select Case timeUnits
Case TimePeriodSecond
    TimePeriodUnitsToString = TimePeriodNameSeconds
Case TimePeriodMinute
    TimePeriodUnitsToString = TimePeriodNameMinutes
Case TimePeriodHour
    TimePeriodUnitsToString = TimePeriodNameHours
Case TimePeriodDay
    TimePeriodUnitsToString = TimePeriodNameDays
Case TimePeriodWeek
    TimePeriodUnitsToString = TimePeriodNameWeeks
Case TimePeriodMonth
    TimePeriodUnitsToString = TimePeriodNameMonths
Case TimePeriodYear
    TimePeriodUnitsToString = TimePeriodNameYears
Case TimePeriodVolume
    TimePeriodUnitsToString = TimePeriodNameVolumeIncrement
Case TimePeriodTickVolume
    TimePeriodUnitsToString = TimePeriodNameTickVolumeIncrement
Case TimePeriodTickMovement
    TimePeriodUnitsToString = TimePeriodNameTickIncrement
End Select
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


'================================================================================
' Helper Functions
'================================================================================


