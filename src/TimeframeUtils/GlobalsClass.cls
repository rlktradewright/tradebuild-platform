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

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

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


