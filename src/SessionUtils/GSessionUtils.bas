Attribute VB_Name = "GSessionUtils"
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GSessionUtils"

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

Public Function CalcOffsetSessionTimes( _
                ByVal pTimestamp As Date, _
                ByVal pOffset As Long, _
                ByVal pStartTime As Date, _
                ByVal pEndTime As Date, _
                ByVal pIgnoreWeekends As Boolean) As SessionTimes
Const ProcName As String = "CalcOffsetSessionTimes"
On Error GoTo Err

Dim lDatumSessionTimes As SessionTimes

Dim lTargetDate As Date

If pIgnoreWeekends Then
    lDatumSessionTimes = GSessionUtils.GetSessionTimesIgnoringWeekend(pTimestamp, pStartTime, pEndTime)
    lTargetDate = DateValue(lDatumSessionTimes.StartTime) + pOffset
Else
    lDatumSessionTimes = GSessionUtils.CalcSessionTimes(pTimestamp, pStartTime, pEndTime)
    
    Dim lBasedate As Date
    lBasedate = lDatumSessionTimes.EndTime - IIf(pEndTime = 1#, 1, 0)
    
    Dim lTargetWorkingDayNum As Long
    lTargetWorkingDayNum = WorkingDayNumber(lBasedate) + pOffset

    lTargetDate = WorkingDayDate(lTargetWorkingDayNum, Int(lBasedate))
    If sessionSpansMidnight(pStartTime, pEndTime) Then lTargetDate = lTargetDate - 1
End If

With CalcOffsetSessionTimes
    .StartTime = lTargetDate + pStartTime
    .EndTime = lTargetDate + pEndTime
    If sessionSpansMidnight(pStartTime, pEndTime) Then .EndTime = .EndTime + 1
End With

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

' !!!!!!!!!!!!!!!!!!!!!!!!!!!
' calcSessionTimes needs to be amended
' to take account of holidays
Public Function CalcSessionTimes( _
                ByVal pTimestamp As Date, _
                ByVal pStartTime As Date, _
                ByVal pEndTime As Date) As SessionTimes
Const ProcName As String = "CalcSessionTimes"
On Error GoTo Err

CalcSessionTimes = GSessionUtils.GetSessionTimesIgnoringWeekend(pTimestamp, _
                        pStartTime, _
                        pEndTime)

Dim lWeekday As VbDayOfWeek
lWeekday = DatePart("w", CalcSessionTimes.StartTime)
If sessionSpansMidnight(pStartTime, pEndTime) Then
    ' session DOES span midnight
    If lWeekday = vbFriday Then
        CalcSessionTimes.StartTime = CalcSessionTimes.StartTime - 1
        CalcSessionTimes.EndTime = CalcSessionTimes.EndTime - 1
    ElseIf lWeekday = vbSaturday Then
        CalcSessionTimes.StartTime = CalcSessionTimes.StartTime - 2
        CalcSessionTimes.EndTime = CalcSessionTimes.EndTime - 2
    End If
Else
    ' session doesn't span midnight or 24-hour session or no session times known
    If lWeekday = vbSunday Then
        CalcSessionTimes.StartTime = CalcSessionTimes.StartTime - 2
        CalcSessionTimes.EndTime = CalcSessionTimes.EndTime - 2
    ElseIf lWeekday = vbSaturday Then
        CalcSessionTimes.StartTime = CalcSessionTimes.StartTime - 1
        CalcSessionTimes.EndTime = CalcSessionTimes.EndTime - 1
    End If
End If

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateSessionBuilder( _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                Optional ByVal pTimeZone As TimeZone, _
                Optional ByVal pInitialSessionTime As Date) As SessionBuilder
Const ProcName As String = "CreateSessionBuilder"
On Error GoTo Err

Set CreateSessionBuilder = New SessionBuilder
CreateSessionBuilder.Initialise pSessionStartTime, pSessionEndTime, pTimeZone
CreateSessionBuilder.SetSessionCurrentTime pInitialSessionTime

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateSessionBuilderFuture(ByVal pSessionFuture As IFuture) As IFuture
Const ProcName As String = "CreateSessionBuilderFuture"
On Error GoTo Err

Dim lBuilder As New SessionBuilderFutBldr
lBuilder.Initialise pSessionFuture
Set CreateSessionBuilderFuture = lBuilder.Future

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateSessionFuture(ByVal pSessionBuilderFuture As IFuture) As IFuture
Const ProcName As String = "CreateSessionFuture"
On Error GoTo Err

Dim lBuilder As New SessionFutureBuilder
lBuilder.Initialise pSessionBuilderFuture
Set CreateSessionFuture = lBuilder.Future

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetOffsetSessionTimes( _
                ByVal Timestamp As Date, _
                ByVal offset As Long, _
                Optional ByVal StartTime As Date, _
                Optional ByVal EndTime As Date) As SessionTimes
Const ProcName As String = "GetOffsetSessionTimes"
On Error GoTo Err

GetOffsetSessionTimes = GSessionUtils.CalcOffsetSessionTimes(Timestamp, _
                            offset, _
                            GSessionUtils.NormaliseSessionStartTime(StartTime), _
                            GSessionUtils.NormaliseSessionEndTime(EndTime), _
                            False)

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetOffsetSessionTimesIgnoringWeekend( _
                ByVal Timestamp As Date, _
                ByVal offset As Long, _
                Optional ByVal StartTime As Date, _
                Optional ByVal EndTime As Date) As SessionTimes
Const ProcName As String = "GetOffsetSessionTimesIgnoringWeekend"
On Error GoTo Err

GetOffsetSessionTimesIgnoringWeekend = GSessionUtils.CalcOffsetSessionTimes(Timestamp, _
                            offset, _
                            GSessionUtils.NormaliseSessionStartTime(StartTime), _
                            GSessionUtils.NormaliseSessionEndTime(EndTime), _
                            True)

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetSessionTimes( _
                ByVal Timestamp As Date, _
                Optional ByVal StartTime As Date, _
                Optional ByVal EndTime As Date) As SessionTimes
Const ProcName As String = "GetSessionTimes"
On Error GoTo Err

GetSessionTimes = GSessionUtils.CalcSessionTimes(Timestamp, _
                            GSessionUtils.NormaliseSessionStartTime(StartTime), _
                            GSessionUtils.NormaliseSessionEndTime(EndTime))

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetSessionTimesIgnoringWeekend( _
                ByVal Timestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As SessionTimes
Const ProcName As String = "GetSessionTimesIgnoringWeekend"
On Error GoTo Err

Dim referenceDate As Date
referenceDate = DateValue(Timestamp)

Dim referenceTime As Date
referenceTime = TimeValue(Timestamp)

If referenceDate = MinDate Then
    GetSessionTimesIgnoringWeekend.StartTime = MinDate
    GetSessionTimesIgnoringWeekend.EndTime = MinDate
    Exit Function
End If

If referenceDate = MaxDate Then
    GetSessionTimesIgnoringWeekend.StartTime = MaxDate
    GetSessionTimesIgnoringWeekend.EndTime = MaxDate
    Exit Function
End If

Dim sessionStartDate As Date

If pSessionStartTime < pSessionEndTime Then
    If referenceTime < pSessionStartTime Then
        sessionStartDate = referenceDate - 1
    Else
        sessionStartDate = referenceDate
    End If
Else
    If referenceTime >= pSessionStartTime Then
        sessionStartDate = referenceDate
    Else
        sessionStartDate = referenceDate - 1
    End If
End If

GetSessionTimesIgnoringWeekend.StartTime = sessionStartDate + pSessionStartTime
If pSessionEndTime > pSessionStartTime Then
    GetSessionTimesIgnoringWeekend.EndTime = sessionStartDate + pSessionEndTime
Else
    GetSessionTimesIgnoringWeekend.EndTime = sessionStartDate + 1 + pSessionEndTime
End If

Exit Function

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function NormaliseSessionEndTime( _
            ByVal Timestamp As Date) As Date
NormaliseSessionEndTime = CDate(Round(86400# * (CDbl(Timestamp) - CDbl(Int(Timestamp)))) / 86400#)
If NormaliseSessionEndTime = 0# Then NormaliseSessionEndTime = 1#
End Function

Public Function NormaliseSessionStartTime( _
            ByVal Timestamp As Date) As Date
NormaliseSessionStartTime = CDate(Round(86400# * (CDbl(Timestamp) - CDbl(Int(Timestamp)))) / 86400#)
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function sessionSpansMidnight( _
                ByVal pStartTime As Date, _
                ByVal pEndTime As Date) As Boolean
sessionSpansMidnight = (pStartTime > pEndTime)
End Function






