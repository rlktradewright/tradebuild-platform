Attribute VB_Name = "Globals"
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

Private Const ProjectName                           As String = "SessionUtils27"
Private Const ModuleName                            As String = "Globals"

Public Const OneSecond                              As Double = 1 / 86400
Public Const OneMinute                              As Double = 1 / 1440
Public Const OneHour                                As Double = 1 / 24

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
Static sLogger As FormattingLogger
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("sessionutils", ProjectName)
Set gLogger = sLogger
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gCalcOffsetSessionTimes( _
                ByVal pTimestamp As Date, _
                ByVal pOffset As Long, _
                ByVal pStartTime As Date, _
                ByVal pEndTime As Date, _
                ByVal pIgnoreWeekends As Boolean) As SessionTimes
Const ProcName As String = "gCalcOffsetSessionTimes"
On Error GoTo Err

Dim lDatumSessionTimes As SessionTimes

Dim lTargetDate As Date

If pIgnoreWeekends Then
    lDatumSessionTimes = gGetSessionTimesIgnoringWeekend(pTimestamp, pStartTime, pEndTime)
    lTargetDate = DateValue(lDatumSessionTimes.StartTime) + pOffset
Else
    lDatumSessionTimes = gCalcSessionTimes(pTimestamp, pStartTime, pEndTime)
    
    Dim lBasedate As Date
    lBasedate = lDatumSessionTimes.EndTime - IIf(pEndTime = 1#, 1, 0)
    
    Dim lTargetWorkingDayNum As Long
    lTargetWorkingDayNum = WorkingDayNumber(lBasedate) + pOffset

    lTargetDate = WorkingDayDate(lTargetWorkingDayNum, Int(lBasedate))
    If sessionSpansMidnight(pStartTime, pEndTime) Then lTargetDate = lTargetDate - 1
End If

With gCalcOffsetSessionTimes
    .StartTime = lTargetDate + pStartTime
    .EndTime = lTargetDate + pEndTime
    If sessionSpansMidnight(pStartTime, pEndTime) Then .EndTime = .EndTime + 1
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

' !!!!!!!!!!!!!!!!!!!!!!!!!!!
' calcSessionTimes needs to be amended
' to take account of holidays
Public Function gCalcSessionTimes( _
                ByVal pTimestamp As Date, _
                ByVal pStartTime As Date, _
                ByVal pEndTime As Date) As SessionTimes
Const ProcName As String = "gCalcSessionTimes"
On Error GoTo Err

gCalcSessionTimes = gGetSessionTimesIgnoringWeekend(pTimestamp, _
                        pStartTime, _
                        pEndTime)

Dim lWeekday As VbDayOfWeek
lWeekday = DatePart("w", gCalcSessionTimes.StartTime)
If sessionSpansMidnight(pStartTime, pEndTime) Then
    ' session DOES span midnight
    If lWeekday = vbFriday Then
        gCalcSessionTimes.StartTime = gCalcSessionTimes.StartTime - 1
        gCalcSessionTimes.EndTime = gCalcSessionTimes.EndTime - 1
    ElseIf lWeekday = vbSaturday Then
        gCalcSessionTimes.StartTime = gCalcSessionTimes.StartTime - 2
        gCalcSessionTimes.EndTime = gCalcSessionTimes.EndTime - 2
    End If
Else
    ' session doesn't span midnight or 24-hour session or no session times known
    If lWeekday = vbSunday Then
        gCalcSessionTimes.StartTime = gCalcSessionTimes.StartTime - 2
        gCalcSessionTimes.EndTime = gCalcSessionTimes.EndTime - 2
    ElseIf lWeekday = vbSaturday Then
        gCalcSessionTimes.StartTime = gCalcSessionTimes.StartTime - 1
        gCalcSessionTimes.EndTime = gCalcSessionTimes.EndTime - 1
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetSessionTimesIgnoringWeekend( _
                ByVal Timestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As SessionTimes
Const ProcName As String = "gGetSessionTimesIgnoringWeekend"
On Error GoTo Err

Dim referenceDate As Date
referenceDate = DateValue(Timestamp)

Dim referenceTime As Date
referenceTime = TimeValue(Timestamp)

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

gGetSessionTimesIgnoringWeekend.StartTime = sessionStartDate + pSessionStartTime
If pSessionEndTime > pSessionStartTime Then
    gGetSessionTimesIgnoringWeekend.EndTime = sessionStartDate + pSessionEndTime
Else
    gGetSessionTimesIgnoringWeekend.EndTime = sessionStartDate + 1 + pSessionEndTime
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Public Function gNormaliseSessionEndTime( _
            ByVal Timestamp As Date) As Date
gNormaliseSessionEndTime = CDate(Round(86400# * (CDbl(Timestamp) - CDbl(Int(Timestamp)))) / 86400#)
If gNormaliseSessionEndTime = 0# Then gNormaliseSessionEndTime = 1#
End Function

Public Function gNormaliseSessionStartTime( _
            ByVal Timestamp As Date) As Date
gNormaliseSessionStartTime = CDate(Round(86400# * (CDbl(Timestamp) - CDbl(Int(Timestamp)))) / 86400#)
End Function

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

'@================================================================================
' Helper Functions
'@================================================================================

Private Function sessionSpansMidnight( _
                ByVal pStartTime As Date, _
                ByVal pEndTime As Date) As Boolean
sessionSpansMidnight = (pStartTime > pEndTime)
End Function




