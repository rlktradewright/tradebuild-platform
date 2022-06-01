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

Public Const ProjectName                            As String = "HistDataUtils27"
Private Const ModuleName                            As String = "Globals"

Public Const Latest                                 As String = "LATEST"
Public Const Today                                  As String = "TODAY"
Public Const Tomorrow                               As String = "TOMORROW"
Public Const Yesterday                              As String = "YESTERDAY"
Public Const EndOfWeek                              As String = "ENDOFWEEK"
Public Const StartOfWeek                            As String = "STARTOFWEEK"
Public Const StartOfPreviousWeek                    As String = "STARTOFPREVIOUSWEEK"

Public Const All                                    As Long = &H7FFFFFFF

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
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("histdatautils", ProjectName)
Set gLogger = sLogger
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gBarTypeToString(ByVal pBarType As BarTypes) As String
Select Case pBarType
Case BarTypeTrade
    gBarTypeToString = "TRADE"
Case BarTypeBid
    gBarTypeToString = "BID"
Case BarTypeAsk
    gBarTypeToString = "ASK"
Case Else
    AssertArgument False, "Invalid bar type"
End Select
End Function

Public Function gCreateBarDataSpecifier( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pToTime As Date, _
                ByVal pFromTime As Date, _
                ByVal pMaxNumberOfBars As Long, _
                ByVal pBarType As BarTypes, _
                ByVal pExcludeCurrentBar As Boolean, _
                ByVal pIncludeBarsOutsideSession As Boolean, _
                ByVal pNormaliseDailyTimestamps As Boolean, _
                ByVal pCustomSessionStartTime As Date, _
                ByVal pCustomSessionEndTime As Date) As BarDataSpecifier
Const ProcName As String = "gCreateBarDataSpecifier"
On Error GoTo Err

AssertArgument Not pBarTimePeriod Is Nothing

Dim lBarDataSpecifier As New BarDataSpecifier
lBarDataSpecifier.Initialise _
                            pBarTimePeriod, _
                            pToTime, _
                            pFromTime, _
                            pMaxNumberOfBars, _
                            pBarType, _
                            pExcludeCurrentBar, _
                            pIncludeBarsOutsideSession, _
                            pNormaliseDailyTimestamps, _
                            pCustomSessionStartTime, _
                            pCustomSessionEndTime
Set gCreateBarDataSpecifier = lBarDataSpecifier

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateBufferedBarWriter( _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pOutputMonitor As IBarOutputMonitor, _
                ByVal pContractFuture As IFuture) As IBarWriter
Const ProcName As String = "gCreateBufferedBarWriter"
On Error GoTo Err

Dim lBufferedWriter As New BufferedBarWriter
Dim lWriter As IBarWriter
Set lWriter = pHistDataStore.CreateBarWriter(lBufferedWriter, pContractFuture)
lBufferedWriter.Initialise pOutputMonitor, lWriter, pContractFuture
Set gCreateBufferedBarWriter = lBufferedWriter

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

Public Function gSpecialTimeToDate( _
                ByVal pSpecialTime As String, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                Optional ByVal pClock As Clock) As Date
Const ProcName As String = "gSpecialTimeToDate"
On Error GoTo Err

Dim lTimestamp As Date
If pClock Is Nothing Then
    lTimestamp = GetTimestamp
Else
    lTimestamp = pClock.Timestamp
End If

pSpecialTime = UCase$(pSpecialTime)

If pSpecialTime = Today Then
    gSpecialTimeToDate = todayDate(lTimestamp, pSessionStartTime, pSessionEndTime)
ElseIf pSpecialTime = Yesterday Then
    gSpecialTimeToDate = yesterdayDate(lTimestamp, pSessionStartTime, pSessionEndTime)
ElseIf pSpecialTime = StartOfWeek Then
    gSpecialTimeToDate = startOfWeekDate(lTimestamp, pSessionStartTime, pSessionEndTime)
ElseIf pSpecialTime = StartOfPreviousWeek Then
    gSpecialTimeToDate = startOfPreviousWeekDate(lTimestamp, pSessionStartTime, pSessionEndTime)
ElseIf pSpecialTime = Latest Then
    gSpecialTimeToDate = MaxDate
ElseIf pSpecialTime = Tomorrow Then
    gSpecialTimeToDate = tomorrowDate(lTimestamp, pSessionStartTime, pSessionEndTime)
ElseIf pSpecialTime = EndOfWeek Then
    gSpecialTimeToDate = endOfWeekDate(lTimestamp, pSessionStartTime, pSessionEndTime)
Else
    AssertArgument False, "Invalid special time '" & pSpecialTime & "'"
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function endOfWeekDate( _
                ByVal pTimestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
endOfWeekDate = GetSessionTimes( _
                 Int(pTimestamp) - DatePart("w", pTimestamp, vbMonday) + vbFriday - 1 + 0.5, _
                 pSessionStartTime, _
                 pSessionEndTime).StartTime
End Function

Private Function startOfPreviousWeekDate( _
                ByVal pTimestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
startOfPreviousWeekDate = GetSessionTimes( _
                Int(pTimestamp) - DatePart("w", pTimestamp, vbMonday) + vbMonday - 8 + 0.5, _
                pSessionStartTime, _
                pSessionEndTime).StartTime
End Function

Private Function startOfWeekDate( _
                ByVal pTimestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
startOfWeekDate = GetSessionTimes( _
                Int(pTimestamp) - DatePart("w", pTimestamp, vbMonday) + vbMonday - 1 + 0.5, _
                pSessionStartTime, _
                pSessionEndTime).StartTime
End Function

Private Function todayDate( _
                ByVal pTimestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
todayDate = GetSessionTimes( _
                Int(WorkingDayDate(WorkingDayNumber(pTimestamp), pTimestamp)) + 0.5, _
                pSessionStartTime, _
                pSessionEndTime).StartTime
End Function

Private Function tomorrowDate( _
                ByVal pTimestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
tomorrowDate = GetSessionTimes( _
                Int(WorkingDayDate(WorkingDayNumber(pTimestamp) + 1, pTimestamp)) + 0.5, _
                pSessionStartTime, _
                pSessionEndTime).StartTime
End Function

Private Function yesterdayDate( _
                ByVal pTimestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
yesterdayDate = GetSessionTimes( _
                Int(WorkingDayDate(WorkingDayNumber(pTimestamp) - 1, pTimestamp)) + 0.5, _
                pSessionStartTime, _
                pSessionEndTime).StartTime
End Function




