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

Public Function gCreateBarDataSpecifierFuture( _
                ByVal pContractFuture As IFuture, _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pToTime As Date, _
                ByVal pFromTime As Date, _
                ByVal pMaxNumberOfBars As Long, _
                ByVal pBarType As BarTypes, _
                ByVal pClockFuture As IFuture, _
                ByVal pExcludeCurrentBar As Boolean, _
                ByVal pIncludeBarsOutsideSession As Boolean, _
                ByVal pNormaliseDailyTimestamps As Boolean, _
                ByVal pCustomSessionStartTime As Date, _
                ByVal pCustomSessionEndTime As Date) As IFuture
Const ProcName As String = "gCreateBarDataSpecifier"
On Error GoTo Err

AssertArgument Not pContractFuture Is Nothing
AssertArgument Not pBarTimePeriod Is Nothing

Dim lBarDataSpecifierFutureBuilder As New BarDataSpecFutureBldr
lBarDataSpecifierFutureBuilder.Initialise pContractFuture, _
                                            pBarTimePeriod, _
                                            pToTime, _
                                            pFromTime, _
                                            pMaxNumberOfBars, _
                                            pBarType, _
                                            pClockFuture, _
                                            pExcludeCurrentBar, _
                                            pIncludeBarsOutsideSession, _
                                            pNormaliseDailyTimestamps, _
                                            pCustomSessionStartTime, _
                                            pCustomSessionEndTime
Set gCreateBarDataSpecifierFuture = lBarDataSpecifierFutureBuilder.Future

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

'@================================================================================
' Helper Functions
'@================================================================================




