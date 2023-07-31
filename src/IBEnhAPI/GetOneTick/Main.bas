Attribute VB_Name = "MainMod"
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

Public Const ProjectName                            As String = "GetOneTick"
Private Const ModuleName                            As String = "Main"

Private Const TwsSwitch                             As String = "TWS"

Private Const DefaultClientId                       As Long = 476113245

'@================================================================================
' Member variables
'@================================================================================

Public gCon                                         As Console

Private mFatalErrorHandler                          As FatalErrorHandler

Private mClp                                        As CommandLineParser

Private mTwsClient                                  As Client
Private mClientId                                   As Long

Private mProcessor                                  As Processor

Private mTerminateRequested                         As Boolean

Private mContractStore                              As IContractStore
Private mMarketDataManager                          As RealTimeDataManager

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

Public Sub gFinish()
gWriteLineToConsole "Finish requested"
mTerminateRequested = True
End Sub

Public Sub gHandleFatalError(ev As ErrorEventData)
On Error Resume Next    ' ignore any further errors that might arise

If gCon Is Nothing Then Exit Sub

gCon.WriteErrorString "Error "
gCon.WriteErrorString CStr(ev.ErrorCode)
gCon.WriteErrorString ": "
gCon.WriteErrorLine ev.ErrorMessage
gCon.WriteErrorLine "At:"
gCon.WriteErrorLine ev.ErrorSource

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

Public Sub gWriteErrorLine( _
                ByVal pMessage As String)
Const ProcName As String = "gWriteErrorLine"

Dim s As String
s = "Error: " & pMessage
gCon.WriteErrorLine s
LogMessage "StdErr: " & s
End Sub

Public Sub gWriteLineToConsole( _
                ByVal pMessage As String, _
                Optional ByVal pIncludeTimestamp As Boolean, _
                Optional ByVal pDontLogit As Boolean)
Const ProcName As String = "gWriteLineToConsole"

If Not pDontLogit Then LogMessage "Con: " & pMessage
Dim lTime As String
If pIncludeTimestamp Then
    gCon.WriteStringToConsole FormatTimestamp(GetTimestamp, TimestampTimeOnlyISO8601) & " "
End If
gCon.WriteLineToConsole pMessage
End Sub

Public Sub gWriteLineToStdOut(ByVal pMessage As String)
Const ProcName As String = "gWriteLineToStdOut"

LogMessage "StdOut: " & pMessage
gCon.WriteLine pMessage
End Sub

Public Sub Main()
Const ProcName As String = "Main"
On Error GoTo Err

Set gCon = GetConsole

InitialiseTWUtilities

Set mFatalErrorHandler = New FatalErrorHandler
ApplicationGroupName = "TradeWright"
ApplicationName = "GetOneTick"
SetupDefaultLogging Command, True, True

logProgramId

Set mClp = CreateCommandLineParser(Command)

If Not setupTwsApi(mClp.SwitchValue(TwsSwitch), _
                mClientId) Then
    'showUsage
    Exit Sub
End If

Process

TerminateTWUtilities

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getInputLineAndWait( _
                Optional ByVal pDontReadInput As Boolean = False, _
                Optional ByVal pWaitTimeMIllisecs As Long = 5, _
                Optional ByVal pPrompt As String = ":") As String
Const ProcName As String = "getInputLine"
On Error GoTo Err

Dim lWaitUntilTime As Double
lWaitUntilTime = GetTimestampUTC + pWaitTimeMIllisecs / (86400# * 1000#)

If Not pDontReadInput Then getInputLineAndWait = Trim$(gCon.ReadLine(pPrompt))

Do
    ' allow queued system messages to be handled
    Wait 5
Loop Until GetTimestampUTC >= lWaitUntilTime

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub logProgramId()
Const ProcName As String = "logProgramId"
On Error GoTo Err

Dim s As String
s = App.ProductName & _
    " V" & _
    App.Major & _
    "." & App.Minor & _
    "." & App.Revision & _
    IIf(App.FileDescription <> "", "-" & App.FileDescription, "") & _
    vbCrLf & _
    App.LegalCopyright
gWriteLineToConsole s, , True
s = s & vbCrLf & "Arguments: " & Command
LogMessage s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Process()
Const ProcName As String = "process"
On Error GoTo Err

Dim lSpecString As String
lSpecString = mClp.Arg(0)

Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifierFromString(lSpecString)

Set mProcessor = New Processor
mProcessor.Process lContractSpec, _
                    mContractStore, _
                    mMarketDataManager

Dim lInputString As String

Dim lCounter As Long
Do Until lInputString = gCon.EofString Or _
        UCase$(lInputString) = "EXIT" Or _
        mTerminateRequested
    lCounter = lCounter + 1
    If lCounter Mod 10 = 0 Then gWriteLineToConsole "Waiting for tick"
    lInputString = getInputLineAndWait
Loop

If Not mTwsClient Is Nothing Then
    gWriteLineToConsole "Releasing API connection", True
    mTwsClient.Finish
    ' allow time for the socket connection to be nicely released
    getInputLineAndWait pDontReadInput:=True, pWaitTimeMIllisecs:=10
End If

gWriteLineToConsole "Exiting", True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setupTwsApi( _
                ByVal SwitchValue As String, _
                ByRef pClientId As Long) As Boolean
Const ProcName As String = "setupTwsApi"
On Error GoTo Err

setupTwsApi = True

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(SwitchValue, ",")

Dim server As String
server = lClp.Arg(0)

Dim port As String
port = lClp.Arg(1)
If port = "" Then
    port = 7496
ElseIf Not IsInteger(port, 1024) Then
    gWriteErrorLine "port must be an integer > 1024 and <= 65535"
    setupTwsApi = False
End If
    
Dim clientId As String
clientId = lClp.Arg(2)
If clientId = "" Then
    clientId = DefaultClientId
ElseIf Not IsInteger(clientId, 0, 999999999) Then
    gWriteErrorLine "clientId must be an integer >= 0 and <= 999999999"
    setupTwsApi = False
End If
pClientId = CLng(clientId)

Dim connectionRetryInterval As String
connectionRetryInterval = lClp.Arg(3)
If connectionRetryInterval = "" Then
ElseIf Not IsInteger(connectionRetryInterval, 0, 3600) Then
    gWriteErrorLine "Error: connection retry interval must be an integer >= 0 and <= 3600"
    setupTwsApi = False
End If

If Not setupTwsApi Then Exit Function

If connectionRetryInterval = "" Then
    Set mTwsClient = GetClient(server, _
                            CLng(port), _
                            pClientId)
Else
    Set mTwsClient = GetClient(server, _
                            CLng(port), _
                            pClientId, _
                            pConnectionRetryIntervalSecs:=CLng(connectionRetryInterval))
End If

Set mContractStore = mTwsClient.GetContractStore
Set mMarketDataManager = CreateRealtimeDataManager(mTwsClient.GetMarketDataFactory, mContractStore)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function





