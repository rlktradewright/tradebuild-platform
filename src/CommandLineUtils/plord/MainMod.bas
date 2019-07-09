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

Public Const ProjectName                            As String = "plord27"
Private Const ModuleName                            As String = "MainMod"

Private Const LogFileSwitch                         As String = "LOG"
Private Const MonitorSwitch                         As String = "MONITOR"
Private Const ResultsDirSwitch                      As String = "RESULTSDIR"
Private Const RecoveryFileDirSwitch                 As String = "RECOVERYFILEDIR"
Private Const ScopeNameSwitch                       As String = "SCOPENAME"
Private Const SimulateOrdersSwitch                  As String = "SIMULATEORDERS"
Private Const StageOrdersSwitch                     As String = "STAGEORDERS"
Private Const StopAtSwitch                          As String = "STOPAT"
Private Const TwsSwitch                             As String = "TWS"

Public Const CancelAfterSwitch                      As String = "CANCELAFTER"
Public Const CancelPriceSwitch                      As String = "CANCELPRICE"
Public Const OffsetSwitch                           As String = "OFFSET"
Public Const PriceSwitch                            As String = "PRICE"
Public Const TIFSwitch                              As String = "TIF"
Public Const TrailBySwitch                          As String = "TRAILBY"
Public Const TrailPercentSwitch                     As String = "TRAILPERCENT"
Public Const TriggerPriceSwitch                     As String = "TRIGGER"
Public Const TriggerPriceSwitch1                    As String = "TRIGGERPRICE"

Public Const CurrencySwitch                         As String = "CURRENCY"
Public Const CurrencySwitch1                        As String = "CURR"
Public Const ExchangeSwitch                         As String = "EXCHANGE"
Public Const ExchangeSwitch1                        As String = "EXCH"
Public Const ExpirySwitch                           As String = "EXPIRY"
Public Const ExpirySwitch1                          As String = "EXP"
Public Const GoodAfterTimeSwitch                    As String = "GOODAFTERTIME"
Public Const GoodTillDateSwitch                     As String = "GOODTILLDATE"
Public Const LocalSymbolSwitch                      As String = "LOCALSYMBOL"
Public Const LocalSymbolSwitch1                     As String = "LOCAL"
Public Const MultiplierSwitch                       As String = "MULTIPLIER"
Public Const MultiplierSwitch1                      As String = "MULT"
Public Const RightSwitch                            As String = "RIGHT"
Public Const SecTypeSwitch                          As String = "SECTYPE"
Public Const SecTypeSwitch1                         As String = "SEC"
Public Const SymbolSwitch                           As String = "SYMBOL"
Public Const SymbolSwitch1                          As String = "SYMB"
Public Const StrikeSwitch                           As String = "STRIKE"
Public Const StrikeSwitch1                          As String = "STR"
Public Const TimezoneSwitch                         As String = "TIMEZONE"

Public Const BracketCommand                         As String = "BRACKET"
Public Const CloseoutCommand                        As String = "CLOSEOUT"
Public Const ContractCommand                        As String = "CONTRACT"
Public Const EndBracketCommand                      As String = "ENDBRACKET"
Public Const EndOrdersCommand                       As String = "ENDORDERS"
Public Const EntryCommand                           As String = "ENTRY"
Public Const ExitCommand                            As String = "EXIT"
Public Const GroupCommand                           As String = "GROUP"
Public Const HelpCommand                            As String = "HELP"
Public Const Help1Command                           As String = "?"
Public Const ListCommand                            As String = "LIST"
Public Const OrderCommand                           As String = "ORDER"
Public Const QuitCommand                            As String = "QUIT"
Public Const ResetCommand                           As String = "RESET"
Public Const StageOrdersCommand                     As String = "STAGEORDERS"
Public Const StopLossCommand                        As String = "STOPLOSS"
Public Const TargetCommand                          As String = "TARGET"

Public Const GroupsSubcommand                       As String = "GROUPS"
Public Const PositionsSubcommand                    As String = "POSITIONS"
Public Const TradesSubcommand                       As String = "TRADES"

Public Const ValueSeparator                         As String = ":"
Public Const SwitchPrefix                           As String = "/"

Private Const All                                   As String = "ALL"
Private Const Default                               As String = "DEFAULT"
Private Const Yes                                   As String = "YES"
Private Const No                                    As String = "NO"

Public Const TickDesignator                         As String = "T"

Private Const DefaultClientId                       As Long = 906564398

Private Const DefaultOrderGroupName                 As String = "$"

Private Const DefaultPrompt                         As String = ":"

'@================================================================================
' Member variables
'@================================================================================

Public gCon                                         As Console

Public gLogToConsole                                As Boolean

' This flag is set when a Contract command is being processed: this is because
' the contract fetch and starting the market data happen asynchronously, and we
' don't want any further user input until we know whether it has succeeded.
Public gInputPaused                                 As Boolean

Public gNumberOfOrdersPlaced                        As Long

Private mErrorCount                                 As Long

Private mFatalErrorHandler                          As FatalErrorHandler

Private mContractStore                              As IContractStore
Private mMarketDataManager                          As IMarketDataManager

Private mScopeName                                  As String
Private mOrderManager                               As New OrderManager
Private mOrderSubmitterFactory                      As IOrderSubmitterFactory

Private mGroupName                                  As String

Private mStageOrdersDefault                         As Boolean
Private mStageOrders                                As Boolean

Private mContractProcessor                          As ContractProcessor
Private mCloseoutProcessor                          As CloseoutProcessor

Private mContractProcessors                         As New EnumerableCollection

Private mCommandNumber                              As Long

Private mValidNextCommands()                        As String

Private mClp                                        As CommandLineParser

Private mMonitor                                    As Boolean

Private mLineNumber                                 As Long

Private mRecoveryFileDir                            As String

Private mOrderPersistenceDataStore                  As IOrderPersistenceDataStore
Private mOrderRecoveryAgent                         As IOrderRecoveryAgent

Private mConfigStore                                As ConfigurationStore

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
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("plord", ProjectName)
Set gLogger = sLogger
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gGenerateSwitch(ByVal pName As String, ByVal pValue As String) As String
If InStr(1, pValue, " ") <> 0 Then pValue = """" & pValue & """"
gGenerateSwitch = SwitchPrefix & pName & IIf(pValue = "", " ", ValueSeparator & pValue & " ")
End Function

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

Public Sub gSetValidNextCommands(ParamArray values() As Variant)
ReDim mValidNextCommands(UBound(values)) As String
Dim i As Long
For i = 0 To UBound(values)
    mValidNextCommands(i) = values(i)
Next
End Sub

Public Sub gWriteErrorLine(ByVal pMessage As String)
Const ProcName As String = "gWriteErrorLine"

Dim s As String
s = "Error on line " & mLineNumber & ": " & pMessage
gCon.WriteErrorLine s
gLogger.Log "StdErr: " & s, ProcName, ModuleName
mErrorCount = mErrorCount + 1
End Sub

Public Sub gWriteLineToConsole(ByVal pMessage As String)
Const ProcName As String = "gWriteLineToConsole"

gLogger.Log "Con: " & pMessage, ProcName, ModuleName
gCon.WriteLineToConsole pMessage
End Sub

Public Sub gWriteLineToStdOut(ByVal pMessage As String)
Const ProcName As String = "gWriteLineToStdOut"

gLogger.Log "StdOut: " & pMessage, ProcName, ModuleName
LogMessage pMessage
gCon.WriteLine pMessage
End Sub

Public Sub Main()
Const ProcName As String = "Main"
On Error GoTo Err

InitialiseTWUtilities

Set mFatalErrorHandler = New FatalErrorHandler
ApplicationGroupName = "TradeWright"
ApplicationName = "plord"
SetupDefaultLogging command, True, True

Set mConfigStore = gGetConfigStore

Set mCloseoutProcessor = New CloseoutProcessor
mCloseoutProcessor.Initialise mOrderManager

Set gCon = GetConsole

Set mClp = CreateCommandLineParser(command)

If Trim$(command) = "/?" Or Trim$(command) = "-?" Then
    showUsage
    Exit Sub
End If

If setupTws(mClp.SwitchValue(TwsSwitch), mClp.Switch(SimulateOrdersSwitch)) Then
    process
Else
    showUsage
    Exit Sub
End If

If mMonitor And gNumberOfOrdersPlaced > 0 Then
    Do While True
        Wait 10
    Loop
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getPrompt() As String
If mContractProcessor Is Nothing Then
    getPrompt = DefaultPrompt
ElseIf mContractProcessor.Contract Is Nothing Then
    getPrompt = DefaultPrompt
Else
    getPrompt = mContractProcessor.Contract.Specifier.LocalSymbol & _
                "@" & _
                mContractProcessor.Contract.Specifier.Exchange & _
                DefaultPrompt
End If
End Function

Private Function isCommandValid(ByVal pCommand As String) As Boolean
Dim i As Long
For i = 0 To UBound(mValidNextCommands)
    If pCommand = mValidNextCommands(i) Then
        isCommandValid = True
        Exit Function
    End If
Next
End Function

Private Sub listGroupNames()
Const ProcName As String = "listGroupNames"
On Error GoTo Err

Dim lVar As Variant
For Each lVar In mOrderManager.GetGroupNames
    Dim lGroupName As String: lGroupName = lVar
    gWriteLineToConsole IIf(lGroupName = mGroupName, "* ", "  ") & lGroupName
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub listPositions()
Const ProcName As String = "listPositions"
On Error GoTo Err

Dim lPM As PositionManager
For Each lPM In mOrderManager.PositionManagersLive
    If lPM.PositionSize <> 0 Or lPM.PendingPositionSize <> 0 Then
        Dim lContract As IContract
        Set lContract = lPM.ContractFuture.Value
        gWriteLineToConsole padStringRight(lPM.GroupName, 15) & " " & _
                            padStringRight(lContract.Specifier.LocalSymbol & "@" & lContract.Specifier.Exchange, 25) & _
                            " Size=" & padStringleft(lPM.PositionSize, 5) & _
                            IIf(lPM.PendingPositionSize <> 0, " PendingSize=" & padStringleft(lPM.PendingPositionSize, 5), "") & _
                            " Profit=" & padStringleft(Format(lPM.Profit, "0.00"), 8)
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub listTrades()
Const ProcName As String = "listTrades"
On Error GoTo Err

Dim lPM As PositionManager
For Each lPM In mOrderManager.PositionManagersLive
    Dim lContract As IContract
    Set lContract = lPM.ContractFuture.Value
    gWriteLineToConsole padStringRight(lPM.GroupName, 15) & " " & _
                        padStringRight(lContract.Specifier.LocalSymbol & "@" & lContract.Specifier.Exchange, 25)
    
    Dim lTrade As Execution
    For Each lTrade In lPM.Executions
        gWriteLineToConsole "  " & FormatTimestamp(lTrade.FillTime, TimestampTimeOnlyISO8601 + TimestampNoMillisecs) & " " & _
        padStringRight(IIf(lTrade.Action = OrderActionBuy, "BUY", "SELL"), 5) & _
        padStringleft(lTrade.Quantity, 5) & _
        padStringleft(FormatPrice(lTrade.Price, lContract.Specifier.SecType, lContract.TickSize), 9) & _
        " " & lTrade.TimezoneName
    Next
Next


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function padStringRight(ByVal pInput As String, ByVal pLength As Long) As String
pInput = Left$(pInput, pLength)
padStringRight = pInput & Space$(pLength - Len(pInput))
End Function

Public Function padStringleft(ByVal pInput As String, ByVal pLength As Long) As String
pInput = Right$(pInput, pLength)
padStringleft = Space$(pLength - Len(pInput)) & pInput
End Function

Private Sub process()
Const ProcName As String = "process"
On Error GoTo Err

If Not setMonitor Then Exit Sub
If Not setStageOrders Then Exit Sub
setOrderRecovery

gSetValidNextCommands ListCommand, StageOrdersCommand, GroupCommand, ContractCommand, CloseoutCommand

Dim inString As String
inString = Trim$(gCon.ReadLine(DefaultPrompt))

Do While inString <> gCon.EofString
    mLineNumber = mLineNumber + 1
    
    If inString = "" Then
        ' ignore blank lines - and don't write to the log because
        ' the FileAutoReader program sends blank lines very frequently
    ElseIf Left$(inString, 1) = "#" Then
        gLogger.Log "StdIn: " & inString, ProcName, ModuleName
        ' ignore comments
    Else
        gLogger.Log "StdIn: " & inString, ProcName, ModuleName
        If Not processCommand(inString) Then Exit Do
    End If
    
    Do While gInputPaused
        Wait 20
    Loop
    inString = Trim$(gCon.ReadLine(getPrompt))
Loop

If Not mOrderPersistenceDataStore Is Nothing Then mOrderPersistenceDataStore.Finish
Set mOrderPersistenceDataStore = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processCommand(ByVal pInstring As String) As Boolean
Const ProcName As String = "processCommand"
On Error GoTo Err

Dim command As String
command = UCase$(Split(pInstring, " ")(0))

Dim params As String
params = Trim$(Right$(pInstring, Len(pInstring) - Len(command)))

If command = ExitCommand Or command = QuitCommand Then
    gWriteLineToConsole "Exiting"
    processCommand = False
    Exit Function
End If

If command = Help1Command Then
    gCon.WriteLine "Valid commands at this point are: " & Join(mValidNextCommands, ",")
ElseIf command = HelpCommand Then
    showStdInHelp
ElseIf Not isCommandValid(command) Then
    gWriteErrorLine "Valid commands at this point are: " & Join(mValidNextCommands, ",")
Else
    Select Case command
    Case ContractCommand
        Set mContractProcessor = Nothing
        Set mContractProcessor = processContractCommand(params)
    Case StageOrdersCommand
        processStageOrdersCommand params
    Case GroupCommand
        processGroupCommand params
    Case BracketCommand
        mContractProcessor.ProcessBracketCommand params
    Case EntryCommand
        mContractProcessor.ProcessEntryCommand params
    Case StopLossCommand
        mContractProcessor.ProcessStopLossCommand params
    Case TargetCommand
        mContractProcessor.ProcessTargetCommand params
    Case EndBracketCommand
        mContractProcessor.ProcessEndBracketCommand
    Case EndOrdersCommand
        processEndOrdersCommand
    Case ResetCommand
        processResetCommand
    Case ListCommand
        processListCommand params
    Case CloseoutCommand
        processCloseoutCommand params
    Case Else
        gWriteErrorLine "Invalid command '" & command & "'"
    End Select
End If

processCommand = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processCloseoutCommand( _
                ByVal pParams As String)
Const ProcName As String = "processCloseoutCommand"
On Error GoTo Err

If UCase$(pParams) = All Then
    mCloseoutProcessor.CloseoutAll
    gInputPaused = True
ElseIf UCase$(pParams) <> "" Then
    mCloseoutProcessor.CloseoutGroup UCase$(pParams)
    gInputPaused = True
Else
    gWriteErrorLine CloseoutCommand & " parameter must be a group name or ALL"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processContractCommand(ByVal pParams As String) As ContractProcessor
Const ProcName As String = "processContractCommand"
On Error GoTo Err

Dim lProcessor As New ContractProcessor
lProcessor.StageOrders = mStageOrders
mContractProcessors.Add lProcessor
        
lProcessor.Initialise mContractStore, mMarketDataManager, mOrderManager, mScopeName, mGroupName, mOrderSubmitterFactory

gInputPaused = True
If Not lProcessor.processContractCommand(pParams) Then
    gInputPaused = False
    Set lProcessor = Nothing
End If

Set processContractCommand = lProcessor

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processEndOrdersCommand()
Const ProcName As String = "processEndOrdersCommand"
On Error GoTo Err

If mErrorCount <> 0 Then
    gWriteErrorLine mErrorCount & " errors have been found - no orders will be placed"
    mErrorCount = 0
    If Not mContractProcessor Is Nothing Then
        gSetValidNextCommands ListCommand, GroupCommand, BracketCommand, ResetCommand, CloseoutCommand
    Else
        gSetValidNextCommands ListCommand, GroupCommand, ContractCommand, ResetCommand, CloseoutCommand
    End If
    Exit Sub
End If

' if there have been no errors, we need to tell each ContractProcessor to submit
' its orders. To avoid exceeding the API's input message limits, we do this
' asynchronously with a task, which means we need to inhibit user input
' until it has completed.
gInputPaused = True

Dim t As New PlaceOrdersTask
t.Initialise mContractProcessors, mStageOrders, mMonitor
StartTask t, PriorityNormal

setupResultsLogging mClp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processGroupCommand( _
                ByVal pParams As String)
mGroupName = pParams
End Sub

Private Sub processListCommand( _
                ByVal pParams As String)
Const ProcName As String = "processListCommand"
On Error GoTo Err

If UCase$(pParams) = GroupsSubcommand Then
    listGroupNames
ElseIf UCase$(pParams) = PositionsSubcommand Then
    listPositions
ElseIf UCase$(pParams) = TradesSubcommand Then
    listTrades
Else
    gWriteErrorLine ListCommand & " parameter must be one of " & GroupsSubcommand & ", " & PositionsSubcommand & " or " & TradesSubcommand
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processResetCommand()
mContractProcessors.Clear
Set mContractProcessor = Nothing
mStageOrdersDefault = mStageOrdersDefault
mErrorCount = 0
gSetValidNextCommands ListCommand, StageOrdersCommand, GroupCommand, ContractCommand, ResetCommand, CloseoutCommand
End Sub

Private Sub processStageOrdersCommand( _
                ByVal pParams As String)
Select Case UCase$(pParams)
Case Default
    mStageOrders = mStageOrdersDefault
Case Yes
    If Not (mOrderSubmitterFactory.Capabilities And OrderSubmitterCapabilityCanStageOrders) = OrderSubmitterCapabilityCanStageOrders Then
        gWriteErrorLine StageOrdersCommand & " parameter cannot be YES with current configuration"
        Exit Sub
    End If
    mStageOrders = True
Case No
    mStageOrders = False
Case Else
    gWriteErrorLine StageOrdersCommand & " parameter must be either YES or NO or DEFAULT"
End Select

gSetValidNextCommands ListCommand, GroupCommand, ContractCommand, StageOrdersCommand, ResetCommand, CloseoutCommand
End Sub

Private Function setMonitor() As Boolean
setMonitor = True
If Not mClp.Switch(MonitorSwitch) Then
    mMonitor = False
ElseIf mClp.SwitchValue(MonitorSwitch) = "" Then
    mMonitor = True
ElseIf UCase$(mClp.SwitchValue(MonitorSwitch)) = Yes Then
    mMonitor = True
ElseIf UCase$(mClp.SwitchValue(MonitorSwitch)) = No Then
    mMonitor = False
Else
    gWriteErrorLine "The /" & MonitorSwitch & " switch has an invalid value: it must be either YES or NO (not case-sensitive, default is YES)"
    setMonitor = False
End If
End Function

Private Sub setOrderRecovery()
Const ProcName As String = "setOrderRecovery"
On Error GoTo Err

If Not mClp.Switch(ScopeNameSwitch) Then Exit Sub
mScopeName = mClp.SwitchValue(ScopeNameSwitch)
If mScopeName = "" Then Exit Sub

mGroupName = DefaultOrderGroupName
mRecoveryFileDir = ApplicationSettingsFolder
If mClp.SwitchValue(RecoveryFileDirSwitch) <> "" Then mRecoveryFileDir = mClp.SwitchValue(RecoveryFileDirSwitch)

If mOrderPersistenceDataStore Is Nothing Then Set mOrderPersistenceDataStore = CreateOrderPersistenceDataStore(mRecoveryFileDir)

mOrderManager.RecoverOrdersFromPreviousSession mScopeName, mOrderPersistenceDataStore, mOrderRecoveryAgent, mMarketDataManager, mOrderSubmitterFactory

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setStageOrders() As Boolean
setStageOrders = True
If mClp.SwitchValue(StageOrdersSwitch) = "" Then
    mStageOrdersDefault = False
ElseIf UCase$(mClp.SwitchValue(StageOrdersSwitch)) = Yes Then
    If Not (mOrderSubmitterFactory.Capabilities And OrderSubmitterCapabilityCanStageOrders) = OrderSubmitterCapabilityCanStageOrders Then
        gWriteErrorLine "The /" & StageOrdersSwitch & " switch has an invalid value: it cannot be YES with the current configuration"
        setStageOrders = False
        Exit Function
    End If
    mStageOrdersDefault = True
ElseIf UCase$(mClp.SwitchValue(StageOrdersSwitch)) = No Then
    mStageOrdersDefault = False
Else
    gWriteErrorLine "The /" & StageOrdersSwitch & " switch has an invalid value: it must be either YES or NO (not case-sensitive)"
    setStageOrders = False
End If
End Function

Private Sub setupResultsLogging(ByVal pClp As CommandLineParser)
Const ProcName As String = "setupResultsLogging"
On Error GoTo Err

Static sSetup As Boolean
If sSetup Then Exit Sub
sSetup = True

Dim lResultsPath As String

If pClp.Switch(ResultsDirSwitch) Then
    If pClp.SwitchValue(ResultsDirSwitch) <> "" Then lResultsPath = pClp.SwitchValue(ResultsDirSwitch)
ElseIf pClp.Switch(LogFileSwitch) Then
    If pClp.SwitchValue(LogFileSwitch) Then
        Dim fso As New FileSystemObject
        lResultsPath = fso.GetParentFolderName(pClp.SwitchValue(LogFileSwitch))
    End If
End If

If lResultsPath = "" Then lResultsPath = ApplicationSettingsFolder & "\Results\"
If Right$(lResultsPath, 1) <> "\" Then lResultsPath = lResultsPath & "\"

Dim lFilenameSuffix As String
lFilenameSuffix = FormatTimestamp(GetTimestamp, TimestampDateAndTime + TimestampNoMillisecs)

Dim lLogfile As FileLogListener
Set lLogfile = CreateFileLogListener(lResultsPath & "Logs\" & _
                                        ProjectName & _
                                        "-" & lFilenameSuffix & ".log", _
                                    includeTimestamp:=True, _
                                    includeLogLevel:=False)
GetLogger("log").AddLogListener lLogfile
GetLogger("position.order").AddLogListener lLogfile
GetLogger("position.simulatedorder").AddLogListener lLogfile

If mMonitor Then
    Set lLogfile = CreateFileLogListener(lResultsPath & "Orders\" & _
                                            ProjectName & _
                                            "-" & lFilenameSuffix & _
                                            ".log", _
                                        includeTimestamp:=False, _
                                        includeLogLevel:=False)
    GetLogger("position.orderdetail").AddLogListener lLogfile
    
    Set lLogfile = CreateFileLogListener(lResultsPath & "Orders\" & _
                                            ProjectName & _
                                            "-" & lFilenameSuffix & _
                                            "-Profile" & ".log", _
                                        includeTimestamp:=False, _
                                        includeLogLevel:=False)
    GetLogger("position.bracketorderprofilestring").AddLogListener lLogfile
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setupTws( _
                ByVal SwitchValue As String, _
                ByVal pSimulateOrders As Boolean) As Boolean
Const ProcName As String = "setupTws"
On Error GoTo Err

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(SwitchValue, ",")

Dim server As String
server = lClp.Arg(0)

Dim port As String
port = lClp.Arg(1)

Dim clientId As String
clientId = lClp.Arg(2)

If port = "" Then
    port = 7496
ElseIf Not IsInteger(port, 0) Then
    gWriteErrorLine "port must be an integer > 0"
    setupTws = False
End If
    
If clientId = "" Then
    clientId = DefaultClientId
ElseIf Not IsInteger(clientId, 0, 999999999) Then
    gWriteErrorLine "clientId must be an integer >= 0 and <= 999999999"
    setupTws = False
End If

Dim lTwsClient As Client
Set lTwsClient = GetClient(server, CLng(port), CLng(clientId))

Set mContractStore = lTwsClient.GetContractStore
Set mMarketDataManager = CreateRealtimeDataManager(lTwsClient.GetMarketDataFactory)
mMarketDataManager.LoadFromConfig gGetMarketDataSourcesConfig(mConfigStore)

Set mOrderRecoveryAgent = lTwsClient

If pSimulateOrders Then
    Set mOrderSubmitterFactory = New SimOrderSubmitterFactory
Else
    Set mOrderSubmitterFactory = lTwsClient
End If
    
setupTws = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub showCloseoutHelp()
gCon.WriteLineToConsole "    closeoutcommand   ::= closeout [ <groupname> | ALL ])"
End Sub

Public Sub showContractHelp()
gCon.WriteLineToConsole "    contractcommand  ::= contract [ <localsymbol>[@<exchangename>]"
gCon.WriteLineToConsole "                                  | <localsymbol>@SMART/<routinghint>"
gCon.WriteLineToConsole "                                  | /<specifier>[;/<specifier>]..."
gCon.WriteLineToConsole "                                  ] NEWLINE"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "    specifier ::= [ local[symbol]:<localsymbol>"
gCon.WriteLineToConsole "                  | symb[ol]:<symbol>"
gCon.WriteLineToConsole "                  | sec[type]:[ STK | FUT | FOP | CASH | OPT]"
gCon.WriteLineToConsole "                  | exch[ange]:<exchangename>"
gCon.WriteLineToConsole "                  | curr[ency]:<currencycode>"
gCon.WriteLineToConsole "                  | exp[iry]:[yyyymm | yyyymmdd | <offset>]"
gCon.WriteLineToConsole "                  | mult[iplier]:<multiplier>"
gCon.WriteLineToConsole "                  | str[ike]:<price>"
gCon.WriteLineToConsole "                  | right:[ CALL | PUT ]"
gCon.WriteLineToConsole "                  ]"
gCon.WriteLineToConsole ""
End Sub

Private Sub showListHelp()
gCon.WriteLineToConsole "    listcommand   ::= list [groups | positions | trades])"
End Sub

Private Sub showOrderHelp()
gCon.WriteLineToConsole "    ordercommand   ::= order <action> <quantity> <entryordertype> "
gCon.WriteLineToConsole "                           [/<orderattr>]... NEWLINE"
gCon.WriteLineToConsole "    bracketcommand ::= bracket <action> <quantity> [/<bracketattr>]... NEWLINE"
gCon.WriteLineToConsole "                       entry <entryordertype> [/<orderattr>]...  NEWLINE"
gCon.WriteLineToConsole "                       [stoploss <stoplossorderType> [/<orderattr>]...  NEWLINE]"
gCon.WriteLineToConsole "                       [target <targetorderType> [/<orderattr>]...  ] NEWLINE"
gCon.WriteLineToConsole "                       endbracket NEWLINE"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "    action     ::= [ buy | sell ]"
gCon.WriteLineToConsole "    quantity   ::= INTEGER >= 1"
gCon.WriteLineToConsole "    entryordertype  ::= [ mkt"
gCon.WriteLineToConsole "                        | lmt"
gCon.WriteLineToConsole "                        | stp"
gCon.WriteLineToConsole "                        | stplmt"
gCon.WriteLineToConsole "                        | mit"
gCon.WriteLineToConsole "                        | lit"
gCon.WriteLineToConsole "                        | moc"
gCon.WriteLineToConsole "                        | loc"
gCon.WriteLineToConsole "                        | trail"
gCon.WriteLineToConsole "                        | traillmt"
gCon.WriteLineToConsole "                        | mtl"
gCon.WriteLineToConsole "                        | moo"
gCon.WriteLineToConsole "                        | loo"
gCon.WriteLineToConsole "                        | bid"
gCon.WriteLineToConsole "                        | ask"
gCon.WriteLineToConsole "                        | last"
gCon.WriteLineToConsole "                        ]"
gCon.WriteLineToConsole "    stoplossordertype  ::= [ mkt"
gCon.WriteLineToConsole "                           | stp"
gCon.WriteLineToConsole "                           | stplmt"
gCon.WriteLineToConsole "                           | auto"
gCon.WriteLineToConsole "                           | trail"
gCon.WriteLineToConsole "                           | traillmt"
gCon.WriteLineToConsole "                           ]"
gCon.WriteLineToConsole "    targetordertype  ::= [ mkt"
gCon.WriteLineToConsole "                         | lmt"
gCon.WriteLineToConsole "                         | mit"
gCon.WriteLineToConsole "                         | lit"
gCon.WriteLineToConsole "                         | mtl"
gCon.WriteLineToConsole "                         | auto"
gCon.WriteLineToConsole "                         ]"
gCon.WriteLineToConsole "    orderattr  ::= [ price:<price>"
gCon.WriteLineToConsole "                   | trigger[price]:<price>"
gCon.WriteLineToConsole "                   | trailby:<numberofticks>"
gCon.WriteLineToConsole "                   | trailpercent:<percentage>"
gCon.WriteLineToConsole "                   | offset:<numberofticks>"
gCon.WriteLineToConsole "                   | tif:<tifvalue>"
gCon.WriteLineToConsole "                   ]"
gCon.WriteLineToConsole "    bracketattr  ::= [ cancelafter:<canceltime>"
gCon.WriteLineToConsole "                     | cancelprice:<price>"
gCon.WriteLineToConsole "                     | goodaftertime:DATETIME"
gCon.WriteLineToConsole "                     | goodtilldate:DATETIME"
gCon.WriteLineToConsole "                     | timezone:TIMEZONENAME"
gCon.WriteLineToConsole "                     ]"
gCon.WriteLineToConsole "    price  ::= DOUBLE"
gCon.WriteLineToConsole "    numberofticks  ::= INTEGER"
gCon.WriteLineToConsole "    percentage  ::= DOUBLE <= 10.0"
gCon.WriteLineToConsole "    tifvalue  ::= [ DAY"
gCon.WriteLineToConsole "                  | GTC"
gCon.WriteLineToConsole "                  | IOC"
gCon.WriteLineToConsole "                  ]"

End Sub

Private Sub showReverseHelp()
gCon.WriteLineToConsole "    reversecommand   ::= reverse [ <groupname> | ALL ])"
End Sub

Private Sub showStdInHelp()
gCon.WriteLineToConsole "StdIn Format:"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "#comment NEWLINE"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "[stageorders [yes|no|default] NEWLINE]"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "[group [<groupname>] NEWLINE]"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "<contractcommand>"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "[<ordercommand>|<bracketcommand>]..."
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "endorders NEWLINE"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "reset NEWLINE"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "<listcommand>"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "<closeoutcommand>"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "<reversecommand>"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "where"
gCon.WriteLineToConsole ""
showContractHelp
gCon.WriteLineToConsole ""
showOrderHelp
gCon.WriteLineToConsole ""
showListHelp
gCon.WriteLineToConsole ""
showCloseoutHelp
gCon.WriteLineToConsole ""
showReverseHelp
End Sub

Private Sub showStdOutHelp()
gCon.WriteLineToConsole "StdOut Format:"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "CONTRACT <contractdetails>"
gCon.WriteLineToConsole "TIME <timestamp>"
gCon.WriteLineToConsole "<orderdetails>"
gCon.WriteLineToConsole "<bracketorderdetails>"
gCon.WriteLineToConsole "ENDORDERS"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "  where"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "    contractdetails      ::= same format as for contract input"
gCon.WriteLineToConsole "    timestamp            ::= yyyy-mm-dd hh:mm:ss"
gCon.WriteLineToConsole "    orderdetails         ::= Same format as single order input"
gCon.WriteLineToConsole "    bracketorderdetails  ::= Same format as bracket order input"
End Sub

Private Sub showUsage()
gCon.WriteLineToConsole "Usage:"
gCon.WriteLineToConsole "plord27 -tws[:[<twsserver>][,[<port>][,[<clientid>]]]] [-monitor[:[yes|no]]] "
gCon.WriteLineToConsole "       [-resultsdir:<resultspath>] [-stopAt:<hh:mm>] [-log:<logfilepath>]"
gCon.WriteLineToConsole "       [-loglevel:[ I | N | D | M | H }]"
gCon.WriteLineToConsole "       [-stageorders[:[yes|no]]]"
gCon.WriteLineToConsole "       [-scopename:<scope>] [-recoveryfiledir:<recoverypath>]"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "  where"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "    twsserver  ::= STRING name or IP address of computer where TWS/Gateway"
gCon.WriteLineToConsole "                          is running"
gCon.WriteLineToConsole "    port       ::= INTEGER port to be used for API connection"
gCon.WriteLineToConsole "    clientid   ::= INTEGER client id >=0 to be used for API connection (default"
gCon.WriteLineToConsole "                           value is " & DefaultClientId & ")"
gCon.WriteLineToConsole "    resultspath ::= path to the folder in which results files  are to be created"
gCon.WriteLineToConsole "                    (defaults to the logfile path)"
gCon.WriteLineToConsole "    logfilepath ::= path to the folder where the program logfile is to be"
gCon.WriteLineToConsole "                    created"
gCon.WriteLineToConsole "    scope       ::= order recovery scope name"
gCon.WriteLineToConsole "    recoverypath ::= path to the folder in which order recovery files  are to"
gCon.WriteLineToConsole "                     be created(defaults to the logfile path)"
gCon.WriteLineToConsole ""
showStdInHelp
gCon.WriteLineToConsole ""
showStdOutHelp
gCon.WriteLineToConsole ""
End Sub



