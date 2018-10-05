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

Public Const ProjectName                            As String = "plord"
Private Const ModuleName                            As String = "MainMod"

Private Const TwsSwitch                             As String = "TWS"

Public Const CancelAfterSwitch                      As String = "CANCELAFTER"
Public Const CancelPriceSwitch                      As String = "CANCELPRICE"
Public Const LogFileSwitch                          As String = "LOG"
Public Const MonitorSwitch                          As String = "MONITOR"
Public Const OffsetSwitch                           As String = "OFFSET"
Public Const PriceSwitch                            As String = "PRICE"
Public Const ResultsDirSwitch                       As String = "RESULTSDIR"
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

Public Const BracketCommand                         As String = "BRACKET"
Public Const ContractCommand                        As String = "CONTRACT"
Public Const EndBracketCommand                      As String = "ENDBRACKET"
Public Const EndOrdersCommand                       As String = "ENDORDERS"
Public Const EntryCommand                           As String = "ENTRY"
Public Const ExitCommand                            As String = "EXIT"
Public Const HelpCommand                            As String = "HELP"
Public Const Help1Command                           As String = "?"
Public Const OrderCommand                           As String = "ORDER"
Public Const QuitCommand                            As String = "QUIT"
Public Const StopLossCommand                        As String = "STOPLOSS"
Public Const TargetCommand                          As String = "TARGET"
Public Const StageOrdersCommand                     As String = "STAGEORDERS"

Public Const ValueSeparator                         As String = ":"
Public Const SwitchPrefix                           As String = "/"

Private Const Yes                                   As String = "YES"
Private Const No                                    As String = "NO"

'@================================================================================
' Member variables
'@================================================================================

Public gCon                                         As Console

Public gLogToConsole                                As Boolean

' This flag is set when a Contract command is being processed: this is because
' the contract fetch and starting the market data happen asynchronously, and we
' don't want any further user input until we know whether it has succeeded.
Public gInputPaused                                 As Boolean

Public gErrorCount                                  As Long

Public gNumberOfOrdersPlaced                        As Long

Private mFatalErrorHandler                          As FatalErrorHandler

Private mContractStore                              As IContractStore
Private mMarketDataManager                          As IMarketDataManager
Private mOrderSubmitterFactory                      As IOrderSubmitterFactory

Private mStageOrders                                As Boolean

Private mProcessor                                  As Processor

Private mContractSpec                               As IContractSpecifier

Private mProcessors                                 As New EnumerableCollection

Private mCommandNumber                              As Long

Private mValidNextCommands()                        As String

Private mClp                                        As CommandLineParser

Private mMonitor                                    As Boolean

Private mLineNumber                                 As Long

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
gGenerateSwitch = SwitchPrefix & pName & IIf(pValue = "", "", ValueSeparator & pValue & " ")
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
gErrorCount = gErrorCount + 1
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

Set gCon = GetConsole

Set mClp = CreateCommandLineParser(command)

If Trim$(command) = "/?" Or Trim$(command) = "-?" Then
    showUsage
ElseIf mClp.Switch(TwsSwitch) Then
    If setupTws(mClp.switchValue(TwsSwitch)) Then process
ElseIf setupTws("") Then
    process
Else
    showUsage
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

Private Function isCommandValid(ByVal pCommand As String) As Boolean
Dim i As Long
For i = 0 To UBound(mValidNextCommands)
    If pCommand = mValidNextCommands(i) Then
        isCommandValid = True
        Exit Function
    End If
Next
End Function

Private Sub process()
Const ProcName As String = "process"
On Error GoTo Err

If mClp.Switch(MonitorSwitch) Then
    mMonitor = True
ElseIf mClp.switchValue(MonitorSwitch) = "" Then
    mMonitor = True
ElseIf UCase$(mClp.switchValue(MonitorSwitch)) = "YES" Then
    mMonitor = True
ElseIf UCase$(mClp.switchValue(MonitorSwitch)) = "NO" Then
    mMonitor = False
Else
    gWriteErrorLine "The /" & MonitorSwitch & " switch has an invalid value: it must be either YES or NO (not case-sensitive)"
End If

gSetValidNextCommands StageOrdersCommand

Dim inString As String
inString = Trim$(gCon.ReadLine(":"))

Do While inString <> gCon.EofString
    gLogger.Log "StdIn: " & inString, ProcName, ModuleName
    mLineNumber = mLineNumber + 1
    If inString = "" Then
        ' ignore blank lines
    ElseIf Left$(inString, 1) = "#" Then
        ' ignore comments
    Else
        Dim lExit As Boolean
        processCommand inString, lExit
        If lExit Then Exit Do
    End If
    
    Do While gInputPaused
        Wait 20
    Loop
    inString = Trim$(gCon.ReadLine(":"))
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processCommand(ByVal pInstring As String, ByRef pExit As Boolean)
Const ProcName As String = "processCommand"
On Error GoTo Err

Dim command As String
command = UCase$(Split(pInstring, " ")(0))

Dim params As String
params = Trim$(Right$(pInstring, Len(pInstring) - Len(command)))

If command = ExitCommand Then
    gWriteLineToConsole "Exiting"
    pExit = True
    Exit Sub
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
        Set mProcessor = processContractCommand(params)
    Case StageOrdersCommand
        processStageOrdersCommand params
    Case BracketCommand
        mProcessor.ProcessBracketCommand params
    Case EntryCommand
        mProcessor.ProcessEntryCommand params
    Case StopLossCommand
        mProcessor.ProcessStopLossCommand params
    Case TargetCommand
        mProcessor.ProcessTargetCommand params
    Case EndBracketCommand
        mProcessor.ProcessEndBracketCommand
    Case EndOrdersCommand
        processEndOrdersCommand
    Case Else
        gWriteErrorLine "Invalid command '" & command & "'"
    End Select
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processContractCommand(ByVal pParams As String) As Processor
Const ProcName As String = "processContractCommand"
On Error GoTo Err

Dim lProcessor As New Processor
lProcessor.StageOrders = mStageOrders
mProcessors.Add lProcessor
        
lProcessor.Initialise mContractStore, mMarketDataManager, mOrderSubmitterFactory

gInputPaused = True
If Not lProcessor.processContractCommand(pParams) Then gInputPaused = False

Set processContractCommand = lProcessor

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processEndOrdersCommand()
Const ProcName As String = "processEndOrdersCommand"
On Error GoTo Err

If gErrorCount <> 0 Then
    gWriteErrorLine gErrorCount & " errors have been found - no orders will be placed"
    gErrorCount = 0
    If Not mProcessor Is Nothing Then
        gSetValidNextCommands BracketCommand
    Else
        gSetValidNextCommands ContractCommand
    End If
    Exit Sub
End If

' if there have been no errors, we need to tell each processor to submit
' its orders. To avoid exceeding the API's input message limits, we do this
' asynchronously with a task, which means we need to inhibit user input
' until it has completed.
gInputPaused = True

Dim t As New PlaceOrdersTask
t.Initialise mProcessors, mStageOrders
StartTask t, PriorityNormal

setupResultsLogging mClp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processStageOrdersCommand( _
                ByVal pParams As String)
Select Case UCase$(pParams)
Case Yes
    mStageOrders = True
Case No
    mStageOrders = False
Case Else
    gWriteErrorLine StageOrdersCommand & " parameter must be either YES or NO"
End Select

gSetValidNextCommands ContractCommand
End Sub

Private Sub setupResultsLogging(ByVal pClp As CommandLineParser)
Const ProcName As String = "setupResultsLogging"
On Error GoTo Err

Static sSetup As Boolean
If sSetup Then Exit Sub
sSetup = True

Dim lResultsPath As String

If pClp.Switch(ResultsDirSwitch) Then
    If pClp.switchValue(ResultsDirSwitch) <> "" Then lResultsPath = pClp.switchValue(ResultsDirSwitch)
ElseIf pClp.Switch(LogFileSwitch) Then
    If pClp.switchValue(LogFileSwitch) Then
        Dim fso As New FileSystemObject
        lResultsPath = fso.GetParentFolderName(pClp.switchValue(LogFileSwitch))
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
                ByVal switchValue As String) As Boolean
Const ProcName As String = "setupTws"
On Error GoTo Err

Dim mClp As CommandLineParser
Set mClp = CreateCommandLineParser(switchValue, ",")

Dim server As String
server = mClp.Arg(0)

Dim port As String
port = mClp.Arg(1)

Dim clientId As String
clientId = mClp.Arg(2)

If port = "" Then
    port = 7496
ElseIf Not IsInteger(port, 0) Then
    gWriteErrorLine "Error: port must be an integer > 0"
    setupTws = False
End If
    
If clientId = "" Then
    clientId = &H71A3DD2E
ElseIf Not IsInteger(clientId, 0) Then
    gWriteErrorLine "Error: clientId must be an integer >= 0"
    setupTws = False
End If

Dim lTwsClient As Client
Set lTwsClient = GetClient(server, CLng(port), CLng(clientId))

Set mContractStore = lTwsClient.GetContractStore
Set mMarketDataManager = CreateRealtimeDataManager(lTwsClient.GetMarketDataFactory)
Set mOrderSubmitterFactory = lTwsClient
    
setupTws = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub showContractHelp()
gCon.WriteLineToConsole "    contractcommand  ::= contract  [(/<specifier>[;/<specifier>]...) NEWLINE]"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "    specifier ::= [ local[symbol]:<localsymbol>"
gCon.WriteLineToConsole "                  | symb[ol]:<symbol>"
gCon.WriteLineToConsole "                  | sec[type]:[ STK | FUT | FOP | CASH | OPT]"
gCon.WriteLineToConsole "                  | exch[ange]:<exchangename>"
gCon.WriteLineToConsole "                  | curr[ency]:<currencycode>"
gCon.WriteLineToConsole "                  | exp[iry]:[yyyymm | yyyymmdd]"
gCon.WriteLineToConsole "                  | mult[iplier]:<multiplier>"
gCon.WriteLineToConsole "                  | str[ike]:<price>"
gCon.WriteLineToConsole "                  | right:[ CALL | PUT ]"
gCon.WriteLineToConsole "                  ]"
gCon.WriteLineToConsole ""
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
gCon.WriteLineToConsole "                     ]"
gCon.WriteLineToConsole "    price  ::= DOUBLE"
gCon.WriteLineToConsole "    numberofticks  ::= INTEGER"
gCon.WriteLineToConsole "    percentage  ::= DOUBLE <= 10.0"
gCon.WriteLineToConsole "    tifvalue  ::= [ DAY"
gCon.WriteLineToConsole "                  | GTC"
gCon.WriteLineToConsole "                  | IOC"
gCon.WriteLineToConsole "                  ]"

End Sub

Private Sub showStdInHelp()
gCon.WriteLineToConsole "StdIn Format:"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "#comment"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "stageorders [yes|no]"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "<contractcommand>"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "[<ordercommand>|<bracketcommand>]..."
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "endorders NEWLINE"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "where"
gCon.WriteLineToConsole ""
showContractHelp
gCon.WriteLineToConsole ""
showOrderHelp
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
gCon.WriteLineToConsole "plord27 -tws[:[<twsserver>][,[<port>][,[<clientid>]]]] [-monitor:[yes|no]] "
gCon.WriteLineToConsole "       [-resultsdir:<resultspath>] [-stopAt:<hh:mm>] [-log:<logfilepath>"
gCon.WriteLineToConsole "       [-loglevel:[ I | N | D | M | H }"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "  where"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "    twsserver  ::= STRING name or IP address of computer where TWS/Gateway"
gCon.WriteLineToConsole "                          is running"
gCon.WriteLineToConsole "    port       ::= INTEGER port to be used for API connection"
gCon.WriteLineToConsole "    clientid   ::= INTEGER client id >=0 to be used for API connection (default"
gCon.WriteLineToConsole "                           value is 1906564398"
gCon.WriteLineToConsole "    resultspath ::= path to the folder in which results files  are to be created"
gCon.WriteLineToConsole "                    (defaults to the logfile path)"
gCon.WriteLineToConsole "    logfilepath ::= path to the folder where the program logfile is to be"
gCon.WriteLineToConsole "                    created"
gCon.WriteLineToConsole ""
showStdInHelp
gCon.WriteLineToConsole ""
showStdOutHelp
gCon.WriteLineToConsole ""
End Sub



