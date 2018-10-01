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
Public Const OffsetSwitch                           As String = "OFFSET"
Public Const PriceSwitch                            As String = "PRICE"
Public Const TIFSwitch                              As String = "TIF"
Public Const TrailBySwitch                          As String = "TRAILBY"
Public Const TrailPercentSwitch                     As String = "TRAILPERCENT"
Public Const TriggerPriceSwitch                     As String = "TRIGGER"
Public Const TriggerPriceSwitch1                    As String = "TRIGGERPRICE"

Public Const BracketCommand                         As String = "BRACKET"
Public Const ContractCommand                        As String = "CONTRACT"
Public Const EndCommand                             As String = "END"
Public Const EndBracketCommand                      As String = "ENDBRACKET"
Public Const EntryCommand                           As String = "ENTRY"
Public Const ExitCommand                            As String = "EXIT"
Public Const HelpCommand                            As String = "HELP"
Public Const Help1Command                           As String = "?"
Public Const OrderCommand                           As String = "ORDER"
Public Const QuitCommand                            As String = "QUIT"
Public Const StopLossCommand                        As String = "STOPLOSS"
Public Const TargetCommand                          As String = "TARGET"
Public Const TransmitCommand                        As String = "TRANSMIT"

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



Private mFatalErrorHandler                          As FatalErrorHandler

Private mContractStore                              As IContractStore
Private mMarketDataManager                          As IMarketDataManager
Private mOrderSubmitterFactory                      As IOrderSubmitterFactory

Private mTransmit                                   As Boolean

Private mContractSpec                               As IContractSpecifier

Private mProcessors                                 As New EnumerableCollection

Private mCommandNumber                              As Long

Private mValidNextCommands()                        As String


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
gCon.WriteErrorLine pMessage
gErrorCount = gErrorCount + 1
End Sub

Public Sub gWriteLineToConsole(ByVal pMessage As String)
gCon.WriteLineToConsole pMessage
End Sub

Public Sub Main()
Const ProcName As String = "Main"
On Error GoTo Err

InitialiseTWUtilities

Set mFatalErrorHandler = New FatalErrorHandler
ApplicationGroupName = "TradeWright"
ApplicationName = "plord"
SetupDefaultLogging command

Set gCon = GetConsole

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(command)

If clp.Switch("?") Or _
    clp.NumberOfSwitches = 0 _
Then
    showUsage
ElseIf clp.Switch(TwsSwitch) Then
    If setupTws(clp.switchValue(TwsSwitch)) Then process
Else
    showUsage
End If

TerminateTWUtilities

Exit Sub

Err:
If Not gCon Is Nothing Then
    gCon.WriteErrorLine Err.Description
    gCon.WriteErrorLine "At:"
    gCon.WriteErrorLine Err.Source
End If

TerminateTWUtilities
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

gSetValidNextCommands TransmitCommand

Dim inString As String
inString = Trim$(gCon.ReadLine(":"))

Do While inString <> gCon.EofString
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

If command = EndCommand Then
    processEndCommand
ElseIf Not isCommandValid(command) Then
    gWriteErrorLine "Valid commands at this point are: " & Join(mValidNextCommands, ",")
Else
    Static sProcessor As Processor
    Select Case command
    Case ContractCommand
        Set sProcessor = processContractCommand(params)
    Case HelpCommand, Help1Command
        showStdInHelp
    Case TransmitCommand
        processTransmitCommand params
    Case BracketCommand
        sProcessor.ProcessBracketCommand params
    Case EntryCommand
        sProcessor.ProcessEntryCommand params
    Case StopLossCommand
        sProcessor.ProcessStopLossCommand params
    Case TargetCommand
        sProcessor.ProcessTargetCommand params
    Case EndBracketCommand
        sProcessor.ProcessEndBracketCommand
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
lProcessor.Transmit = mTransmit
mProcessors.Add lProcessor
        
lProcessor.Initialise mContractStore, mMarketDataManager, mOrderSubmitterFactory

gInputPaused = True
If Not lProcessor.processContractCommand(pParams) Then gInputPaused = False

Set processContractCommand = lProcessor

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processEndCommand()
Const ProcName As String = "processEndCommand"
On Error GoTo Err

If gErrorCount <> 0 Then
    gWriteErrorLine gErrorCount & " errors have been found - no orders will be placed"
    Exit Sub
End If

' if there have been no errors, we need to tell each processor to submit
' its orders. To avoid exceeding the API's input message limits, we do this
' asynchronously with a task, which means we need to inhibit user input
' until it has completed.
gInputPaused = True

Dim t As New PlaceOrdersTask
t.Initialise mProcessors
StartTask t, PriorityNormal

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTransmitCommand( _
                ByVal pParams As String)
Select Case UCase$(pParams)
Case Yes
    mTransmit = True
    gSetValidNextCommands ContractCommand
Case No
    mTransmit = False
    gSetValidNextCommands ContractCommand
Case Else
    gWriteErrorLine "Transmit must be either YES or NO"
End Select

End Sub

Private Function setupTws( _
                ByVal switchValue As String) As Boolean
Const ProcName As String = "setupTws"
On Error GoTo Err

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(switchValue, ",")

Dim server As String
server = clp.Arg(0)

Dim port As String
port = clp.Arg(1)

Dim clientId As String
clientId = clp.Arg(2)

If port = "" Then
    port = 7496
ElseIf Not IsInteger(port, 0) Then
    gWriteErrorLine "Error: port must be an integer > 0"
    setupTws = False
End If
    
If clientId = "" Then
    clientId = &H7A92DC3F
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
gCon.WriteLineToConsole "contract  [(/specifier[;/specifier]...)]"

gCon.WriteLineToConsole "   specifier := [ local[symbol]:<localsymbol>"
gCon.WriteLineToConsole "                | symb[ol]:<symbol>"
gCon.WriteLineToConsole "                | sec[type]:[ STK | FUT | FOP | CASH ]"
gCon.WriteLineToConsole "                | exch[ange]:<exchangename>"
gCon.WriteLineToConsole "                | curr[ency]:<currencycode>"
gCon.WriteLineToConsole "                | exp[iry]:[yyyymm | yyyymmdd]"
gCon.WriteLineToConsole "                | mult[iplier]:<multiplier>"
gCon.WriteLineToConsole "                | str[ike]:<price>"
gCon.WriteLineToConsole "                | right:[ CALL | PUT ]"
gCon.WriteLineToConsole "                ]"

gCon.WriteLineToConsole ""
End Sub

Private Sub showOrderHelp()
gCon.WriteLineToConsole "[ order <action> <quantity> <entryordertype> [/<orderattr>]..."
gCon.WriteLineToConsole "| bracket <action> <quantity> [/<bracketattr>]... "
gCon.WriteLineToConsole "  entry <entryordertype> [/<orderattr>]...  "
gCon.WriteLineToConsole "  [stoploss <stoplossorderType> [/<orderattr>]...  ]"
gCon.WriteLineToConsole "  [target <targetorderType> [/<orderattr>]...  ]"
gCon.WriteLineToConsole "  endbracket"
gCon.WriteLineToConsole "]"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "where"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "    action     := [ buy | sell ]"
gCon.WriteLineToConsole "    quantity   := INTEGER >= 1"
gCon.WriteLineToConsole "    entryordertype  := [ mkt"
gCon.WriteLineToConsole "                       | lmt"
gCon.WriteLineToConsole "                       | stp"
gCon.WriteLineToConsole "                       | stplmt"
gCon.WriteLineToConsole "                       | mit"
gCon.WriteLineToConsole "                       | lit"
gCon.WriteLineToConsole "                       | moc"
gCon.WriteLineToConsole "                       | loc"
gCon.WriteLineToConsole "                       | trail"
gCon.WriteLineToConsole "                       | traillmt"
gCon.WriteLineToConsole "                       | mtl"
gCon.WriteLineToConsole "                       | moo"
gCon.WriteLineToConsole "                       | loo"
gCon.WriteLineToConsole "                       ]"
gCon.WriteLineToConsole "    stoplossordertype  := [ mkt"
gCon.WriteLineToConsole "                          | stp"
gCon.WriteLineToConsole "                          | stplmt"
gCon.WriteLineToConsole "                          | trail"
gCon.WriteLineToConsole "                          | traillmt"
gCon.WriteLineToConsole "                          ]"
gCon.WriteLineToConsole "    targetordertype  := [ mkt"
gCon.WriteLineToConsole "                        | lmt"
gCon.WriteLineToConsole "                        | mit"
gCon.WriteLineToConsole "                        | lit"
gCon.WriteLineToConsole "                        | mtl"
gCon.WriteLineToConsole "                        ]"
gCon.WriteLineToConsole "    orderattr  := [ price:<price>"
gCon.WriteLineToConsole "                  | trigger[price]:<price>"
gCon.WriteLineToConsole "                  | trailby:<numberofticks>"
gCon.WriteLineToConsole "                  | trailpercent:<percentage>"
gCon.WriteLineToConsole "                  | offset:<numberofticks>"
gCon.WriteLineToConsole "                  | tif:<tifvalue>"
gCon.WriteLineToConsole "                  ]"
gCon.WriteLineToConsole "    bracketattr  := [ cancelafter:<canceltime>"
gCon.WriteLineToConsole "                    | cancelprice:<price>"
gCon.WriteLineToConsole "                    ]"
gCon.WriteLineToConsole "    price  := DOUBLE"
gCon.WriteLineToConsole "    numberofticks  := INTEGER"
gCon.WriteLineToConsole "    percentage  := DOUBLE <= 10.0"
gCon.WriteLineToConsole "    tifvalue  := [ DAY"
gCon.WriteLineToConsole "                 | GTC"
gCon.WriteLineToConsole "                 | IOC"

End Sub

Private Sub showStdInHelp()
gCon.WriteLineToConsole "StdIn Format:"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "#comment"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "transmit [yes|no]"
gCon.WriteLineToConsole ""

showContractHelp

showOrderHelp

gCon.WriteLineToConsole "end"
End Sub

Private Sub showUsage()
gCon.WriteLineToConsole "Usage:"
gCon.WriteLineToConsole "plord27 -tws:[twsserver][,[port][,[clientid]]] [-monitor] [-stopAt:<hh:mm>]"
gCon.WriteLineToConsole ""
showStdInHelp
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "StdOut Format:"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "timestamp,open,high,low,close,volume,tickvolume"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "  where"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "    timestamp ::= yyyy-mm-dd hh:mm:ss.nnn"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole ""
End Sub



