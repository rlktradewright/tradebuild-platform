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

Private Const BatchOrdersSwitch                     As String = "BATCHORDERS"
Private Const LogFileSwitch                         As String = "LOG"
Private Const LogApiMessages                        As String = "LOGAPIMESSAGES"
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

Public Const BatchOrdersCommand                     As String = "BATCHORDERS"
Public Const BracketCommand                         As String = "BRACKET"
Public Const BuyCommand                             As String = "BUY"
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
Public Const QuitCommand                            As String = "QUIT"
Public Const ResetCommand                           As String = "RESET"
Public Const SellCommand                            As String = "SELL"
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
Private mGroupContractProcessors                    As SortedDictionary

Private mCommandNumber                              As Long

Private mValidNextCommands()                        As String

Private mClp                                        As CommandLineParser

Private mMonitor                                    As Boolean

Private mLineNumber                                 As Long

Private mRecoveryFileDir                            As String

Private mOrderPersistenceDataStore                  As IOrderPersistenceDataStore
Private mOrderRecoveryAgent                         As IOrderRecoveryAgent

Private mConfigStore                                As ConfigurationStore

Private mBatchOrdersDefault                         As Boolean
Private mBatchOrders                                As Boolean

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

Public Property Get gRegExp() As RegExp
Static lRegexp As RegExp
If lRegexp Is Nothing Then Set lRegexp = New RegExp
Set gRegExp = lRegexp
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

Public Function gNotifyContractFutureAvailable( _
                ByVal pContractFuture As IFuture) As ContractProcessor
Const ProcName As String = "gNotifyContractAvailable"
On Error GoTo Err

Const PositionManagerNameSeparator As String = "&&"

If pContractFuture.IsCancelled Then
    gWriteErrorLine "Contract fetch was cancelled"
    Exit Function
End If

If pContractFuture.IsFaulted Then
    gWriteErrorLine pContractFuture.ErrorMessage
    Exit Function
End If

Assert TypeOf pContractFuture.Value Is IContract, "Unexpected future value"

Dim lContract As IContract
Set lContract = pContractFuture.Value

If IsContractExpired(lContract) Then
    gWriteErrorLine "Contract has expired"
    Exit Function
End If
    
Dim lContractProcessorName As String
lContractProcessorName = mGroupName & _
                        PositionManagerNameSeparator & _
                        lContract.Specifier.Key

If Not mContractProcessors.TryItem(lContractProcessorName, mContractProcessor) Then
    Set mContractProcessor = New ContractProcessor
    mContractProcessor.Initialise lContractProcessorName, pContractFuture, mMarketDataManager, mOrderManager, mScopeName, mGroupName, mOrderSubmitterFactory
    mContractProcessors.Add mContractProcessor, lContractProcessorName
End If

If mGroupContractProcessors.Contains(mGroupName) Then mGroupContractProcessors.Remove mGroupName
mGroupContractProcessors.Add mContractProcessor, mGroupName

Set gNotifyContractFutureAvailable = mContractProcessor

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Public Sub gProcessOrders()
Const ProcName As String = "gProcessOrders"
On Error GoTo Err

' we need to tell each ContractProcessor to submit
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

If Trim$(command) = "/?" Or Trim$(command) = "-?" Then
    showUsage
    Exit Sub
End If

InitialiseTWUtilities

Set mFatalErrorHandler = New FatalErrorHandler
ApplicationGroupName = "TradeWright"
ApplicationName = "plord"
SetupDefaultLogging command, True, True

Set mConfigStore = gGetConfigStore

Set mCloseoutProcessor = New CloseoutProcessor
mCloseoutProcessor.Initialise mOrderManager

Set gCon = GetConsole

logProgramId

Set mGroupContractProcessors = CreateSortedDictionary(KeyTypeString)

Set mClp = CreateCommandLineParser(command)
If setupTws(mClp.SwitchValue(TwsSwitch), mClp.Switch(SimulateOrdersSwitch), mClp.Switch(LogApiMessages)) Then
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
    getPrompt = mGroupName & DefaultPrompt
ElseIf mContractProcessor.Contract Is Nothing Then
    getPrompt = mGroupName & DefaultPrompt
Else
    getPrompt = mGroupName & "!" & _
                mContractProcessor.Contract.Specifier.LocalSymbol & _
                "@" & _
                mContractProcessor.Contract.Specifier.Exchange & _
                DefaultPrompt
End If
End Function

Private Function getNumberOfUnprocessedOrders() As Long
Const ProcName As String = "getNumberOfUnprocessedOrders"
On Error GoTo Err

Dim lProcessor As ContractProcessor
For Each lProcessor In mContractProcessors
    getNumberOfUnprocessedOrders = getNumberOfUnprocessedOrders + lProcessor.BracketOrders.Count
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Private Function isGroupValid(ByVal pGroup As String) As Boolean
Const ProcName As String = "isGroupValid"
On Error GoTo Err

gRegExp.Global = True
gRegExp.Pattern = "^[a-zA-Z0-9][\w-]*$"
isGroupValid = gRegExp.Test(pGroup)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
    If lPM.PositionSize <> 0 Or _
        lPM.PendingBuyPositionSize <> 0 Or _
        lPM.PendingSellPositionSize <> 0 _
    Then
        Dim lContract As IContract
        Set lContract = lPM.ContractFuture.Value
        gWriteLineToConsole padStringRight(lPM.GroupName, 15) & " " & _
                            padStringRight(lContract.Specifier.LocalSymbol & "@" & lContract.Specifier.Exchange, 25) & _
                            " Size=" & padStringleft(lPM.PositionSize & _
                                                    "(" & lPM.PendingBuyPositionSize & _
                                                    "/" & _
                                                    lPM.PendingSellPositionSize & ")", 10) & _
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

Private Sub logProgramId()
Const ProcName As String = "logProgramId"
On Error GoTo Err

Dim s As String
s = App.ProductName & " V" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
    App.LegalCopyright
gWriteLineToConsole s
s = s & vbCrLf & "Arguments: " & command
gLogger.Log s, ProcName, ModuleName

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

Private Function parseContractSpec( _
                ByVal pClp As CommandLineParser) As IContractSpecifier
Const ProcName As String = "parseContractSpec"
On Error GoTo Err

If pClp.Arg(0) = "?" Or _
    pClp.Switch("?") Or _
    (pClp.Arg(0) = "" And pClp.NumberOfSwitches = 0) _
Then
    showContractHelp
    Exit Function
End If

Dim validParams As Boolean
validParams = True

Dim lSectypeStr As String: lSectypeStr = pClp.SwitchValue(SecTypeSwitch)
If lSectypeStr = "" Then lSectypeStr = pClp.SwitchValue(SecTypeSwitch1)

Dim lExchange As String: lExchange = pClp.SwitchValue(ExchangeSwitch)
If lExchange = "" Then lExchange = pClp.SwitchValue(ExchangeSwitch1)

Dim lLocalSymbol As String: lLocalSymbol = pClp.SwitchValue(LocalSymbolSwitch)
If lLocalSymbol = "" Then lLocalSymbol = pClp.SwitchValue(LocalSymbolSwitch1)

Dim lSymbol As String: lSymbol = pClp.SwitchValue(SymbolSwitch)
If lSymbol = "" Then lSymbol = pClp.SwitchValue(SymbolSwitch1)

Dim lCurrency As String: lCurrency = pClp.SwitchValue(CurrencySwitch)
If lCurrency = "" Then lCurrency = pClp.SwitchValue(CurrencySwitch1)

Dim lExpiry As String: lExpiry = pClp.SwitchValue(ExpirySwitch)
If lExpiry = "" Then lExpiry = pClp.SwitchValue(ExpirySwitch1)

Dim lMultiplier As String: lMultiplier = pClp.SwitchValue(MultiplierSwitch)
If lMultiplier = "" Then lMultiplier = pClp.SwitchValue(MultiplierSwitch1)
If lMultiplier = "" Then lMultiplier = "1.0"

Dim lStrike As String: lStrike = pClp.SwitchValue(StrikeSwitch)
If lStrike = "" Then lStrike = pClp.SwitchValue(StrikeSwitch1)
If lStrike = "" Then lStrike = "0.0"

Dim lRight As String: lRight = pClp.SwitchValue(RightSwitch)

Dim lSectype As SecurityTypes
lSectype = SecTypeFromString(lSectypeStr)
If lSectypeStr <> "" And lSectype = SecTypeNone Then
    gWriteErrorLine "Invalid Sectype '" & lSectypeStr & "'"
    validParams = False
End If

If lExpiry <> "" Then
    If IsInteger(lExpiry, 0, MaxContractExpiryOffset) Then
    ElseIf IsDate(lExpiry) Then
        lExpiry = Format(CDate(lExpiry), "yyyymmdd")
    ElseIf Len(lExpiry) = 6 Then
        If Not IsDate(Left$(lExpiry, 4) & "/" & Right$(lExpiry, 2) & "/01") Then
            gWriteErrorLine "Invalid Expiry '" & lExpiry & "'"
            validParams = False
        End If
    ElseIf Len(lExpiry) = 8 Then
        If Not IsDate(Left$(lExpiry, 4) & "/" & Mid$(lExpiry, 5, 2) & "/" & Right$(lExpiry, 2)) Then
            gWriteErrorLine "Invalid Expiry '" & lExpiry & "'"
            validParams = False
        End If
    Else
        gWriteErrorLine "Invalid Expiry '" & lExpiry & "'"
        validParams = False
    End If
End If
            
Dim Multiplier As Double
If lMultiplier = "" Then
    Multiplier = 1#
ElseIf IsNumeric(lMultiplier) Then
    Multiplier = CDbl(lMultiplier)
Else
    gWriteErrorLine "Invalid multiplier '" & lMultiplier & "'"
    validParams = False
End If
            
Dim Strike As Double
If lStrike <> "" Then
    If IsNumeric(lStrike) Then
        Strike = CDbl(lStrike)
    Else
        gWriteErrorLine "Invalid strike '" & lStrike & "'"
        validParams = False
    End If
End If

Dim optRight As OptionRights
optRight = OptionRightFromString(lRight)
If lRight <> "" And optRight = OptNone Then
    gWriteErrorLine "Invalid right '" & lRight & "'"
    validParams = False
End If

        
If validParams Then
    Set parseContractSpec = CreateContractSpecifier(lLocalSymbol, _
                                            lSymbol, _
                                            lExchange, _
                                            lSectype, _
                                            lCurrency, _
                                            lExpiry, _
                                            Multiplier, _
                                            Strike, _
                                            optRight)
End If

Exit Function

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Function

Private Sub process()
Const ProcName As String = "process"
On Error GoTo Err

If Not setMonitor Then Exit Sub
If Not setStageOrders Then Exit Sub
If Not setBatchOrders Then Exit Sub

setOrderRecovery

gSetValidNextCommands ListCommand, StageOrdersCommand, BatchOrdersCommand, GroupCommand, ContractCommand, BuyCommand, SellCommand, CloseoutCommand, ResetCommand

Dim inString As String
inString = Trim$(gCon.ReadLine(getPrompt))

Do While inString <> gCon.EofString
    If inString = "" Then
        ' ignore blank lines - and don't write to the log because
        ' the FileAutoReader program sends blank lines very frequently
    ElseIf Left$(inString, 1) = "#" Then
        ' ignore comments except log and write to console
        gLogger.Log "StdIn: " & inString, ProcName, ModuleName
        gWriteLineToConsole inString
    Else
        mLineNumber = mLineNumber + 1
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

Dim Params As String
Params = Trim$(Right$(pInstring, Len(pInstring) - Len(command)))

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
        processContractCommand Params
    Case BatchOrdersCommand
        processBatchOrdersCommand Params
    Case StageOrdersCommand
        processStageOrdersCommand Params
    Case GroupCommand
        processGroupCommand Params
    Case BuyCommand
        ProcessBuyCommand Params
    Case SellCommand
        ProcessSellCommand Params
    Case BracketCommand
        mContractProcessor.ProcessBracketCommand Params
    Case EntryCommand
        mContractProcessor.ProcessEntryCommand Params
    Case StopLossCommand
        mContractProcessor.ProcessStopLossCommand Params
    Case TargetCommand
        mContractProcessor.ProcessTargetCommand Params
    Case EndBracketCommand
        mContractProcessor.ProcessEndBracketCommand
        If mErrorCount = 0 And Not mBatchOrders Then gProcessOrders
    Case EndOrdersCommand
        processEndOrdersCommand
    Case ResetCommand
        processResetCommand
    Case ListCommand
        processListCommand Params
    Case CloseoutCommand
        processCloseoutCommand Params
    Case Else
        gWriteErrorLine "Invalid command '" & command & "'"
    End Select
End If

processCommand = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processBatchOrdersCommand( _
                ByVal pParams As String)
Select Case UCase$(pParams)
Case Default
    mBatchOrders = mBatchOrdersDefault
Case Yes
    mBatchOrders = True
Case No
    mBatchOrders = False
Case Else
    gWriteErrorLine BatchOrdersCommand & " parameter must be either YES or NO or DEFAULT"
End Select

gSetValidNextCommands ListCommand, GroupCommand, ContractCommand, BuyCommand, SellCommand, StageOrdersCommand, BatchOrdersCommand, ResetCommand, CloseoutCommand
End Sub

Private Sub ProcessBuyCommand( _
                ByVal pParams As String)
Const ProcName As String = "processBuyCommand"
On Error GoTo Err

processBuyOrSellCommand OrderActionBuy, pParams

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processBuyOrSellCommand( _
                ByVal pAction As OrderActions, _
                ByVal pParams As String)
Const ProcName As String = "processBuyOrSellCommand"
On Error GoTo Err

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pParams)
Dim lArg0 As String: lArg0 = lClp.Arg(0)

If Not IsInteger(lArg0, 1) Then
    ' the first arg is a contract spec
    Dim lContractSpec As IContractSpecifier
    Set lContractSpec = CreateContractSpecifierFromString(lArg0)

    If lContractSpec Is Nothing Then Exit Sub
    
    Dim lData As New BuySellCommandData
    lData.Action = pAction
    lData.Params = Right$(pParams, Len(pParams) - Len(lArg0))
    
    Dim lResolver As New ContractResolver
    lResolver.Initialise lContractSpec, mContractStore, lData, mBatchOrders
    
    gInputPaused = True
ElseIf Not mContractProcessor Is Nothing Then
    If pAction = OrderActionBuy Then
        If mContractProcessor.ProcessBuyCommand(pParams) And Not mBatchOrders Then gProcessOrders
    Else
        If mContractProcessor.ProcessSellCommand(pParams) And Not mBatchOrders Then gProcessOrders
    End If
Else
    gWriteErrorLine "No contract has been specified in this group"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ProcessSellCommand( _
                ByVal pParams As String)
Const ProcName As String = "processBuyCommand"
On Error GoTo Err

processBuyOrSellCommand OrderActionSell, pParams

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processCloseoutCommand( _
                ByVal pParams As String)
Const ProcName As String = "processCloseoutCommand"
On Error GoTo Err

If UCase$(pParams) = All Then
    mCloseoutProcessor.CloseoutAll
    gInputPaused = True
ElseIf UCase$(pParams) = DefaultOrderGroupName Then
    mCloseoutProcessor.CloseoutGroup DefaultOrderGroupName
    gInputPaused = True
ElseIf pParams <> "" Then
    If Not isGroupValid(pParams) Then
        gWriteErrorLine "Invalid group name: first character must be letter or digit; remaining characters must be letter, digit, hyphen or underscore"
    Else
        mCloseoutProcessor.CloseoutGroup pParams
        gInputPaused = True
    End If
Else
    gWriteErrorLine CloseoutCommand & " parameter must be a group name or ALL"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processContractCommand(ByVal pParams As String)
Const ProcName As String = "processContractCommand"
On Error GoTo Err

If Trim$(pParams) = "" Then Exit Sub

If Trim$(pParams) = HelpCommand Or Trim$(pParams) = Help1Command Then
    showContractHelp
    Exit Sub
End If

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pParams, " ")

Dim lSpecString As String
lSpecString = lClp.Arg(0)

Dim lContractSpec As IContractSpecifier
If lSpecString <> "" Then
    Set lContractSpec = CreateContractSpecifierFromString(lSpecString)
Else
    Set lContractSpec = parseContractSpec(lClp)
End If

If lContractSpec Is Nothing Then
    If Not mContractProcessor Is Nothing Then
        gSetValidNextCommands ListCommand, ContractCommand, BuyCommand, SellCommand, EndOrdersCommand, ResetCommand
    Else
        gSetValidNextCommands ListCommand, ContractCommand, EndOrdersCommand, ResetCommand
    End If
    Exit Sub
End If

Dim lResolver As New ContractResolver
lResolver.Initialise lContractSpec, mContractStore, Nothing, mBatchOrders

gInputPaused = True

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Sub

Private Sub processEndOrdersCommand()
Const ProcName As String = "processEndOrdersCommand"
On Error GoTo Err

If mErrorCount <> 0 Then
    gWriteLineToConsole mErrorCount & " errors have been found - no orders will be placed"
    mErrorCount = 0
    gSetValidNextCommands ListCommand, GroupCommand, ContractCommand, BracketCommand, BuyCommand, SellCommand, StageOrdersCommand, BatchOrdersCommand, ResetCommand, CloseoutCommand
    Exit Sub
End If

If getNumberOfUnprocessedOrders = 0 Then
    gWriteLineToConsole "No orders have been defined"
    Exit Sub
End If

gProcessOrders

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processGroupCommand( _
                ByVal pParams As String)
If pParams = "" Or pParams = DefaultOrderGroupName Then
    mGroupName = DefaultOrderGroupName
ElseIf Not isGroupValid(pParams) Then
    gWriteErrorLine "Invalid group name: first character must be letter or digit; remaining characters must be letter, digit, hyphen or underscore"
Else
    mGroupName = pParams
End If
If Not mGroupContractProcessors.TryItem(mGroupName, mContractProcessor) Then Set mContractProcessor = Nothing

If Not mContractProcessor Is Nothing Then
    gSetValidNextCommands ListCommand, ContractCommand, BracketCommand, BuyCommand, SellCommand, EndOrdersCommand, ResetCommand
Else
    gSetValidNextCommands ListCommand, ContractCommand, BuyCommand, SellCommand, EndOrdersCommand, ResetCommand
End If
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
mStageOrders = mStageOrdersDefault
mBatchOrders = mBatchOrdersDefault
mErrorCount = 0
gSetValidNextCommands ListCommand, StageOrdersCommand, BatchOrdersCommand, GroupCommand, ContractCommand, BuyCommand, SellCommand, ResetCommand, CloseoutCommand
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

gSetValidNextCommands ListCommand, GroupCommand, ContractCommand, BuyCommand, SellCommand, StageOrdersCommand, BatchOrdersCommand, ResetCommand, CloseoutCommand
End Sub

Private Function setBatchOrders() As Boolean
setBatchOrders = True
If mClp.SwitchValue(BatchOrdersSwitch) = "" Then
    mBatchOrdersDefault = False
ElseIf UCase$(mClp.SwitchValue(BatchOrdersSwitch)) = Yes Then
    mBatchOrdersDefault = True
ElseIf UCase$(mClp.SwitchValue(BatchOrdersSwitch)) = No Then
    mBatchOrdersDefault = False
Else
    gWriteErrorLine "The /" & BatchOrdersSwitch & " switch has an invalid value: it must be either YES or NO (not case-sensitive)"
    setBatchOrders = False
End If
End Function

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
                ByVal pSimulateOrders As Boolean, _
                ByVal pLogApiMessages As Boolean) As Boolean
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
Set lTwsClient = GetClient(server, CLng(port), CLng(clientId), pLogTwsMessages:=pLogApiMessages)

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
gCon.WriteLineToConsole "    buycommand  ::= buy [<contract>] <quantity> <entryordertype> "
gCon.WriteLineToConsole "                        [<price> [<triggerprice]]"
gCon.WriteLineToConsole "                        [/<bracketattr> | /<orderattr>]... NEWLINE"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "    sellcommand ::= sell [<contract>] <quantity> <entryordertype> "
gCon.WriteLineToConsole "                         [/<bracketattr> | /<orderattr>]... NEWLINE"
gCon.WriteLineToConsole ""
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
gCon.WriteLineToConsole "[<buycommand>|<sellcommand>|<bracketcommand>]..."
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
gCon.WriteLineToConsole "       [-simulateorders] [-logapimessages]"
gCon.WriteLineToConsole "       [-scopename:<scope>] [-recoveryfiledir:<recoverypath>]"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "  where"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "    twsserver  ::= STRING name or IP address of computer where TWS/Gateway"
gCon.WriteLineToConsole "                          is running"
gCon.WriteLineToConsole "    port       ::= INTEGER port to be used for API connection"
gCon.WriteLineToConsole "    clientid   ::= INTEGER client id >=0 to be used for API connection (default"
gCon.WriteLineToConsole "                           value is " & DefaultClientId & ")"
gCon.WriteLineToConsole "    resultspath ::= path to the folder in which results files are to be created"
gCon.WriteLineToConsole "                    (defaults to the logfile path)"
gCon.WriteLineToConsole "    logfilepath ::= path to the folder where the program logfile is to be"
gCon.WriteLineToConsole "                    created"
gCon.WriteLineToConsole "    scope       ::= order recovery scope name"
gCon.WriteLineToConsole "    recoverypath ::= path to the folder in which order recovery files are to"
gCon.WriteLineToConsole "                     be created(defaults to the logfile path)"
gCon.WriteLineToConsole ""
showStdInHelp
gCon.WriteLineToConsole ""
showStdOutHelp
gCon.WriteLineToConsole ""
End Sub



