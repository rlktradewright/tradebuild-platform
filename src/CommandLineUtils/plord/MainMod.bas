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

Private Const ApiMessageLoggingSwitch               As String = "APIMESSAGELOGGING"
Private Const BatchOrdersSwitch                     As String = "BATCHORDERS"
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
Public Const CloseSwitch                            As String = "CLOSE"
Public Const DaysSwitch                             As String = "DAYS"
Public Const DescriptionSwitch                      As String = "DESCRIPTION"
Public Const EntrySwitch                            As String = "ENTRY"
Public Const OffsetSwitch                           As String = "OFFSET"
Public Const IgnoreRTHSwitch                        As String = "IGNORERTH"
Public Const PriceSwitch                            As String = "PRICE"
Public Const ReasonSwitch                           As String = "REASON"
Public Const TIFSwitch                              As String = "TIF"
Public Const TimeSwitch                             As String = "TIME"
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

Public Const GroupsSubcommand                       As String = "GROUPS"
Public Const PositionsSubcommand                    As String = "POSITIONS"
Public Const TradesSubcommand                       As String = "TRADES"

Public Const ValueSeparator                         As String = ":"
Public Const SwitchPrefix                           As String = "/"

Private Const All                                   As String = "ALL"
Private Const Default                               As String = "DEFAULT"
Private Const Yes                                   As String = "YES"
Private Const No                                    As String = "NO"

Public Const TickOffsetDesignator                   As String = "T"
Public Const PercentOffsetDesignator                As String = "%"

' Legacy pseudo-order types from early versions
Public Const AskPseudoOrderType                     As String = "ASK"
Public Const BidPseudoOrderType                     As String = "BID"
Public Const LastPseudoOrderType                    As String = "Last"
Public Const AutoPseudoOrderType                    As String = "AUTO"

Private Const DefaultClientId                       As Long = 906564398

Private Const DefaultGroupName                      As String = "$"
Private Const AllGroups                             As String = "ALL"

Private Const DefaultPrompt                         As String = ":"

Private Const CloseoutMarket                        As String = "MKT"
Private Const CloseoutLimit                         As String = "LMT"

'@================================================================================
' Member variables
'@================================================================================

Public gCon                                         As Console

Public gLogToConsole                                As Boolean

' This flag is set when various asynchronous operations are being performed, and
' we' don't want any further user input until we know whether it has succeeded.
Public gInputPaused                                 As Boolean

Public gNumberOfOrdersPlaced                        As Long

Public gPlaceOrdersTask                             As PlaceOrdersTask

Public gBracketOrderListener                        As New BracketOrderListener

Public gCommands                                    As New Commands
Public gCommandListAlways                           As New CommandList
Public gCommandListOrderCreation                    As New CommandList
Public gCommandListOrderSpecification               As New CommandList
Public gCommandListOrderCompletion                  As New CommandList
Public gCommandListGeneral                          As New CommandList


Private mErrorCount                                 As Long

Private mFatalErrorHandler                          As FatalErrorHandler

Private mContractStore                              As IContractStore
Private mMarketDataManager                          As IMarketDataManager

Private mScopeName                                  As String
Private mOrderManager                               As OrderManager
Private mOrderSubmitterFactory                      As IOrderSubmitterFactory

Private mStageOrdersDefault                         As Boolean
Private mStageOrders                                As Boolean

Private mCurrentGroup                               As GroupResources
Private mGroups                                     As Groups

Private mCommandNumber                              As Long

Private mNextCommands                               As New NextCommands

Private mClp                                        As CommandLineParser

Private mMonitor                                    As Boolean

Private mRecoveryFileDir                            As String

Private mOrderPersistenceDataStore                  As IOrderPersistenceDataStore
Private mOrderRecoveryAgent                         As IOrderRecoveryAgent

Private mConfigStore                                As ConfigurationStore

Private mBatchOrdersDefault                         As Boolean
Private mBatchOrders                                As Boolean

Private mBracketOrderDefinitionInProgress           As Boolean

Private mClientId                                   As Long

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

Public Property Get gErrorCount() As Long
gErrorCount = mErrorCount
End Property

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

Public Function gGenerateContractProcessorName( _
                ByVal pGroupName As String, _
                ByVal pContractSpec As IContractSpecifier) As String
Const PositionManagerNameSeparator As String = "&&"

gGenerateContractProcessorName = UCase$(pGroupName & _
                                        PositionManagerNameSeparator & _
                                        pContractSpec.Key)
End Function

Public Function gGenerateSwitch(ByVal pName As String, ByVal pValue As String) As String
If InStr(1, pValue, " ") <> 0 Then pValue = """" & pValue & """"
gGenerateSwitch = SwitchPrefix & pName & IIf(pValue = "", " ", ValueSeparator & pValue & " ")
End Function

Public Function gGetContractName(ByVal pcontract As IContract) As String
AssertArgument Not pcontract Is Nothing
gGetContractName = pcontract.Specifier.LocalSymbol & "@" & pcontract.Specifier.Exchange
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
ReDim lCommandLists(UBound(values)) As CommandList
Dim i As Long
For i = 0 To UBound(values)
    Set lCommandLists(i) = values(i)
Next
mNextCommands.SetValidNextCommandLists lCommandLists
End Sub

Public Sub gWriteErrorLine( _
                ByVal pMessage As String, _
                Optional ByVal pDontIncrementErrorCount As Boolean = False)
Const ProcName As String = "gWriteErrorLine"

Dim s As String
s = "Error: " & pMessage
gCon.WriteErrorLine s
LogMessage "StdErr: " & s
If Not pDontIncrementErrorCount Then mErrorCount = mErrorCount + 1
End Sub

Public Sub gWriteLineToConsole(ByVal pMessage As String, Optional ByVal pLogit As Boolean)
Const ProcName As String = "gWriteLineToConsole"

If pLogit Then LogMessage "Con: " & pMessage
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

If Trim$(Command) = "/?" Or Trim$(Command) = "-?" Then
    showUsage
    Exit Sub
End If

InitialiseTWUtilities

Set mFatalErrorHandler = New FatalErrorHandler
ApplicationGroupName = "TradeWright"
ApplicationName = "plord"
SetupDefaultLogging Command, True, True

Set mConfigStore = gGetConfigStore

logProgramId

setupCommandLists

Set mClp = CreateCommandLineParser(Command)

Dim lLogApiMessages As ApiMessageLoggingOptions
Dim lLogRawApiMessages As ApiMessageLoggingOptions
Dim lLogApiMessageStats As Boolean
If Not validateApiMessageLogging( _
                mClp.SwitchValue(ApiMessageLoggingSwitch), _
                lLogApiMessages, _
                lLogRawApiMessages, _
                lLogApiMessageStats) Then
    gWriteLineToConsole "API message logging setting is invalid", True
    Exit Sub
End If

If Not setupTwsApi(mClp.SwitchValue(TwsSwitch), _
                mClp.Switch(SimulateOrdersSwitch), _
                lLogApiMessages, _
                lLogRawApiMessages, _
                lLogApiMessageStats) Then
    showUsage
    Exit Sub
End If

mScopeName = mClp.SwitchValue(ScopeNameSwitch)

Set mGroups = New Groups
mGroups.Initialise mContractStore, _
                    mMarketDataManager, _
                    mOrderManager, _
                    mScopeName, _
                    mOrderSubmitterFactory

Set gPlaceOrdersTask = New PlaceOrdersTask
gPlaceOrdersTask.Initialise mGroups
StartTask gPlaceOrdersTask, PriorityNormal

process

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getInputLine() As String
Const ProcName As String = "getInputLine"
On Error GoTo Err

Do While gInputPaused
    Wait 20
Loop
getInputLine = Trim$(gCon.ReadLine(getPrompt))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getNumberOfUnprocessedOrders() As Long
Const ProcName As String = "getNumberOfUnprocessedOrders"
On Error GoTo Err

Dim lProcessor As ContractProcessor
Dim lGroupResources As GroupResources
For Each lGroupResources In mGroups
    For Each lProcessor In lGroupResources.ContractProcessors
        getNumberOfUnprocessedOrders = getNumberOfUnprocessedOrders + lProcessor.BracketOrders.Count
    Next
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getPrompt() As String
If mCurrentGroup.CurrentContractProcessor Is Nothing Then
    getPrompt = mCurrentGroup.GroupName & DefaultPrompt
Else
    getPrompt = mCurrentGroup.GroupName & "!" & _
                mCurrentGroup.CurrentContractProcessor.ContractName & _
                DefaultPrompt
End If
End Function

Private Function isGroupValid(ByVal pGroup As String) As Boolean
Const ProcName As String = "isGroupValid"
On Error GoTo Err

gRegExp.Global = True
gRegExp.IgnoreCase = True
gRegExp.Pattern = "^(\$|ALL|[a-zA-Z0-9][\w-]*)$"
isGroupValid = gRegExp.Test(pGroup)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isProhibitedGroupName(ByVal pGroup As String) As Boolean
Const ProcName As String = "isProhibitedGroupName"
On Error GoTo Err

pGroup = UCase$(pGroup)

isProhibitedGroupName = True
If pGroup = AllGroups Then Exit Function

Dim lOrderType As OrderTypes: lOrderType = OrderTypes.OrderTypeNone
lOrderType = OrderTypeFromString(pGroup)
If lOrderType <> OrderTypeNone Then Exit Function

isProhibitedGroupName = False

Exit Function

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then Resume Next
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub listGroups()
Const ProcName As String = "listGroups"
On Error GoTo Err

Dim lRes As GroupResources
For Each lRes In mGroups
    Dim lGroupName As String: lGroupName = lRes.GroupName
    Dim lContractProcessor As ContractProcessor
    Set lContractProcessor = lRes.CurrentContractProcessor
    
    Dim lContractName As String
    If lContractProcessor Is Nothing Then
        lContractName = ""
    Else
        lContractName = lContractProcessor.ContractName
    End If
    gWriteLineToConsole IIf(lRes Is mCurrentGroup, "* ", "  ") & _
                        padStringRight(lGroupName, 20) & _
                        padStringRight(lContractName, 25), _
                        True
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
    Dim lContract As IContract
    Set lContract = lPM.ContractFuture.Value
    gWriteLineToConsole padStringRight(lPM.GroupName, 15) & " " & _
                        padStringRight(gGetContractName(lContract), 30) & _
                        " Size=" & padStringleft(lPM.PositionSize & _
                                                "(" & lPM.PendingBuyPositionSize & _
                                                "/" & _
                                                lPM.PendingSellPositionSize & ")", 10) & _
                        " Profit=" & padStringleft(Format(lPM.Profit, "0.00"), 9), _
                        True
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
                        padStringRight(gGetContractName(lContract), 30), _
                        True
    
    Dim lTrade As Execution
    For Each lTrade In lPM.Executions
        gWriteLineToConsole "  " & FormatTimestamp(lTrade.FillTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & " " & _
                        padStringRight(IIf(lTrade.Action = OrderActionBuy, "BUY", "SELL"), 5) & _
                        padStringleft(lTrade.Quantity, 5) & _
                        padStringleft(FormatPrice(lTrade.Price, lContract.Specifier.SecType, lContract.TickSize), 9) & _
                        padStringRight(" " & lTrade.FillingExchange, 10) & _
                        " " & lTrade.TimezoneName, _
        True
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
s = App.ProductName & _
    " V" & _
    App.Major & _
    "." & App.Minor & _
    "." & App.Revision & _
    IIf(App.FileDescription <> "", "." & App.FileDescription, "") & _
    vbCrLf & _
    App.LegalCopyright
gWriteLineToConsole s, False
s = s & vbCrLf & "Arguments: " & Command
LogMessage s

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
                ByVal pClp As CommandLineParser, _
                ByRef pMaxExpenditure As Long, _
                ByRef pUnderlyingExchange As String) As IContractSpecifier
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
    If IsValidExpiry(lExpiry) Then
    ElseIf IsDate(lExpiry) Then
        lExpiry = Format(CDate(lExpiry), "yyyymmdd")
    ElseIf Len(lExpiry) = 6 Then
        If Not IsDate(Left$(lExpiry, 4) & "/" & Right$(lExpiry, 2) & "/01") Then
            gWriteErrorLine "Invalid Expiry '" & lExpiry & "'"
            validParams = False
        End If
    ElseIf Len(lExpiry) = 8 Then
        If Not IsDate(Left$(lExpiry, 4) & "/" & mId$(lExpiry, 5, 2) & "/" & Right$(lExpiry, 2)) Then
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
    ElseIf parseStrikeExtension(lStrike, pMaxExpenditure, pUnderlyingExchange) Then
        Strike = 0#
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

Private Function parseLegacyCloseoutCommand( _
                ByVal pParams As String, _
                ByRef pGroupName As String, _
                ByRef pCloseoutMode As CloseoutModes, _
                ByRef pPriceSspec As PriceSpecifier) As Boolean
Const ProcName As String = "parseLegacyCloseoutCommand"
On Error GoTo Err

gRegExp.Global = False
gRegExp.IgnoreCase = True

Dim p As String: p = _
    "(?:" & _
        "^" & _
        "(?!mkt)(?!lmt)" & _
        "(?:" & _
            "(?: *)|(all|\$)|([a-zA-Z0-9][\w-]*)" & _
        ")" & _
        "(?:" & _
            " +" & _
            "(?:" & _
                "(mkt)|" & _
                "(?:" & _
                    "(lmt)" & _
                    "(?:" & _
                        ":(-)?(\d{1,3})" & _
                    ")?" & _
                ")" & _
            ")" & _
        ")?" & _
        "$" & _
    ")" & _
    "|"
p = p & _
    "(?:" & _
        "^" & _
        " *" & _
        "(?:" & _
            "(mkt)|" & _
            "(?:" & _
                "(lmt)" & _
                "(?:" & _
                    ":(-)?(\d{1,3})" & _
                ")?" & _
            ")" & _
        ")" & _
        "$" & _
    ")"
gRegExp.Pattern = p

Dim lMatches As MatchCollection
Set lMatches = gRegExp.Execute(Trim$(pParams))

If lMatches.Count <> 1 Then
    Exit Function
End If

Dim lMatch As Match: Set lMatch = lMatches(0)


If UCase$(lMatch.SubMatches(0)) = AllGroups Then
    pGroupName = AllGroups
ElseIf lMatch.SubMatches(1) = DefaultGroupName Then
    pGroupName = DefaultGroupName
Else
    pGroupName = UCase$(lMatch.SubMatches(1))
End If
    
Dim lUseLimitOrders As Boolean: lUseLimitOrders = (UCase$(lMatch.SubMatches(3)) = CloseoutLimit Or _
                                                    UCase$(lMatch.SubMatches(7)) = CloseoutLimit)
If lUseLimitOrders Then
    pCloseoutMode = CloseoutModeLimit
    
    Dim lSpreadFactorSign As Long: lSpreadFactorSign = IIf(lMatch.SubMatches(4) = "-" Or lMatch.SubMatches(8) = "-", -1, 1)
    Dim lSpreadFactor As Long
    If lMatch.SubMatches(5) <> "" Then
        lSpreadFactor = lMatch.SubMatches(5) * lSpreadFactorSign
    ElseIf lMatch.SubMatches(9) <> "" Then
        lSpreadFactor = lMatch.SubMatches(9) * lSpreadFactorSign
    End If
    
    Set pPriceSspec = NewPriceSpecifier(pPriceType:=PriceValueTypeBidOrAsk, _
                                        pOffset:=lSpreadFactor, _
                                        pOffsetType:=PriceOffsetTypeBidAskPercent)
Else
    pCloseoutMode = CloseoutModeMarket
End If

parseLegacyCloseoutCommand = True

Exit Function

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description, True
    Exit Function
End If
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function parseStrikeExtension( _
                ByVal pValue As String, _
                ByRef pMaxExpenditure As Long, _
                ByRef pUnderlyingExchange As String) As Boolean
Const ProcName As String = "parseStrikeExtension"
On Error GoTo Err

Const MaxExpenditure As Long = 9999999
Const StrikeFormat As String = "^(?:(?:([1-9]\d{1,6})\$(?:(?:;|,)([a-zA-Z0-9]+))?)?)?$"

gRegExp.Pattern = StrikeFormat

Dim lMatches As MatchCollection
Set lMatches = gRegExp.Execute(Trim$(pValue))

If lMatches.Count <> 1 Then Exit Function

Dim lResult As Boolean: lResult = True
Dim lMatch As Match: Set lMatch = lMatches(0)

Dim lMaxExpenditure As String
lMaxExpenditure = lMatch.SubMatches(0)
If lMaxExpenditure = "" Then
    pMaxExpenditure = 0
ElseIf IsInteger(lMaxExpenditure, 0, MaxExpenditure) Then
    pMaxExpenditure = CInt(lMaxExpenditure)
Else
    lResult = False
End If

pUnderlyingExchange = lMatch.SubMatches(1)

parseStrikeExtension = lResult

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub process()
Const ProcName As String = "process"
On Error GoTo Err

If Not setMonitor Then Exit Sub
If Not setStageOrders Then Exit Sub
If Not setBatchOrders Then Exit Sub

setOrderRecovery

gSetValidNextCommands gCommandListAlways, _
                    gCommandListGeneral, _
                    gCommandListOrderCreation

Dim inString As String: inString = getInputLine

Do While inString <> gCon.EofString
    If inString = "" Then
        ' ignore blank lines - and don't write to the log because
        ' the FileAutoReader program sends blank lines very frequently
    ElseIf Left$(inString, 1) = "#" Then
        ' ignore comments except log and write to console
        LogMessage "StdIn: " & inString
        gWriteLineToConsole inString
    Else
        LogMessage "StdIn: " & inString
        If Not processCommand(inString) Then Exit Do
    End If
    
    inString = getInputLine
Loop

If Not mOrderPersistenceDataStore Is Nothing Then mOrderPersistenceDataStore.Finish
Set mOrderPersistenceDataStore = Nothing

gWriteLineToConsole "Exiting", True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processCommand(ByVal pInstring As String) As Boolean
Const ProcName As String = "processCommand"
On Error GoTo Err

Dim lCommandName As String
lCommandName = UCase$(Split(pInstring, " ")(0))

Dim Params As String
Params = Trim$(Right$(pInstring, Len(pInstring) - Len(lCommandName)))

Dim lCommand As Command
Set lCommand = gCommands.ParseCommand(lCommandName)

If lCommand Is gCommands.ExitCommand Then
    processCommand = False
    Exit Function
End If

Dim lContractProcessor As ContractProcessor: Set lContractProcessor = mCurrentGroup.CurrentContractProcessor

If lCommand Is gCommands.Help1Command Then
    gCon.WriteLine "Valid commands at this point are: " & mNextCommands.ValidCommandNames
ElseIf lCommand Is gCommands.HelpCommand Then
    showStdInHelp
ElseIf Not mNextCommands.IsCommandValid(lCommand) Then
    gWriteErrorLine "Valid commands at this point are: " & mNextCommands.ValidCommandNames, Not mBracketOrderDefinitionInProgress
ElseIf lCommand Is gCommands.ContractCommand Then
    processContractCommand Params, lContractProcessor
ElseIf lCommand Is gCommands.BatchOrdersCommand Then
    processBatchOrdersCommand Params
ElseIf lCommand Is gCommands.StageOrdersCommand Then
    processStageOrdersCommand Params
ElseIf lCommand Is gCommands.GroupCommand Then
    processGroupCommand Params, lContractProcessor
ElseIf lCommand Is gCommands.BuyCommand Then
    setupResultsLogging mClp
    ProcessBuyCommand Params, lContractProcessor
ElseIf lCommand Is gCommands.BuyAgainCommand Then
    ProcessBuyAgainCommand Params, lContractProcessor
ElseIf lCommand Is gCommands.SellCommand Then
    setupResultsLogging mClp
    ProcessSellCommand Params, lContractProcessor
ElseIf lCommand Is gCommands.SellAgainCommand Then
    ProcessSellAgainCommand Params, lContractProcessor
ElseIf lCommand Is gCommands.BracketCommand Then
    mBracketOrderDefinitionInProgress = True
    lContractProcessor.ProcessBracketCommand Params
ElseIf lCommand Is gCommands.EntryCommand Then
    lContractProcessor.ProcessEntryCommand Params
ElseIf lCommand Is gCommands.StopLossCommand Then
    lContractProcessor.ProcessStopLossCommand Params
ElseIf lCommand Is gCommands.TargetCommand Then
    lContractProcessor.ProcessTargetCommand Params
ElseIf lCommand Is gCommands.RolloverCommand Then
    lContractProcessor.ProcessRolloverCommand Params
ElseIf lCommand Is gCommands.QuitCommand Then
    lContractProcessor.ProcessQuitCommand
    mErrorCount = 0
ElseIf lCommand Is gCommands.EndBracketCommand Then
    mBracketOrderDefinitionInProgress = False
    lContractProcessor.ProcessEndBracketCommand
    If mErrorCount = 0 And Not mBatchOrders Then
        setupResultsLogging mClp
        processOrders
    Else
        gWriteLineToConsole mErrorCount & " errors have been found - order will not be placed", True
        mErrorCount = 0
    End If
ElseIf lCommand Is gCommands.EndOrdersCommand Then
    gWriteLineToStdOut lCommandName
    processEndOrdersCommand
ElseIf lCommand Is gCommands.ResetCommand Then
    processResetCommand
ElseIf lCommand Is gCommands.ListCommand Then
    processListCommand Params
ElseIf lCommand Is gCommands.CloseoutCommand Then
    setupResultsLogging mClp
    processCloseoutCommand Params
ElseIf lCommand Is gCommands.QuoteCommand Then
    processQuoteCommand Params, lContractProcessor
ElseIf lCommand Is gCommands.PurgeCommand Then
    processPurgeCommand Params
Else
    gWriteErrorLine "Invalid command '" & Command & "'", True
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
    gWriteErrorLine gCommands.BatchOrdersCommand.Name & " parameter must be either YES or NO or DEFAULT", True
End Select
End Sub

Private Sub ProcessBuyCommand( _
                ByVal pParams As String, _
                ByVal pContractProcessor As ContractProcessor)
Const ProcName As String = "processBuyCommand"
On Error GoTo Err

processBuyOrSellCommand OrderActionBuy, pParams, pContractProcessor

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ProcessBuyAgainCommand( _
                ByVal pParams As String, _
                ByVal pContractProcessor As ContractProcessor)
Const ProcName As String = "ProcessBuyAgainCommand"
On Error GoTo Err

If pContractProcessor Is Nothing Then
    gWriteErrorLine "No Buy command to repeat", True
ElseIf pContractProcessor.LatestBuyCommandParams = "" Then
    gWriteErrorLine "No Buy command to repeat", True
Else
    ProcessBuyCommand pContractProcessor.LatestBuyCommandParams, pContractProcessor
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processBuyOrSellCommand( _
                ByVal pAction As OrderActions, _
                ByVal pParams As String, _
                ByVal pContractProcessor As ContractProcessor)
Const ProcName As String = "processBuyOrSellCommand"
On Error GoTo Err

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pParams)
Dim lArg0 As String: lArg0 = lClp.Arg(0)

If Not IsInteger(lArg0, 1) Then
    ' the first arg is a contract spec
    Dim lContractSpec As IContractSpecifier
    Set lContractSpec = CreateContractSpecifierFromString(lArg0)

    If lContractSpec Is Nothing Then Exit Sub
    
    Set pContractProcessor = mCurrentGroup.AddContractProcessor(lContractSpec, _
                                        mBatchOrders, _
                                        mStageOrders, _
                                        0, _
                                        "")
    
    pParams = Right$(pParams, Len(pParams) - Len(lArg0))
End If

If Not pContractProcessor Is Nothing Then
    If pAction = OrderActionBuy Then
        If pContractProcessor.ProcessBuyCommand(pParams) And Not mBatchOrders Then processOrders
    Else
        If pContractProcessor.ProcessSellCommand(pParams) And Not mBatchOrders Then processOrders
    End If
Else
    gWriteErrorLine "No contract has been specified in this group", True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ProcessSellCommand( _
                ByVal pParams As String, _
                ByVal pContractProcessor As ContractProcessor)
Const ProcName As String = "processBuyCommand"
On Error GoTo Err

processBuyOrSellCommand OrderActionSell, pParams, pContractProcessor

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ProcessSellAgainCommand( _
                ByVal pParams As String, _
                ByVal pContractProcessor As ContractProcessor)
Const ProcName As String = "ProcessSellAgainCommand"
On Error GoTo Err

If pContractProcessor Is Nothing Then
    gWriteErrorLine "No Sell command to repeat", True
ElseIf pContractProcessor.LatestSellCommandParams = "" Then
    gWriteErrorLine "No Sell command to repeat", True
Else
    ProcessSellCommand pContractProcessor.LatestSellCommandParams, pContractProcessor
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processCloseoutCommand( _
                ByVal pParams As String)
Const ProcName As String = "processCloseoutCommand"
On Error GoTo Err

Dim lLimitOrder As String: lLimitOrder = OrderTypeToShortString(OrderTypeLimit)
Dim lMarketOrder As String: lMarketOrder = OrderTypeToShortString(OrderTypeMarket)

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pParams)

Dim lGroupName As String
Dim lOrderTypeName As String
Dim lPriceStr As String

Dim lCloseoutMode As CloseoutModes
Dim lPriceSpec As PriceSpecifier

Dim lError As Boolean

If parseLegacyCloseoutCommand(pParams, lGroupName, lCloseoutMode, lPriceSpec) Then
ElseIf lClp.NumberOfArgs = 3 Then
    lGroupName = lClp.Arg(0)
    lOrderTypeName = UCase$(lClp.Arg(1))
    lPriceStr = lClp.Arg(2)
    
    If lOrderTypeName = lMarketOrder Then
        lCloseoutMode = CloseoutModeMarket
    ElseIf lOrderTypeName = lLimitOrder Then
        lCloseoutMode = CloseoutModeLimit
    Else
        lError = True
        gWriteErrorLine "Second argument must be either 'MKT' or 'LMT'", True
    End If
ElseIf lClp.NumberOfArgs = 2 Then
    lOrderTypeName = UCase$(lClp.Arg(0))
    If lOrderTypeName = lLimitOrder Then
        lCloseoutMode = CloseoutModeLimit
        lPriceStr = lClp.Arg(1)
    ElseIf lOrderTypeName = lMarketOrder Then
        lCloseoutMode = CloseoutModeMarket
        lPriceStr = lClp.Arg(1)
    Else
        lGroupName = lClp.Arg(0)
        lOrderTypeName = UCase$(lClp.Arg(1))
        If lOrderTypeName = lMarketOrder Then lCloseoutMode = CloseoutModeMarket
    End If
ElseIf lClp.NumberOfArgs = 1 Then
    lOrderTypeName = UCase$(lClp.Arg(0))
    If lOrderTypeName = lMarketOrder Then
        lCloseoutMode = CloseoutModeMarket
    ElseIf lOrderTypeName = lLimitOrder Then
        lCloseoutMode = CloseoutModeLimit
    Else
        lGroupName = lClp.Arg(0)
    End If
ElseIf lClp.NumberOfArgs = 0 Then
    lCloseoutMode = CloseoutModeMarket
Else
    lError = True
    gWriteErrorLine "Too many arguments", True
End If

If lGroupName = "" Then
    lGroupName = mCurrentGroup.GroupName
ElseIf UCase$(lGroupName) = AllGroups Then
    lGroupName = AllGroups
ElseIf lGroupName = DefaultGroupName Then
ElseIf Not isGroupValid(lGroupName) Then
    lError = True
    gWriteErrorLine lGroupName & " is not a valid group name", True
ElseIf Not mGroups.Contains(lGroupName) Then
    lError = True
    gWriteErrorLine "No such group", True
End If

If lCloseoutMode = CloseoutModeMarket Then
    If lPriceStr <> "" Then
        lError = True
        gWriteErrorLine "A price specification cannot be supplied for closeout at market"
    End If
ElseIf lPriceStr = "" And lPriceSpec Is Nothing Then
    lError = True
    gWriteErrorLine "Closeout price specification is missing", True
ElseIf Not lPriceSpec Is Nothing Then
ElseIf Not ParsePriceAndOffset(lPriceSpec, lPriceStr, SecTypeNone, 0#, True) Then
    lError = True
    gWriteErrorLine "Invalid price specification", True
End If

If lError Then Exit Sub

Dim lCloseoutProcessor As New CloseoutProcessor
lCloseoutProcessor.Initialise mOrderManager, lCloseoutMode, lPriceSpec

If lGroupName = AllGroups Then
    lCloseoutProcessor.CloseoutAll
ElseIf Not mGroups.Contains(lGroupName) Then
    gWriteErrorLine "No such group", True
Else
    lCloseoutProcessor.CloseoutGroup lGroupName
End If

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description, True
    Exit Sub
End If
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processContractCommand( _
                ByVal pParams As String, _
                ByVal pContractProcessor As ContractProcessor)
Const ProcName As String = "processContractCommand"
On Error GoTo Err

If pParams = "" Then Exit Sub

pParams = UCase$(pParams)
If gCommands.HelpCommand.parse(pParams) Or gCommands.Help1Command.parse(pParams) Then
    showContractHelp
    Exit Sub
End If

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pParams, " ")

Dim lSpecString As String
lSpecString = lClp.Arg(0)

Dim lContractSpec As IContractSpecifier
Dim lMaxExpenditure As Long
Dim lUnderlyingExchange As String
If lSpecString <> "" Then
    Set lContractSpec = CreateContractSpecifierFromString(lSpecString)
Else
    Set lContractSpec = parseContractSpec(lClp, lMaxExpenditure, lUnderlyingExchange)
End If

If lContractSpec Is Nothing Then
    If Not pContractProcessor Is Nothing Then
        gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation
    Else
        gSetValidNextCommands gCommandListAlways, gCommandListGeneral
    End If
    Exit Sub
End If

mCurrentGroup.AddContractProcessor lContractSpec, _
                                    mBatchOrders, _
                                    mStageOrders, _
                                    lMaxExpenditure, _
                                    lUnderlyingExchange

gInputPaused = True

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description, True
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Sub

Private Sub processEndOrdersCommand()
Const ProcName As String = "processEndOrdersCommand"
On Error GoTo Err

If mErrorCount <> 0 Then
    gWriteLineToConsole mErrorCount & " errors have been found - no orders will be placed", True
    mErrorCount = 0
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation
    Exit Sub
End If

If getNumberOfUnprocessedOrders = 0 Then
    gWriteLineToConsole "No orders have been defined", True
    Exit Sub
End If

processOrders

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processGroupCommand( _
                ByVal pParams As String, _
                ByVal pContractProcessor As ContractProcessor)
Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pParams)

Dim lGroupName As String: lGroupName = lClp.Arg(0)
If isProhibitedGroupName(lGroupName) Then
    gWriteErrorLine "Invalid group name: you can't create a group called '" & lGroupName & "'", True
    Exit Sub
ElseIf Not isGroupValid(lGroupName) Then
    gWriteErrorLine "Invalid group name: first character must be letter or digit; remaining characters must be letter, digit, hyphen or underscore", True
    Exit Sub
End If

If Not mGroups.TryItem(UCase$(lGroupName), mCurrentGroup) Then
    Dim lPMs As PositionManagers: Set lPMs = mOrderManager.GetPositionManagersForGroup(lGroupName)
    
    ' get the 'canonical' group name spelling
    If lPMs.Count <> 0 Then
        Dim en As Enumerator: Set en = lPMs.Enumerator
        en.MoveNext
        Dim lPM As PositionManager: Set lPM = en.Current
        lGroupName = lPM.GroupName
    End If
    Set mCurrentGroup = mGroups.Add(lGroupName)
End If

If lClp.NumberOfArgs > 1 Or lClp.NumberOfSwitches > 0 Then
    Dim lContractArgs As String
    lContractArgs = Trim$(Right$(pParams, Len(pParams) - InStr(1, pParams, " ")))
    processContractCommand lContractArgs, pContractProcessor
End If

gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation
End Sub

Private Sub processOrders()
Const ProcName As String = "processOrders"
On Error GoTo Err

' To avoid exceeding the API's input message limits, we process orders
' asynchronously with a task

gPlaceOrdersTask.AddContractProcessors mCurrentGroup.ContractProcessors, mStageOrders

gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processListCommand( _
                ByVal pParams As String)
Const ProcName As String = "processListCommand"
On Error GoTo Err

If UCase$(pParams) = GroupsSubcommand Then
    listGroups
ElseIf UCase$(pParams) = PositionsSubcommand Then
    listPositions
ElseIf UCase$(pParams) = TradesSubcommand Then
    listTrades
Else
    gWriteErrorLine gCommands.ListCommand.Name & " parameter must be one of " & GroupsSubcommand & ", " & PositionsSubcommand & " or " & TradesSubcommand, True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processPurgeCommand( _
                ByVal pParams As String)
Const ProcName As String = "processPurgeCommand"
On Error GoTo Err

Dim lRes As GroupResources

If pParams = "" Then
    gWriteLineToConsole "Purging " & mCurrentGroup.GroupName
    mCurrentGroup.Purge
    mGroups.Remove mCurrentGroup.GroupName
    Set mCurrentGroup = mGroups.Add(DefaultGroupName)
ElseIf UCase$(pParams) = AllGroups Then
    For Each lRes In mGroups
        gWriteLineToConsole "Purging " & lRes.GroupName
        lRes.Purge
    Next
    mGroups.Clear
    Set mCurrentGroup = mGroups.Add(DefaultGroupName)
ElseIf isGroupValid(pParams) Then
    If Not mGroups.TryItem(pParams, lRes) Then
        gWriteErrorLine "No such group", True
        Exit Sub
    End If
    gWriteLineToConsole "Purging " & lRes.GroupName
    lRes.Purge
    mGroups.Remove pParams
    Set mCurrentGroup = mGroups.Add(DefaultGroupName)
Else
    gWriteErrorLine "Invalid group name", True
    Exit Sub
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processQuoteCommand( _
                ByVal pParams As String, _
                ByVal pContractProcessor As ContractProcessor)
Const ProcName As String = "processQuoteCommand"
On Error GoTo Err

If pParams = "" Then
    If Not pContractProcessor Is Nothing Then
        gWriteLineToConsole pContractProcessor.ContractName & _
                            ": " & _
                            GetCurrentTickSummary(pContractProcessor.DataSource)
    Else
        gWriteLineToConsole "No current contract"
    End If
Else
    Dim lClp As CommandLineParser
    Set lClp = CreateCommandLineParser(pParams, " ")
    
    Dim lSpecString As String: lSpecString = lClp.Arg(0)
    
    Dim lContractSpec As IContractSpecifier
    If lSpecString <> "" Then
        Set lContractSpec = CreateContractSpecifierFromString(lSpecString)
    Else
        Dim lMaxExpenditure As Long
        Dim lUnderlyingExchangeName As String
        Set lContractSpec = parseContractSpec(lClp, lMaxExpenditure, lUnderlyingExchangeName)
    End If
    If lContractSpec Is Nothing Then Exit Sub
    
    Dim lQuoteFetcher As New QuoteFetcher
    lQuoteFetcher.FetchQuote lContractSpec, mContractStore, mMarketDataManager
End If

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description, True
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Sub

Private Sub processResetCommand()
mStageOrders = mStageOrdersDefault
mBatchOrders = mBatchOrdersDefault
mErrorCount = 0
gSetValidNextCommands gCommandListAlways, gCommandListGeneral
End Sub

Private Sub processStageOrdersCommand( _
                ByVal pParams As String)
Select Case UCase$(pParams)
Case Default
    mStageOrders = mStageOrdersDefault
Case Yes
    If Not (mOrderSubmitterFactory.Capabilities And OrderSubmitterCapabilityCanStageOrders) = OrderSubmitterCapabilityCanStageOrders Then
        gWriteErrorLine gCommands.StageOrdersCommand.Name & " parameter cannot be YES with current configuration", True
        Exit Sub
    End If
    mStageOrders = True
Case No
    mStageOrders = False
Case Else
    gWriteErrorLine gCommands.StageOrdersCommand.Name & " parameter must be either YES or NO or DEFAULT", True
End Select

gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation
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
    gWriteErrorLine "The /" & BatchOrdersSwitch & " switch has an invalid value: it must be either YES or NO (not case-sensitive)", True
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
    gWriteErrorLine "The /" & MonitorSwitch & " switch has an invalid value: it must be either YES or NO (not case-sensitive, default is YES)", True
    setMonitor = False
End If
End Function

Private Sub setOrderRecovery()
Const ProcName As String = "setOrderRecovery"
On Error GoTo Err

If mScopeName = "" Then Exit Sub

Set mCurrentGroup = mGroups.Add(DefaultGroupName)
mRecoveryFileDir = ApplicationSettingsFolder
If mClp.SwitchValue(RecoveryFileDirSwitch) <> "" Then mRecoveryFileDir = mClp.SwitchValue(RecoveryFileDirSwitch)

If mOrderPersistenceDataStore Is Nothing Then Set mOrderPersistenceDataStore = CreateOrderPersistenceDataStore(mRecoveryFileDir)

Dim lOrderRecoverer As New OrderRecoverer
lOrderRecoverer.RecoverOrders mOrderManager, mScopeName, mOrderPersistenceDataStore, mOrderRecoveryAgent, mMarketDataManager, mOrderSubmitterFactory, mGroups

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
        gWriteErrorLine "The /" & StageOrdersSwitch & " switch has an invalid value: it cannot be YES with the current configuration", True
        setStageOrders = False
        Exit Function
    End If
    mStageOrdersDefault = True
ElseIf UCase$(mClp.SwitchValue(StageOrdersSwitch)) = No Then
    mStageOrdersDefault = False
Else
    gWriteErrorLine "The /" & StageOrdersSwitch & " switch has an invalid value: it must be either YES or NO (not case-sensitive)", True
    setStageOrders = False
End If
End Function

Private Sub setupCommandLists()
gCommandListAlways.Initialise _
                gCommands.ExitCommand, _
                gCommands.HelpCommand, _
                gCommands.Help1Command, _
                gCommands.ListCommand, _
                gCommands.PurgeCommand, _
                gCommands.QuoteCommand, _
                gCommands.ResetCommand, _
                gCommands.StageOrdersCommand

gCommandListOrderCreation.Initialise _
                gCommands.BracketCommand, _
                gCommands.BuyAgainCommand, _
                gCommands.BuyCommand, _
                gCommands.EndOrdersCommand, _
                gCommands.SellAgainCommand, _
                gCommands.SellCommand

gCommandListOrderSpecification.Initialise _
                gCommands.ContractCommand, _
                gCommands.EntryCommand, _
                gCommands.QuitCommand, _
                gCommands.RolloverCommand, _
                gCommands.StopLossCommand, _
                gCommands.TargetCommand

gCommandListOrderCompletion.Initialise _
                gCommands.EndBracketCommand

gCommandListGeneral.Initialise _
                gCommands.CloseoutCommand, _
                gCommands.ContractCommand, _
                gCommands.EndOrdersCommand, _
                gCommands.GroupCommand
End Sub

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
lFilenameSuffix = FormatTimestamp(GetTimestamp, TimestampDateOnly + TimestampNoMillisecs)

Dim lLogfile As FileLogListener
Set lLogfile = CreateFileLogListener( _
                    lResultsPath & "Logs\" & _
                        ProjectName & _
                        "(" & mClientId & ")" & _
                        "-" & lFilenameSuffix & _
                        ".log", _
                    includeTimestamp:=True, _
                    includeLogLevel:=False)
GetLogger("log").AddLogListener lLogfile
GetLogger("position.order").AddLogListener lLogfile
GetLogger("position.simulatedorder").AddLogListener lLogfile

If mMonitor Then
    Set lLogfile = CreateFileLogListener( _
                    lResultsPath & "Orders\" & _
                        ProjectName & _
                        "(" & mClientId & ")" & _
                        "-" & lFilenameSuffix & _
                        "-Executions" & _
                        ".log", _
                    includeTimestamp:=False, _
                    includeLogLevel:=False)
    GetLogger("position.orderdetail").AddLogListener lLogfile
    
    Set lLogfile = CreateFileLogListener( _
                    lResultsPath & "Orders\" & _
                        ProjectName & _
                        "(" & mClientId & ")" & _
                        "-" & lFilenameSuffix & _
                        "-BracketOrders" & ".log", _
                    includeTimestamp:=False, _
                    includeLogLevel:=False)
    GetLogger("position.bracketorderprofilestring").AddLogListener lLogfile
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setupTwsApi( _
                ByVal SwitchValue As String, _
                ByVal pSimulateOrders As Boolean, _
                ByVal pLogApiMessages As ApiMessageLoggingOptions, _
                ByVal pLogRawApiMessages As ApiMessageLoggingOptions, _
                ByVal pLogApiMessageStats As Boolean) As Boolean
Const ProcName As String = "setupTwsApi"
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
    gWriteErrorLine "port must be an integer > 0", True
    setupTwsApi = False
End If
    
If clientId = "" Then
    clientId = DefaultClientId
ElseIf Not IsInteger(clientId, 0, 999999999) Then
    gWriteErrorLine "clientId must be an integer >= 0 and <= 999999999", True
    setupTwsApi = False
Else
    mClientId = CLng(clientId)
End If

Dim lTwsClient As Client
Set lTwsClient = GetClient(server, _
                        CLng(port), _
                        mClientId, _
                        pLogApiMessages:=pLogApiMessages, _
                        pLogRawApiMessages:=pLogRawApiMessages, _
                        pLogApiMessageStats:=pLogApiMessageStats)

Set mContractStore = lTwsClient.GetContractStore
Set mMarketDataManager = CreateRealtimeDataManager(lTwsClient.GetMarketDataFactory)
mMarketDataManager.LoadFromConfig gGetMarketDataSourcesConfig(mConfigStore)

Set mOrderRecoveryAgent = lTwsClient

If pSimulateOrders Then
    Set mOrderSubmitterFactory = New SimOrderSubmitterFactory
Else
    Set mOrderSubmitterFactory = lTwsClient
End If
    
Set mOrderManager = New OrderManager
mOrderManager.ContractStorePrimary = mContractStore
mOrderManager.MarketDataManager = mMarketDataManager
mOrderManager.OrderSubmitterFactory = mOrderSubmitterFactory

setupTwsApi = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub showCloseoutHelp()
gWriteLineToConsole "    closeoutcommand  ::= closeout [ <groupname> | ALL ]"
gWriteLineToConsole "                                  [ LMT[:<percentofspread>] ]"
gWriteLineToConsole ""
gWriteLineToConsole "    percentofspread  ::= INTEGER"
gWriteLineToConsole ""
End Sub

Public Sub showContractHelp()
gWriteLineToConsole "    contractcommand  ::= contract <contractspec>EOL"
gWriteLineToConsole ""
gWriteLineToConsole "    contractspec     ::= <localsymbol>[@<exchangename>]"
gWriteLineToConsole "                         | <localsymbol>@SMART/<routinghint>"
gWriteLineToConsole "                         | /<specifier>[;/<specifier>]..."
gWriteLineToConsole "    routinghint ::= STRING"
gWriteLineToConsole "    specifier ::=   local[symbol]:<localsymbol>"
gWriteLineToConsole "                  | symb[ol]:<symbol>"
gWriteLineToConsole "                  | sec[type]:<sectype>"
gWriteLineToConsole "                  | exch[ange]:<exchangename>"
gWriteLineToConsole "                  | curr[ency]:<currencycode>"
gWriteLineToConsole "                  | exp[iry]:<yyyymm> | <yyyymmdd> | <expiryoffset>"
gWriteLineToConsole "                  | mult[iplier]:<multiplier>"
gWriteLineToConsole "                  | str[ike]:<price>"
gWriteLineToConsole "                  | right:CALL|PUT"
gWriteLineToConsole "    localsymbol ::= STRING"
gWriteLineToConsole "    symbol ::= STRING"
gWriteLineToConsole "    sectype ::= STK | FUT | FOP | CASH | OPT"
gWriteLineToConsole "    exchangename ::= STRING"
gWriteLineToConsole "    currencycode ::= USD | EUR | GBP | JPY | CHF | etc"
gWriteLineToConsole "    yyyymm       ::= <yyyy><mm>"
gWriteLineToConsole "    yyyymmdd     ::= <yyyy><mm><dd>"
gWriteLineToConsole "    yyyy         ::= INTEGER(1900..2100)"
gWriteLineToConsole "    mm           ::= INTEGER(01..12)"
gWriteLineToConsole "    dd           ::= INTEGER(01..31)"
gWriteLineToConsole "    expiryoffset ::= INTEGER(0..10)"
End Sub

Private Sub showListHelp()
gWriteLineToConsole "    listcommand  ::= list groups|positions|trades"
End Sub

Private Sub showOrderHelp()
gWriteLineToConsole "    buycommand  ::= buy [<contract>] <quantity> <entryordertype> "
gWriteLineToConsole "                        [<priceoroffset> [<triggerprice]]"
gWriteLineToConsole "                        [<attribute>]... EOL"
gWriteLineToConsole ""
gWriteLineToConsole "    sellcommand ::= sell [<contract>] <quantity> <entryordertype> "
gWriteLineToConsole "                         [<priceoroffset> [<triggerprice]]"
gWriteLineToConsole "                         [<attribute>]... EOL"
gWriteLineToConsole ""
gWriteLineToConsole "    bracketcommand ::= bracket <action> <quantity> [<bracketattr>]... EOL"
gWriteLineToConsole "                       entry <entryordertype> [<orderattr>]...  EOL"
gWriteLineToConsole "                       [stoploss <stoplossorderType> [<orderattr>]...  EOL]"
gWriteLineToConsole "                       [target <targetorderType> [<orderattr>]...  ] EOL"
gWriteLineToConsole "                       endbracket EOL"
gWriteLineToConsole ""
gWriteLineToConsole "    action     ::= buy | sell"
gWriteLineToConsole "    quantity   ::= INTEGER >= 1"
gWriteLineToConsole "    entryordertype  ::=   mkt"
gWriteLineToConsole "                        | lmt"
gWriteLineToConsole "                        | stp"
gWriteLineToConsole "                        | stplmt"
gWriteLineToConsole "                        | mit"
gWriteLineToConsole "                        | lit"
gWriteLineToConsole "                        | moc"
gWriteLineToConsole "                        | loc"
gWriteLineToConsole "                        | trail"
gWriteLineToConsole "                        | traillmt"
gWriteLineToConsole "                        | mtl"
gWriteLineToConsole "                        | moo"
gWriteLineToConsole "                        | loo"
gWriteLineToConsole "                        | bid"
gWriteLineToConsole "                        | ask"
gWriteLineToConsole "                        | last"
gWriteLineToConsole "    stoplossordertype  ::=   mkt"
gWriteLineToConsole "                           | stp"
gWriteLineToConsole "                           | stplmt"
gWriteLineToConsole "                           | auto"
gWriteLineToConsole "                           | trail"
gWriteLineToConsole "                           | traillmt"
gWriteLineToConsole "    targetordertype  ::=   mkt"
gWriteLineToConsole "                         | lmt"
gWriteLineToConsole "                         | mit"
gWriteLineToConsole "                         | lit"
gWriteLineToConsole "                         | auto"
gWriteLineToConsole "    attribute  ::= <bracketattr> | <orderattr>"
gWriteLineToConsole "    bracketattr  ::=  /cancelafter:<canceltime>"
gWriteLineToConsole "                     | /cancelprice:<price>"
gWriteLineToConsole "                     | /description:STRING"
gWriteLineToConsole "                     | /goodaftertime:DATETIME"
gWriteLineToConsole "                     | /goodtilldate:DATETIME"
gWriteLineToConsole "                     | /timezone:TIMEZONENAME"
gWriteLineToConsole "    orderattr  ::=   /price:<price>"
gWriteLineToConsole "                   | /reason:STRING"
gWriteLineToConsole "                   | /trigger[price]:<price>"
gWriteLineToConsole "                   | /trailby:<numberofticks>"
gWriteLineToConsole "                   | /trailpercent:<percentage>"
gWriteLineToConsole "                   | /offset:<points>"
gWriteLineToConsole "                   | /offset:<numberofticks>T"
gWriteLineToConsole "                   | /offset:<bidaskspreadpercent>%"
gWriteLineToConsole "                   | /tif:<tifvalue>"
gWriteLineToConsole "                   | /ignorerth"
gWriteLineToConsole "    priceoroffset ::= <price> | <offset>"
gWriteLineToConsole "    price  ::= DOUBLE"
gWriteLineToConsole "    offset ::= DOUBLE"
gWriteLineToConsole "    points ::= DOUBLE"
gWriteLineToConsole "    bidaskspreadpercent ::= [+|-]INTEGEER"
gWriteLineToConsole "    numberofticks  ::= [+|-]INTEGER"
gWriteLineToConsole "    percentage  ::= DOUBLE <= 10.0"
gWriteLineToConsole "    tifvalue  ::=   DAY"
gWriteLineToConsole "                  | GTC"
gWriteLineToConsole "                  | IOC"

End Sub

Private Sub showReverseHelp()
gWriteLineToConsole "    reversecommand   ::= reverse [ <groupname> | ALL ])"
End Sub

Private Sub showStdInHelp()
gWriteLineToConsole "StdIn Format:"
gWriteLineToConsole ""
gWriteLineToConsole "#comment EOL"
gWriteLineToConsole ""
gWriteLineToConsole "[stageorders [yes|no|default] EOL]"
gWriteLineToConsole ""
gWriteLineToConsole "[group [<groupname>] EOL]"
gWriteLineToConsole ""
gWriteLineToConsole "<contractcommand>"
gWriteLineToConsole ""
gWriteLineToConsole "[<buycommand>|<sellcommand>|<bracketcommand>]..."
gWriteLineToConsole ""
gWriteLineToConsole "endorders EOL"
gWriteLineToConsole ""
gWriteLineToConsole "reset EOL"
gWriteLineToConsole ""
gWriteLineToConsole "<listcommand>"
gWriteLineToConsole ""
gWriteLineToConsole "<closeoutcommand>"
gWriteLineToConsole ""
gWriteLineToConsole "<reversecommand>"
gWriteLineToConsole ""
gWriteLineToConsole "quote <contractspecification>"
gWriteLineToConsole ""
gWriteLineToConsole "where"
gWriteLineToConsole ""
showContractHelp
gWriteLineToConsole ""
showOrderHelp
gWriteLineToConsole ""
showListHelp
gWriteLineToConsole ""
showCloseoutHelp
gWriteLineToConsole ""
showReverseHelp
End Sub

Private Sub showStdOutHelp()
gWriteLineToConsole "StdOut Format:"
gWriteLineToConsole ""
gWriteLineToConsole "CONTRACT <contractdetails>"
gWriteLineToConsole "TIME <timestamp>"
gWriteLineToConsole "<orderdetails>"
gWriteLineToConsole "<bracketorderdetails>"
gWriteLineToConsole "ENDORDERS"
gWriteLineToConsole ""
gWriteLineToConsole "  where"
gWriteLineToConsole ""
gWriteLineToConsole "    contractdetails      ::= same format as for contract input"
gWriteLineToConsole "    timestamp            ::= yyyy-mm-dd hh:mm:ss"
gWriteLineToConsole "    orderdetails         ::= Same format as single order input"
gWriteLineToConsole "    bracketorderdetails  ::= Same format as bracket order input"
End Sub

Private Sub showUsage()
gWriteLineToConsole "Usage:"
gWriteLineToConsole "plord27 -tws[:[<twsserver>][,[<port>][,[<clientid>]]]] [-monitor[:[yes|no]]] "
gWriteLineToConsole "       [-resultsdir:<resultspath>] [-stopAt:<hh:mm>] [-log:<logfilepath>]"
gWriteLineToConsole "       [-loglevel:[ I | N | D | M | H }]"
gWriteLineToConsole "       [-stageorders[:[yes|no]]]"
gWriteLineToConsole "       [-simulateorders] [-apimessagelogging:[D|A|N][D|A|N][Y|N]]"
gWriteLineToConsole "       [-scopename:<scope>] [-recoveryfiledir:<recoverypath>]"
gWriteLineToConsole ""
gWriteLineToConsole "  where"
gWriteLineToConsole ""
gWriteLineToConsole "    twsserver  ::= STRING name or IP address of computer where TWS/Gateway"
gWriteLineToConsole "                          is running"
gWriteLineToConsole "    port       ::= INTEGER port to be used for API connection"
gWriteLineToConsole "    clientid   ::= INTEGER client id >=0 to be used for API connection (default"
gWriteLineToConsole "                           value is " & DefaultClientId & ")"
gWriteLineToConsole "    resultspath ::= path to the folder in which results files are to be created"
gWriteLineToConsole "                    (defaults to the logfile path)"
gWriteLineToConsole "    logfilepath ::= path to the folder where the program logfile is to be"
gWriteLineToConsole "                    created"
gWriteLineToConsole "    scope       ::= order recovery scope name"
gWriteLineToConsole "    recoverypath ::= path to the folder in which order recovery files are to"
gWriteLineToConsole "                     be created(defaults to the logfile path)"
gWriteLineToConsole ""
showStdInHelp
gWriteLineToConsole ""
showStdOutHelp
gWriteLineToConsole ""
End Sub

Private Function validateApiMessageLogging( _
                ByVal pApiMessageLogging As String, _
                ByRef pLogApiMessages As ApiMessageLoggingOptions, _
                ByRef pLogRawApiMessages As ApiMessageLoggingOptions, _
                ByRef pLogApiMessageStats As Boolean) As Boolean
Const Always As String = "A"
Const Default As String = "D"
Const No As String = "N"
Const None As String = "N"
Const Yes As String = "Y"

pApiMessageLogging = UCase$(pApiMessageLogging)

validateApiMessageLogging = False
If Len(pApiMessageLogging) = 0 Then pApiMessageLogging = Default & Default & No
If Len(pApiMessageLogging) <> 3 Then Exit Function

Dim s As String
s = Left$(pApiMessageLogging, 1)
If s = None Then
    pLogApiMessages = ApiMessageLoggingOptionNone
ElseIf s = Default Then
    pLogApiMessages = ApiMessageLoggingOptionDefault
ElseIf s = Always Then
    pLogApiMessages = ApiMessageLoggingOptionAlways
Else
    Exit Function
End If

s = mId(pApiMessageLogging, 2, 1)
If s = None Then
    pLogRawApiMessages = ApiMessageLoggingOptionNone
ElseIf s = Default Then
    pLogRawApiMessages = ApiMessageLoggingOptionDefault
ElseIf s = Always Then
    pLogRawApiMessages = ApiMessageLoggingOptionAlways
Else
    Exit Function
End If

s = mId(pApiMessageLogging, 3, 1)
If s = No Then
    pLogApiMessageStats = False
ElseIf s = Yes Then
    pLogApiMessageStats = True
Else
    Exit Function
End If

validateApiMessageLogging = True
End Function

