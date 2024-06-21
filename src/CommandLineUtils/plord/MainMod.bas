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

Public Enum ErrorCountIncrementCriterion
    ErrorCountIncrementNo
    ErrorCountIncrementYes
    ErrorCountIncrementIfNotInteractive
End Enum

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
Public Const QuantitySwitch                         As String = "QUANTITY"
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
Public Const TradingClassSwitch                     As String = "TRADINGCLASS"

Public Const GroupsSubcommand                       As String = "GROUPS"
Public Const OrdersSubcommand                       As String = "ORDERS"
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

Public Const MaxOrderCostSuffix                     As String = "$"
Public Const AccountPercentSuffix                   As String = "%"

' Legacy pseudo-order types from early versions
Public Const AskPseudoOrderType                     As String = "ASK"
Public Const BidPseudoOrderType                     As String = "BID"
Public Const LastPseudoOrderType                    As String = "Last"
Public Const AutoPseudoOrderType                    As String = "AUTO"

Private Const DefaultClientId                       As Long = 906564398

Private Const DefaultGroupName                      As String = "$"
Private Const AllGroups                             As String = "ALL"

Private Const DefaultPrompt                         As String = ">"

Private Const CloseoutMarket                        As String = "MKT"
Private Const CloseoutLimit                         As String = "LMT"

Private Const StrikeSelectionModeNone               As String = ""
Private Const StrikeSelectionModeIncrement          As String = "I"
Private Const StrikeSelectionModeExpenditure        As String = "$"
Private Const StrikeSelectionModeDelta              As String = "D"

Private Const AccountValueAvailableFunds            As String = "AvailableFunds"
Private Const AccountValueCashBalance               As String = "CashBalance"
Private Const AccountValueEquityWithLoan            As String = "EquityWithLoan"
Private Const AccountValueExcessLiquidity           As String = "ExcessLiquidity"
Private Const AccountValueNetLiquidation            As String = "NetLiquidation"

'@================================================================================
' Member variables
'@================================================================================

Public gCon                                         As Console

Public gLogToConsole                                As Boolean

' This flag is set when various asynchronous operations are being performed, and
' we' don't want any further user input until we know whether it has succeeded.
Public gInputPaused                                 As Boolean

Public gPlaceOrdersTask                             As PlaceOrdersTask

Public gBracketOrderListener                        As New BracketOrderListener

Public gCommands                                    As New Commands
Public gCommandListAlways                           As New CommandList
Public gCommandListOrderCreation                    As New CommandList
Public gCommandListOrderSpecification               As New CommandList
Public gCommandListOrderCompletion                  As New CommandList
Public gCommandListGeneral                          As New CommandList

Public gDefaultBalanceAccountValueName              As String

Public gLiveOrders                                  As New LiveOrdersCache



Private mBlockingErrorCount                         As Long
Private mErrorCount                                 As Long

Private mFatalErrorHandler                          As FatalErrorHandler

Private mContractStore                              As IContractStore
Private mMarketDataManager                          As RealTimeDataManager

Private mScopeName                                  As String
Private mOrderManager                               As OrderManager
Private mOrderSubmitterFactory                      As IOrderSubmitterFactory

Private mAccountDataProvider                        As IAccountDataProvider
Private mCurrencyConverter                          As ICurrencyConverter

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

Private mTwsClient                                  As Client
Private mClientId                                   As Long

Private mMoneyManager                               As New MoneyManager

Private mTerminateRequested                         As Boolean

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

Public Property Get gBlockingErrorCount() As Long
gBlockingErrorCount = mBlockingErrorCount
End Property

Public Property Get gErrorCount() As Long
gErrorCount = mErrorCount
End Property

Public Property Get gRegExp() As RegExp
Static sRegexp As RegExp
If sRegexp Is Nothing Then
    Set sRegexp = New RegExp
    sRegexp.IgnoreCase = True
End If
Set gRegExp = sRegexp
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub gCompleteOrderRecovery()
Const ProcName As String = "gCompleteOrderRecovery"
On Error GoTo Err

Dim lPM As PositionManager
For Each lPM In mOrderManager.PositionManagersLive
    If lPM.IsActive Then processGroupCommand lPM.GroupName, Nothing
    Dim lBO As IBracketOrder
    For Each lBO In lPM.BracketOrders
        If lBO.State = BracketOrderStateSubmitted Then
            If lBO.Size <> 0 Then CreateBracketProfitCalculator lBO, lPM.DataSource
            Dim lEntry As LiveOrderEntry: Set lEntry = New LiveOrderEntry
            lEntry.Key = lBO.Key
            lEntry.GroupName = lBO.GroupName
            lEntry.Order = lBO
            'lEntry.Timestamp = lBO.CreationTime
            gLiveOrders.Add lEntry
        End If
        If Not lBO.RolloverSpecification Is Nothing And _
            (lBO.CumBuyPrice <> 0 Or _
            lBO.CumSellPrice <> 0) _
        Then
            gBracketOrderListener.Add lBO
        End If
    Next
Next

Set mCurrentGroup = mGroups.Add(DefaultGroupName)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


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
Const ProcName As String = "gGetContractName"
On Error GoTo Err

AssertArgument Not pcontract Is Nothing, "pcontract is Nothing"
gGetContractName = pcontract.Specifier.LocalSymbol & "@" & pcontract.Specifier.Exchange

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Public Function gIsInteractive() As Boolean
gIsInteractive = (gCon.StdInType = FileTypeChar)
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

Public Function gPadStringleft(ByRef pInput As String, ByVal pLength As Long) As String
Dim lInput As String: lInput = Right$(pInput, pLength)
gPadStringleft = Space$(pLength - Len(lInput)) & lInput
End Function

Public Function gPadStringRight(ByRef pInput As String, ByVal pLength As Long) As String
Dim lInput As String: lInput = Left$(pInput, pLength)
gPadStringRight = lInput & Space$(pLength - Len(lInput))
End Function

Public Sub gSetValidNextCommands(ParamArray values() As Variant)
ReDim lCommandLists(UBound(values)) As CommandList
Dim i As Long
For i = 0 To UBound(values)
    Set lCommandLists(i) = values(i)
Next
mNextCommands.SetValidNextCommandLists lCommandLists
End Sub

Public Function gStrikeSelectionModeToString(ByVal pMode As OptionStrikeSelectionModes) As String
Select Case pMode
Case OptionStrikeSelectionModeNone
    gStrikeSelectionModeToString = StrikeSelectionModeNone
Case OptionStrikeSelectionModeIncrement
    gStrikeSelectionModeToString = StrikeSelectionModeIncrement
Case OptionStrikeSelectionModeExpenditure
    gStrikeSelectionModeToString = StrikeSelectionModeExpenditure
Case OptionStrikeSelectionModeDelta
    gStrikeSelectionModeToString = StrikeSelectionModeDelta
End Select
End Function

Public Sub gTerminate( _
                ByVal pMessage As String)
gWriteLineToConsole pMessage
mTerminateRequested = True
End Sub

Public Sub gWriteErrorLine( _
                ByVal pMessage As String, _
                Optional ByVal pIncrementCriterion As ErrorCountIncrementCriterion = ErrorCountIncrementYes)
Dim s As String
s = "Error: " & pMessage
gCon.WriteErrorLine s
LogMessage "StdErr: " & s
Select Case pIncrementCriterion
Case ErrorCountIncrementNo

Case ErrorCountIncrementYes
    mBlockingErrorCount = mBlockingErrorCount + 1
    mErrorCount = mErrorCount + 1
Case ErrorCountIncrementIfNotInteractive
    If gIsInteractive Then
        mErrorCount = mErrorCount + 1
    Else
        mBlockingErrorCount = mBlockingErrorCount + 1
    End If
End Select
End Sub

Public Sub gWriteLineToConsole( _
                ByVal pMessage As String, _
                Optional ByVal pIncludeTimestamp As Boolean, _
                Optional ByVal pDontLogit As Boolean = False)
Const ProcName As String = "gWriteLineToConsole"

If Not pDontLogit Then LogMessage "Con: " & pMessage
If pIncludeTimestamp Then
    Dim s As String
    s = FormatTimestamp(GetTimestamp, TimestampTimeOnlyISO8601) & " "
    gCon.WriteStringToConsole s
End If
    
Dim lIndent As Long: lIndent = Len(s)
If InStr(1, pMessage, vbCrLf) = 0 Then
    gCon.WriteLineToConsole pMessage
Else
    Dim ar() As String: ar = Split(pMessage, vbCrLf)
    gCon.WriteLineToConsole ar(0)
    Dim i As Long
    For i = 1 To UBound(ar)
        gCon.WriteLineToConsole Space(lIndent) & ar(i)
    Next
End If
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
Debug.Print "Console StdInType = " & gCon.StdInType

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
    gWriteLineToConsole "API message logging setting is invalid"
    Exit Sub
End If

If Not setupTwsApi(mClp.SwitchValue(TwsSwitch), _
                mClp.Switch(SimulateOrdersSwitch), _
                lLogApiMessages, _
                lLogRawApiMessages, _
                lLogApiMessageStats, _
                mClientId) Then
    showUsage
    Exit Sub
End If

mScopeName = mClp.SwitchValue(ScopeNameSwitch)
If mScopeName = "" Then mScopeName = CStr(mClientId)

gDefaultBalanceAccountValueName = AccountValueNetLiquidation

Set mGroups = New Groups
mGroups.Initialise mContractStore, _
                    mMarketDataManager, _
                    mOrderManager, _
                    mScopeName, _
                    mOrderSubmitterFactory, _
                    mMoneyManager, _
                    mAccountDataProvider, _
                    mCurrencyConverter
Set mCurrentGroup = mGroups.Add(DefaultGroupName)

Set gPlaceOrdersTask = New PlaceOrdersTask
gPlaceOrdersTask.Initialise mGroups, mMoneyManager
StartTask gPlaceOrdersTask, PriorityNormal

If mAccountDataProvider.State <> AccountProviderReady Then
    gWriteLineToConsole "Waiting for account provider to be ready", True
    Do While mAccountDataProvider.State <> AccountProviderReady
        getInputLineAndWait pDontReadInput:=True, pWaitTimeMIllisecs:=50
    Loop
    gWriteLineToConsole "Account provider is ready", True
End If

process

TerminateTWUtilities

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getFundsAmount( _
                ByVal pParams As String, _
                ByRef pAmount As Double) As Boolean
Const ProcName As String = "getFundsAmount"
On Error GoTo Err

getFundsAmount = False

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pParams)
If lClp.NumberOfArgs > 2 Then
    gWriteErrorLine "Too many parameters", ErrorCountIncrementIfNotInteractive
    Exit Function
End If

Dim lFundsSpec As String: lFundsSpec = lClp.Arg(0)
Dim lAccountValueName As String: lAccountValueName = lClp.Arg(1)
If lAccountValueName = "" Then lAccountValueName = gDefaultBalanceAccountValueName

If Not isValidAccountValueName(lAccountValueName) Then Exit Function

Const FundsSpecFormat As String = "^(?:(?:([1-9]?\d*(?:\.\d+)?)(\%))|(?:(0?\.\d+(?:\.\d+)?)(\%))|(?:(?:([1-9]\d*)(\$))))$"
gRegExp.Pattern = FundsSpecFormat

Dim lMatches As MatchCollection
Set lMatches = gRegExp.Execute(lFundsSpec)

If lMatches.Count <> 1 Then
    gWriteErrorLine _
        "Invalid funds specification: must be either an integer greater" & vbCrLf & _
        "than 0 followed by $ (an amount of currency), or a number (decimal" & vbCrLf & _
        "places allowed) greater than 0 followed by % (a percentage of the" & vbCrLf & _
        "account balance). " & vbCrLf & _
        "Note that the funds are in the account's base currency.", ErrorCountIncrementIfNotInteractive
    Exit Function
End If

Dim lMatch As Match: Set lMatch = lMatches(0)

If lMatch.SubMatches(1) = AccountPercentSuffix Then
    pAmount = CDbl(lMatch.SubMatches(0)) * CDbl(mAccountDataProvider.GetAccountValue(lAccountValueName, mAccountDataProvider.BaseCurrency).Value) / 100#
ElseIf lMatch.SubMatches(3) = AccountPercentSuffix Then
    pAmount = CDbl(lMatch.SubMatches(2)) * CDbl(mAccountDataProvider.GetAccountValue(lAccountValueName, mAccountDataProvider.BaseCurrency).Value) / 100#
ElseIf lMatch.SubMatches(5) = MaxOrderCostSuffix Then
    pAmount = CDbl(lMatch.SubMatches(4))
End If

getFundsAmount = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getInputLineAndWait( _
                Optional ByVal pDontReadInput As Boolean = False, _
                Optional ByVal pWaitTimeMIllisecs As Long = 5) As String
Const ProcName As String = "getInputLine"
On Error GoTo Err

If mTerminateRequested Then
    getInputLineAndWait = gCon.EofString
    Exit Function
End If

Dim lWaitUntilTime As Double
lWaitUntilTime = GetTimestampUTC + pWaitTimeMIllisecs / (86400# * 1000#)

If Not pDontReadInput Then getInputLineAndWait = Trim$(gCon.ReadLine(getPrompt))

Do
    ' allow queued system messages to be handled
    Wait 5
Loop Until GetTimestampUTC >= lWaitUntilTime

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

Private Function getOrderSummary(ByVal pOrder As IBracketOrder) As String
Const ProcName As String = "getOrderSummary"
On Error GoTo Err

Dim s As String
With pOrder
    s = gGetContractName(.Contract) & " " & _
            IIf(.EntryOrder.Action = OrderActionBuy, "B ", "S ") & _
            .EntryOrder.Quantity.DecimalValue & " " & _
            OrderTypeToShortString(.EntryOrder.OrderType) & " " & _
            .EntryOrder.LimitPriceSpec.PriceString
    If .EntryOrder.TriggerPriceSpec.IsValid Then _
        s = s & ":" & .EntryOrder.TriggerPriceSpec.PriceString
    
    If .StopLossOrder Is Nothing Then
    Else
        s = s & " (SL: " & OrderTypeToShortString(.StopLossOrder.OrderType) & " "
        If .StopLossOrder.LimitPriceSpec.IsValid Then _
            s = s & .StopLossOrder.LimitPriceSpec.PriceString
        If .StopLossOrder.TriggerPriceSpec.IsValid Then _
            s = s & ";" & .StopLossOrder.TriggerPriceSpec.PriceString
        s = s & ")"
    End If
    
    If .TargetOrder Is Nothing Then
    Else
        s = s & " (T: " & OrderTypeToShortString(.TargetOrder.OrderType) & " "
        If .TargetOrder.LimitPriceSpec.IsValid Then _
            s = s & .TargetOrder.LimitPriceSpec.PriceString
        If .TargetOrder.TriggerPriceSpec.IsValid Then _
            s = s & ";" & .TargetOrder.TriggerPriceSpec.PriceString
        s = s & ")"
    End If
End With
getOrderSummary = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getPrompt() As String
If mCurrentGroup Is Nothing Then
ElseIf mCurrentGroup.CurrentContractProcessor Is Nothing Then
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
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isValidAccountValueName( _
                ByRef Value As String) As Boolean
isValidAccountValueName = True
Select Case UCase$(Value)
Case UCase$(AccountValueAvailableFunds)
    Value = AccountValueAvailableFunds
Case UCase$(AccountValueCashBalance)
    Value = AccountValueCashBalance
Case UCase$(AccountValueEquityWithLoan)
    Value = AccountValueEquityWithLoan
Case UCase$(AccountValueExcessLiquidity)
    Value = AccountValueExcessLiquidity
Case UCase$(AccountValueNetLiquidation)
    Value = AccountValueNetLiquidation
Case Else
    gWriteErrorLine "Invalid account value name - must be one of: " & vbCrLf & _
                    "    AvailableFunds" & vbCrLf & _
                    "    CashBalance" & vbCrLf & _
                    "    EquityWithLoan" & vbCrLf & _
                    "    ExcessLiquidity" & vbCrLf & _
                    "    NetLiquidation", _
                    ErrorCountIncrementIfNotInteractive
    isValidAccountValueName = False
End Select
End Function

Private Sub listGroups()
Const ProcName As String = "listGroups"
On Error GoTo Err

gWriteLineToConsole "Groups at " & _
                    FormatTimestamp(GetTimestamp, _
                                    TimestampDateAndTimeISO8601 Or TimestampNoMillisecs) & _
                    ": "
Dim lRes As GroupResources
For Each lRes In mGroups
    Dim lGroupName As String: lGroupName = lRes.GroupName
    gWriteLineToConsole IIf(lRes Is mCurrentGroup, "* ", "  ") & _
                        gPadStringRight(lGroupName, 20)
    
    
    Dim lContractProcessor As ContractProcessor
    
    For Each lContractProcessor In lRes.ContractProcessors
        gWriteLineToConsole "    " & _
                            IIf(lContractProcessor Is lRes.CurrentContractProcessor, "* ", "  ") & _
                            gPadStringRight(lContractProcessor.ContractName, 25)
    Next
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub listOrders(ByVal pParams As String)
Const ProcName As String = "listOrders"
On Error GoTo Err

Const OrdersAll = "ALL"
Const OrdersPending = "PENDING"
Const OrdersSubmitted = "SUBMITTED"
Const OrdersCompleted = "COMPLETED"
Const OrdersCancelled = "CANCELLED"


Dim lSelector As String: lSelector = UCase$(pParams)

Select Case lSelector
Case ""
Case OrdersAll
Case OrdersPending
Case OrdersSubmitted
Case OrdersCompleted
Case OrdersCancelled
Case Else
    gWriteErrorLine "parameter must be omitted or one of: " & _
                        OrdersAll & ", " & _
                        OrdersPending & ", " & _
                        OrdersSubmitted & " or " & _
                        OrdersCompleted & " or " & _
                        OrdersCancelled, _
                        ErrorCountIncrementNo
    Exit Sub
End Select

Dim lEntry As LiveOrderEntry
Dim lStatus As String
Dim lAllow As Boolean
For Each lEntry In gLiveOrders
    If lEntry.Order Is Nothing Then
        If lEntry.Cancelled Then
            lStatus = OrdersCancelled
        Else
            lStatus = OrdersPending
        End If
    Else
        lStatus = BracketOrderStateToString(lEntry.Order.State)
    End If
    
    Select Case lSelector
    Case "", OrdersAll
        lAllow = True
    Case OrdersPending
        lAllow = (lStatus = OrdersPending Or lStatus = BracketOrderStateToString(BracketOrderStateCreated))
    Case OrdersSubmitted
        lAllow = (lStatus = BracketOrderStateToString(BracketOrderStateSubmitted))
    Case OrdersCompleted
        lAllow = (lStatus = BracketOrderStateToString(BracketOrderStateCancelling) Or _
                    lStatus = BracketOrderStateToString(BracketOrderStateClosingOut) Or _
                    lStatus = BracketOrderStateToString(BracketOrderStateClosed) Or _
                    lStatus = BracketOrderStateToString(BracketOrderStateAwaitingOtherOrderCancel))
    Case OrdersCancelled
        lAllow = (lStatus = OrdersCancelled)
    Case Else
        lAllow = False
    End Select
    
    If lAllow Then
        Dim lOrderSummary As String
        If lEntry.BracketOrderSpec Is Nothing Then
            lOrderSummary = getOrderSummary(lEntry.Order)
        Else
            lOrderSummary = lEntry.BracketOrderSpec.ToSummaryString
        End If
        
        gWriteLineToConsole FormatTimestamp(lEntry.Timestamp, TimestampTimeOnlyISO8601 + TimestampNoMillisecs) & _
                            " " & lEntry.Key & _
                            " (" & lStatus & "): " & _
                            lOrderSummary, _
                        False, True
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub listPositions()
Const ProcName As String = "listPositions"
On Error GoTo Err

Dim lQual As String: lQual = IIf(mOrderManager.PositionManagersLive.Count = 0, "none", "")
gWriteLineToConsole "Positions at " & _
                    FormatTimestamp(GetTimestamp, _
                                    TimestampDateAndTimeISO8601 Or TimestampNoMillisecs) & _
                    ": " & lQual
Dim lPM As PositionManager
For Each lPM In mOrderManager.PositionManagersLive
    Dim lContract As IContract
    Set lContract = lPM.ContractFuture.Value
    gWriteLineToConsole gPadStringRight(lPM.GroupName, 15) & " " & _
                        gPadStringRight(gGetContractName(lContract), 30) & _
                        " Size=" & gPadStringleft(lPM.PositionSize & _
                                                "(" & lPM.PendingBuyPositionSize & _
                                                "/" & _
                                                lPM.PendingSellPositionSize & ")", 10) & _
                        " Profit=" & gPadStringleft(Format(lPM.Profit, "0.00"), 9) & _
                        IIf(lPM.IsActive, "", " (inactive)")
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub listTrades()
Const ProcName As String = "listTrades"
On Error GoTo Err

Dim lQual As String: lQual = IIf(mOrderManager.PositionManagersLive.Count = 0, "none", "")
gWriteLineToConsole "Trades at " & _
                    FormatTimestamp(GetTimestamp, _
                                    TimestampDateAndTimeISO8601 Or TimestampNoMillisecs) & _
                    ": " & lQual
Dim lPM As PositionManager
For Each lPM In mOrderManager.PositionManagersLive
    Dim lContract As IContract
    Set lContract = lPM.ContractFuture.Value
    gWriteLineToConsole gPadStringRight(lPM.GroupName, 15) & " " & _
                        gPadStringRight(gGetContractName(lContract), 30)
    
    Dim lTrade As Execution
    For Each lTrade In lPM.Executions
        gWriteLineToConsole "  " & FormatTimestamp(lTrade.FillTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & " " & _
                        gPadStringRight(IIf(lTrade.Action = OrderActionBuy, "BUY", "SELL"), 5) & _
                        gPadStringleft(lTrade.Quantity, 5) & _
                        gPadStringleft(FormatPrice(lTrade.Price, lContract.Specifier.SecType, lContract.TickSize), 9) & _
                        gPadStringRight(" " & lTrade.FillingExchange, 10) & _
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

Private Function parseContractSpec( _
                ByVal pClp As CommandLineParser, _
                ByRef pStrikeSelectionMode As OptionStrikeSelectionModes, _
                ByRef pParameter As Long, _
                ByRef pOperator As OptionStrikeSelectionOperators, _
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

Dim lTradingClass As String: lTradingClass = pClp.SwitchValue(TradingClassSwitch)

Dim lCurrency As String: lCurrency = pClp.SwitchValue(CurrencySwitch)
If lCurrency = "" Then lCurrency = pClp.SwitchValue(CurrencySwitch1)

Dim lExpiry As String: lExpiry = pClp.SwitchValue(ExpirySwitch)
If lExpiry = "" Then lExpiry = pClp.SwitchValue(ExpirySwitch1)

Dim lMultiplier As String: lMultiplier = pClp.SwitchValue(MultiplierSwitch)
If lMultiplier = "" Then lMultiplier = pClp.SwitchValue(MultiplierSwitch1)
If lMultiplier = "" Then lMultiplier = "0.0"

Dim lStrike As String: lStrike = pClp.SwitchValue(StrikeSwitch)
If lStrike = "" Then lStrike = pClp.SwitchValue(StrikeSwitch1)
If lStrike = "" Then lStrike = "0.0"

Dim lRight As String: lRight = pClp.SwitchValue(RightSwitch)

Dim lSectype As SecurityTypes
lSectype = SecTypeFromString(lSectypeStr)
If lSectypeStr <> "" And lSectype = SecTypeNone Then
    gWriteErrorLine "Invalid Sectype '" & lSectypeStr & "'", ErrorCountIncrementIfNotInteractive
    validParams = False
End If

If lExpiry <> "" Then
    If IsValidExpiry(lExpiry) Then
    ElseIf IsDate(lExpiry) Then
        lExpiry = Format(CDate(lExpiry), "yyyymmdd")
    ElseIf Len(lExpiry) = 6 Then
        If Not IsDate(Left$(lExpiry, 4) & "/" & Right$(lExpiry, 2) & "/01") Then
            gWriteErrorLine "Invalid Expiry '" & lExpiry & "'", ErrorCountIncrementIfNotInteractive
            validParams = False
        End If
    ElseIf Len(lExpiry) = 8 Then
        If Not IsDate(Left$(lExpiry, 4) & "/" & Mid$(lExpiry, 5, 2) & "/" & Right$(lExpiry, 2)) Then
            gWriteErrorLine "Invalid Expiry '" & lExpiry & "'", ErrorCountIncrementIfNotInteractive
            validParams = False
        End If
    Else
        gWriteErrorLine "Invalid Expiry '" & lExpiry & "'", ErrorCountIncrementIfNotInteractive
        validParams = False
    End If
End If
            
Dim Multiplier As Double
If lMultiplier = "" Then
    Multiplier = 0#
ElseIf IsNumeric(lMultiplier) Then
    Multiplier = CDbl(lMultiplier)
Else
    gWriteErrorLine "Invalid multiplier '" & lMultiplier & "'", ErrorCountIncrementIfNotInteractive
    validParams = False
End If
            
Dim optRight As OptionRights
optRight = OptionRightFromString(lRight)
If lRight <> "" And optRight = OptNone Then
    gWriteErrorLine "Invalid right '" & lRight & "'", ErrorCountIncrementIfNotInteractive
    validParams = False
End If

Dim Strike As Double
If lStrike <> "" Then
    If IsNumeric(lStrike) Then
        Strike = CDbl(lStrike)
    ElseIf optRight = OptNone Then
        gWriteErrorLine "Right not specified", ErrorCountIncrementIfNotInteractive
        validParams = False
    ElseIf parseStrikeExtension(lStrike, optRight, pStrikeSelectionMode, pParameter, pOperator, pUnderlyingExchange) Then
        Strike = 0#
    Else
        gWriteErrorLine "Invalid strike '" & lStrike & "'", ErrorCountIncrementIfNotInteractive
        validParams = False
    End If
End If

        
If validParams Then
    Set parseContractSpec = CreateContractSpecifier(lLocalSymbol, _
                                            lSymbol, _
                                            lTradingClass, _
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
    gWriteErrorLine Err.Description, ErrorCountIncrementIfNotInteractive
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
    gWriteErrorLine Err.Description, ErrorCountIncrementIfNotInteractive
    Exit Function
End If
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function parseStrikeExtension( _
                ByVal pValue As String, _
                ByVal pOptRight As OptionRights, _
                ByRef pStrikeSelectionMode As OptionStrikeSelectionModes, _
                ByRef pParameter As Long, _
                ByRef pOperator As OptionStrikeSelectionOperators, _
                ByRef pUnderlyingExchange As String) As Boolean
Const ProcName As String = "parseStrikeExtension"
On Error GoTo Err

Const StrikeFormat As String = "^(?:(?:(<|<=|>|>=|)(\-?[1-9]\d{1,6})(\$|D)(?:(?:;|,)([a-zA-Z0-9]+))?)?)?$"

gRegExp.Pattern = StrikeFormat

Dim lMatches As MatchCollection
Set lMatches = gRegExp.Execute(Trim$(pValue))

If lMatches.Count <> 1 Then Exit Function

Dim lResult As Boolean: lResult = True
Dim lMatch As Match: Set lMatch = lMatches(0)

Dim lOperator As String: lOperator = lMatch.SubMatches(0)
Select Case lOperator
Case ""
    pOperator = OptionStrikeSelectionOperatorNone
Case "<"
    pOperator = OptionStrikeSelectionOperatorLT
Case "<="
    pOperator = OptionStrikeSelectionOperatorLE
Case ">"
    pOperator = OptionStrikeSelectionOperatorGT
Case ">="
    pOperator = OptionStrikeSelectionOperatorGE
End Select

Dim lMinParameter As Long
Dim lMaxParameter As Long
Dim lSelectionMode As String: lSelectionMode = UCase$(lMatch.SubMatches(2))
Select Case lSelectionMode
Case StrikeSelectionModeExpenditure
    pStrikeSelectionMode = OptionStrikeSelectionModeExpenditure
    lMinParameter = 10
    lMaxParameter = 9999999
Case StrikeSelectionModeDelta
    pStrikeSelectionMode = OptionStrikeSelectionModeDelta
    If pOptRight = OptCall Then
        lMinParameter = 1
        lMaxParameter = 100
    Else
        lMinParameter = -100
        lMaxParameter = -1
    End If
Case Else
    pStrikeSelectionMode = OptionStrikeSelectionModeNone
End Select

Dim lParameter As String: lParameter = lMatch.SubMatches(1)
If lParameter = "" Then
    pParameter = 0
ElseIf IsInteger(lParameter, lMinParameter, lMaxParameter) Then
    pParameter = CLng(lParameter)
Else
    lResult = False
End If

pUnderlyingExchange = lMatch.SubMatches(3)

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

setupResultsLogging mClp

setOrderRecovery

gSetValidNextCommands gCommandListAlways, _
                    gCommandListGeneral, _
                    gCommandListOrderCreation

Dim inString As String: inString = getInputLineAndWait(gInputPaused)

Do While inString <> gCon.EofString
    If inString = "" Then
        ' ignore blank lines - and don't write to the log because
        ' the FileAutoReader program sends blank lines very frequently
    ElseIf Left$(inString, 1) = "#" Then
        ' ignore comments except log
        LogMessage "StdIn: " & inString
    Else
        If gIsInteractive Then
            LogMessage "StdIn: " & inString
        Else
            LogMessage "StdIn: " & inString
            gWriteLineToConsole ">" & inString
        End If
        
        Dim lCommandObj As Command
        If Left$(inString, 1) = ":" Then
            gRegExp.Global = True
            gRegExp.IgnoreCase = True
            gRegExp.Pattern = "^:([a-zA-Z][0-9a-zA-Z]*(?:[.|\-|:|_][0-9a-zA-Z]+)*) *((?:BUY|B|SELL|S|BRACKET) *(?:.*))$"
            Dim lMatches As MatchCollection
            Set lMatches = gRegExp.Execute(inString)
            
            If lMatches.Count <> 1 Then
                gWriteErrorLine _
                    "Invalid labelled command", ErrorCountIncrementIfNotInteractive
            Else
                Dim lMatch As Match: Set lMatch = lMatches(0)
                Dim lLabel As String: lLabel = lMatch.SubMatches(0)
                Dim lCommand As String: lCommand = lMatch.SubMatches(1)
                If Not processCommand(lCommand, lCommandObj, lLabel) Then
                    gWriteLineToConsole "Exiting due to unprocessed input"
                    Exit Do
                End If
            End If
        ElseIf processCommand(inString, lCommandObj) Then
        
        ElseIf lCommandObj Is gCommands.ExitCommand Then
            gWriteLineToConsole "Exiting"
            Exit Do
        Else
            gWriteLineToConsole "Exiting due to unprocessed input"
            Exit Do
        End If
    End If
    
    inString = getInputLineAndWait(gInputPaused)
Loop

If mTerminateRequested Then Exit Sub

If Not mOrderPersistenceDataStore Is Nothing Then mOrderPersistenceDataStore.Finish
Set mOrderPersistenceDataStore = Nothing

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

Private Function processCommand( _
                ByVal pInstring As String, _
                ByRef pCommand As Command, _
                Optional ByVal pLabel As String) As Boolean
Const ProcName As String = "processCommand"
On Error GoTo Err

Dim lID As String: lID = pLabel

Dim lCommandName As String
lCommandName = UCase$(Split(pInstring, " ")(0))

Dim Params As String
Params = Trim$(Right$(pInstring, Len(pInstring) - Len(lCommandName)))

Set pCommand = gCommands.ParseCommand(lCommandName)

If pCommand Is gCommands.ExitCommand Then
    processCommand = False
    Exit Function
End If

Dim lContractProcessor As ContractProcessor: Set lContractProcessor = mCurrentGroup.CurrentContractProcessor

If pCommand Is gCommands.Help1Command Then
    gCon.WriteLine "Valid commands at this point are: " & mNextCommands.ValidCommandNames
ElseIf pCommand Is gCommands.HelpCommand Then
    showStdInHelp
ElseIf Not mNextCommands.IsCommandValid(pCommand) Then
    gWriteErrorLine "Valid commands at this point are: " & mNextCommands.ValidCommandNames, _
                    IIf(mBracketOrderDefinitionInProgress, ErrorCountIncrementIfNotInteractive, ErrorCountIncrementNo)
ElseIf pCommand Is gCommands.ContractCommand Then
    processContractCommand Params, lContractProcessor
ElseIf pCommand Is gCommands.BatchOrdersCommand Then
    processBatchOrdersCommand Params
ElseIf pCommand Is gCommands.StageOrdersCommand Then
    processStageOrdersCommand Params
ElseIf pCommand Is gCommands.GroupCommand Then
    processGroupCommand Params, lContractProcessor
ElseIf pCommand Is gCommands.SetBalanceCommand Then
    processSetBalanceCommand Params
ElseIf pCommand Is gCommands.ShowBalanceCommand Then
    processShowBalanceCommand
ElseIf pCommand Is gCommands.SetFundsCommand Then
    processSetFundsCommand Params
ElseIf pCommand Is gCommands.SetGroupFundsCommand Then
    processSetGroupFundsCommand Params
ElseIf pCommand Is gCommands.SetRolloverCommand Then
    processSetRolloverCommand Params
ElseIf pCommand Is gCommands.SetGroupRolloverCommand Then
    processSetGroupRolloverCommand Params
ElseIf pCommand Is gCommands.BuyCommand Then
    If lID = "" Then lID = GenerateBracketOrderId
    gWriteLineToConsole "Order id: " & lID
    mErrorCount = 0
    If Not mBatchOrders Then mBlockingErrorCount = 0
    ProcessBuyCommand Params, lContractProcessor, lID
ElseIf pCommand Is gCommands.BuyAgainCommand Then
    If lID = "" Then lID = GenerateBracketOrderId
    gWriteLineToConsole "Order id: " & lID
    ProcessBuyAgainCommand Params, lContractProcessor, lID
ElseIf pCommand Is gCommands.SellCommand Then
    If lID = "" Then lID = GenerateBracketOrderId
    gWriteLineToConsole "Order id: " & lID
    mErrorCount = 0
    If Not mBatchOrders Then mBlockingErrorCount = 0
    ProcessSellCommand Params, lContractProcessor, lID
ElseIf pCommand Is gCommands.SellAgainCommand Then
    If lID = "" Then lID = GenerateBracketOrderId
    gWriteLineToConsole "Order id: " & lID
    ProcessSellAgainCommand Params, lContractProcessor, lID
ElseIf pCommand Is gCommands.BracketCommand Then
    mBracketOrderDefinitionInProgress = True
    If lID = "" Then lID = GenerateBracketOrderId
    gWriteLineToConsole "Order id: " & lID
    mErrorCount = 0
    If Not mBatchOrders Then mBlockingErrorCount = 0
    ProcessBracketCommand Params, lContractProcessor, lID, False
ElseIf pCommand Is gCommands.EntryCommand Then
    lContractProcessor.ProcessEntryCommand Params
ElseIf pCommand Is gCommands.StopLossCommand Then
    lContractProcessor.ProcessStopLossCommand Params
ElseIf pCommand Is gCommands.TargetCommand Then
    lContractProcessor.ProcessTargetCommand Params
ElseIf pCommand Is gCommands.RolloverCommand Then
    lContractProcessor.ProcessRolloverCommand Params
ElseIf pCommand Is gCommands.QuitCommand Then
    lContractProcessor.ProcessQuitCommand
    mBlockingErrorCount = 0
    mErrorCount = 0
ElseIf pCommand Is gCommands.EndBracketCommand Then
    mBracketOrderDefinitionInProgress = False
    lContractProcessor.ProcessEndBracketCommand
    
    If mBlockingErrorCount <> 0 Or mErrorCount <> 0 Then
        gWriteLineToConsole gErrorCount & " errors have been found - order will not be placed"
    Else
        processOrders
    End If
ElseIf pCommand Is gCommands.EndOrdersCommand Then
    processEndOrdersCommand
ElseIf pCommand Is gCommands.ModifyCommand Then
    processModifyCommand Params
ElseIf pCommand Is gCommands.Modify1Command Then
    processModifyCommand Params
ElseIf pCommand Is gCommands.Modify2Command Then
    processModifyCommand Params
ElseIf pCommand Is gCommands.CancelCommand Then
    processCancelCommand Params
ElseIf pCommand Is gCommands.ResetCommand Then
    processResetCommand
ElseIf pCommand Is gCommands.ListCommand Then
    processListCommand Params
ElseIf pCommand Is gCommands.CloseoutCommand Then
    processCloseoutCommand Params
ElseIf pCommand Is gCommands.QuoteCommand Then
    processQuoteCommand Params, lContractProcessor
ElseIf pCommand Is gCommands.PurgeCommand Then
    processPurgeCommand Params
Else
    gWriteErrorLine "Invalid command '" & Command & "'", ErrorCountIncrementIfNotInteractive
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
    gWriteErrorLine gCommands.BatchOrdersCommand.Name & " parameter must be either YES or NO or DEFAULT", ErrorCountIncrementNo
End Select
End Sub

Private Sub ProcessBracketCommand( _
                ByVal pParams As String, _
                ByVal pContractProcessor As ContractProcessor, _
                ByVal pID As String, _
                ByVal pModify As Boolean)
Const ProcName As String = "ProcessBracketCommand"
On Error GoTo Err

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pParams)
Dim lArg0 As String: lArg0 = lClp.Arg(0)

If lClp.NumberOfArgs = 3 Then
    ' the first arg is a contract spec
    Dim lContractSpec As IContractSpecifier
    Set lContractSpec = CreateContractSpecifierFromString(lArg0)

    If lContractSpec Is Nothing Then Exit Sub
    
    Set pContractProcessor = mCurrentGroup.AddContractProcessor(lContractSpec, _
                                        mBatchOrders, _
                                        mStageOrders, _
                                        OptionStrikeSelectionModeNone, _
                                        0, _
                                        OptionStrikeSelectionOperatorNone, _
                                        "")
    pParams = Right$(pParams, Len(pParams) - Len(lArg0))
End If

If Not pContractProcessor Is Nothing Then
    pContractProcessor.ProcessBracketCommand pParams, pID, pModify
Else
    gWriteErrorLine "No contract has been specified in this group", ErrorCountIncrementIfNotInteractive
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ProcessBuyCommand( _
                ByVal pParams As String, _
                ByVal pContractProcessor As ContractProcessor, _
                ByVal pID As String)
Const ProcName As String = "processBuyCommand"
On Error GoTo Err

processBuyOrSellCommand OrderActionBuy, pParams, pContractProcessor, pID, False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ProcessBuyAgainCommand( _
                ByVal pParams As String, _
                ByVal pContractProcessor As ContractProcessor, _
                ByVal pID As String)
Const ProcName As String = "ProcessBuyAgainCommand"
On Error GoTo Err

If pContractProcessor Is Nothing Then
    gWriteErrorLine "No Buy command to repeat", ErrorCountIncrementNo
ElseIf pContractProcessor.LatestBuyCommandParams = "" Then
    gWriteErrorLine "No Buy command to repeat", ErrorCountIncrementNo
Else
    ProcessBuyCommand pContractProcessor.LatestBuyCommandParams, pContractProcessor, pID
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processBuyOrSellCommand( _
                ByVal pAction As OrderActions, _
                ByVal pParams As String, _
                ByVal pContractProcessor As ContractProcessor, _
                ByVal pID As String, _
                ByVal pModify As Boolean)
Const ProcName As String = "processBuyOrSellCommand"
On Error GoTo Err

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pParams)
Dim lArg0 As String: lArg0 = lClp.Arg(0)

If Not IsNumeric(Left$(lArg0, 1)) Then
    ' the first arg is a contract spec
    Dim lContractSpec As IContractSpecifier
    Set lContractSpec = CreateContractSpecifierFromString(lArg0)

    If lContractSpec Is Nothing Then Exit Sub
    
    Set pContractProcessor = mCurrentGroup.AddContractProcessor(lContractSpec, _
                                        mBatchOrders, _
                                        mStageOrders, _
                                        OptionStrikeSelectionModeNone, _
                                        0, _
                                        OptionStrikeSelectionOperatorNone, _
                                        "")
    pParams = Right$(pParams, Len(pParams) - Len(lArg0))
End If

If Not pContractProcessor Is Nothing Then
    If pAction = OrderActionBuy Then
        pContractProcessor.ProcessBuyCommand pParams, pID, pModify
    Else
        pContractProcessor.ProcessSellCommand pParams, pID, pModify
    End If
Else
    gWriteErrorLine "No contract has been specified in this group", ErrorCountIncrementNo
End If

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description, ErrorCountIncrementNo
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Sub

Private Sub ProcessSellCommand( _
                ByVal pParams As String, _
                ByVal pContractProcessor As ContractProcessor, _
                ByVal pID As String)
Const ProcName As String = "processBuyCommand"
On Error GoTo Err

processBuyOrSellCommand OrderActionSell, pParams, pContractProcessor, pID, False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processSetBalanceCommand( _
                ByVal pParams As String)
Const ProcName As String = "processSetBalanceCommand"
On Error GoTo Err

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pParams)
If lClp.NumberOfArgs > 1 Then
    gWriteErrorLine "Too many parameters", ErrorCountIncrementNo
    Exit Sub
End If

Dim lAccountValueName As String: lAccountValueName = lClp.Arg(0)
If lAccountValueName = "" Then lAccountValueName = AccountValueNetLiquidation

If Not isValidAccountValueName(lAccountValueName) Then Exit Sub

gDefaultBalanceAccountValueName = lAccountValueName
Dim lBaseCurrency As String: lBaseCurrency = mAccountDataProvider.BaseCurrency
gWriteLineToConsole "Balance is " & _
                    mAccountDataProvider.GetAccountValue( _
                                        gDefaultBalanceAccountValueName, _
                                        lBaseCurrency).Value & _
                     " " & lBaseCurrency & _
                     " (" & gDefaultBalanceAccountValueName & ")"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processSetFundsCommand( _
                ByVal pParams As String)
Const ProcName As String = "processSetFundsCommand"
On Error GoTo Err

Dim lFunds As Double
If Not getFundsAmount(pParams, lFunds) Then Exit Sub

Dim lGroup As GroupResources
For Each lGroup In mGroups
    lGroup.FixedAccountBalance = lFunds
Next
gWriteLineToConsole "Funds for all groups set to " & CStr(CLng(lFunds))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processSetGroupFundsCommand( _
                ByVal pParams As String)
Const ProcName As String = "processSetGroupFundsCommand"
On Error GoTo Err

Dim lFunds As Double
If Not getFundsAmount(pParams, lFunds) Then Exit Sub

mCurrentGroup.FixedAccountBalance = lFunds
gWriteLineToConsole "Funds for group " & mCurrentGroup.GroupName & " set to " & CStr(CLng(lFunds))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processSetGroupRolloverCommand( _
                ByVal pParams As String)
Const ProcName As String = "processSetGroupRolloverCommand"
On Error GoTo Err

pParams = UCase$(pParams)

Dim ar() As String: ar = Split(pParams, " ")

pParams = Right$(pParams, Len(pParams) - Len(ar(0)))
If ar(0) = "OPTION" Then
    mCurrentGroup.SetDefaultRollover SecTypeOption, pParams
ElseIf ar(0) = "FUTURE" Then
    mCurrentGroup.SetDefaultRollover SecTypeFuture, pParams
Else
    gWriteErrorLine "First parameter must be 'OPTION' or 'FUTURE'", ErrorCountIncrementIfNotInteractive
End If

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description, ErrorCountIncrementIfNotInteractive
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Sub

Private Sub processSetRolloverCommand( _
                ByVal pParams As String)
Const ProcName As String = "processSetRolloverCommand"
On Error GoTo Err

pParams = UCase$(pParams)

Dim ar() As String: ar = Split(pParams, " ")

pParams = Right$(pParams, Len(pParams) - Len(ar(0)))

Dim lGroup As GroupResources
For Each lGroup In mGroups
    If ar(0) = "OPTION" Then
        lGroup.SetDefaultRollover SecTypeOption, pParams
    ElseIf ar(0) = "FUTURE" Then
        lGroup.SetDefaultRollover SecTypeFuture, pParams
    Else
        gWriteErrorLine "First parameter must be 'OPTION' or 'FUTURE'", ErrorCountIncrementIfNotInteractive
    End If
Next

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description, ErrorCountIncrementIfNotInteractive
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Sub

Private Sub ProcessSellAgainCommand( _
                ByVal pParams As String, _
                ByVal pContractProcessor As ContractProcessor, _
                ByVal pID As String)
Const ProcName As String = "ProcessSellAgainCommand"
On Error GoTo Err

If pContractProcessor Is Nothing Then
    gWriteErrorLine "No Sell command to repeat", ErrorCountIncrementNo
ElseIf pContractProcessor.LatestSellCommandParams = "" Then
    gWriteErrorLine "No Sell command to repeat", ErrorCountIncrementNo
Else
    ProcessSellCommand pContractProcessor.LatestSellCommandParams, pContractProcessor, pID
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processShowBalanceCommand()
Const ProcName As String = "processShowBalanceCommand"
On Error GoTo Err

Dim lBaseCurrency As String: lBaseCurrency = mAccountDataProvider.BaseCurrency
gWriteLineToConsole "Balance is " & _
                    mAccountDataProvider.GetAccountValue( _
                                        gDefaultBalanceAccountValueName, _
                                        lBaseCurrency).Value & _
                     " " & lBaseCurrency & _
                     " (" & gDefaultBalanceAccountValueName & ")"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processCancelCommand( _
                ByVal pParams As String)
Const ProcName As String = "processCancelCommand"
On Error GoTo Err

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pParams, " ")

Dim lKey As String
lKey = lClp.Arg(0)

If lKey = "" Then
    gWriteErrorLine "The order key is missing", ErrorCountIncrementNo
    Exit Sub
End If

Dim lEntry As LiveOrderEntry
If IsNumeric(lKey) Then
    If gLiveOrders.TryItemAtIndex(lKey, lEntry) Then
        lKey = lEntry.Key
    End If
Else
    gLiveOrders.TryItem lKey, lEntry
End If

If lEntry Is Nothing Then
    gWriteErrorLine "Order " & lKey & " cannot be cancelled", ErrorCountIncrementNo
    Exit Sub
End If

If lEntry.Cancelled Then
    gWriteErrorLine "Order " & lKey & " has already been cancelled", ErrorCountIncrementNo
    Exit Sub
End If

lEntry.Cancelled = True

Dim lBracketOrder As IBracketOrder
Set lBracketOrder = lEntry.Order

If lBracketOrder Is Nothing Then
    Dim lOrderPlacer As OrderPlacer
    Set lOrderPlacer = mGroups.Item(lEntry.GroupName).OrderPlacers.Item(lKey)
    lOrderPlacer.Cancel "By user"
End If

Dim lEvenIfFilled As Boolean
If lClp.Arg(1) = "" Then
ElseIf lClp.Arg(1) = "!" Then
    lEvenIfFilled = True
Else
    gWriteErrorLine lClp.Arg(1) & " second parameter must be blank or '!' (cancel even if filled)", ErrorCountIncrementNo
End If

If Not lBracketOrder Is Nothing Then
    gWriteLineToConsole "Cancelling " & lKey
    lBracketOrder.Cancel lEvenIfFilled
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
        gWriteErrorLine "Second argument must be either 'MKT' or 'LMT'", ErrorCountIncrementNo
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
    gWriteErrorLine "Too many arguments", ErrorCountIncrementNo
End If

If lGroupName = "" Then
    lGroupName = mCurrentGroup.GroupName
ElseIf UCase$(lGroupName) = AllGroups Then
    lGroupName = AllGroups
ElseIf lGroupName = DefaultGroupName Then
ElseIf Not isGroupValid(lGroupName) Then
    lError = True
    gWriteErrorLine lGroupName & " is not a valid group name", ErrorCountIncrementNo
ElseIf Not mGroups.Contains(lGroupName) Then
    lError = True
    gWriteErrorLine "No such group", True
End If

If lCloseoutMode = CloseoutModeMarket Then
    If lPriceStr <> "" Then
        lError = True
        gWriteErrorLine "A price specification cannot be supplied for closeout at market", ErrorCountIncrementNo
    End If
ElseIf lPriceStr = "" And lPriceSpec Is Nothing Then
    lError = True
    gWriteErrorLine "Closeout price specification is missing", ErrorCountIncrementNo
ElseIf Not lPriceSpec Is Nothing Then
ElseIf Not ParsePriceAndOffset(lPriceSpec, lPriceStr, SecTypeNone, 0#, True) Then
    lError = True
    gWriteErrorLine "Invalid price specification", ErrorCountIncrementNo
End If

If lError Then Exit Sub

Dim lCloseoutProcessor As New CloseoutProcessor
lCloseoutProcessor.Initialise mOrderManager, mGroups, lCloseoutMode, lPriceSpec

If lGroupName = AllGroups Then
    lCloseoutProcessor.CloseoutAll
ElseIf Not mGroups.Contains(lGroupName) Then
    gWriteErrorLine "No such group", ErrorCountIncrementNo
Else
    lCloseoutProcessor.CloseoutGroup lGroupName
End If

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description, ErrorCountIncrementNo
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
If gCommands.HelpCommand.Parse(pParams) Or gCommands.Help1Command.Parse(pParams) Then
    showContractHelp
    Exit Sub
End If

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pParams, " ")

Dim lSpecString As String
lSpecString = lClp.Arg(0)

Dim lContractSpec As IContractSpecifier
Dim lStrikeSelectionMode As OptionStrikeSelectionModes
Dim lParameter As Long
Dim lOperator As OptionStrikeSelectionOperators
Dim lUnderlyingExchange As String

If lSpecString <> "" Then
    Set lContractSpec = CreateContractSpecifierFromString(lSpecString)
Else
    Set lContractSpec = parseContractSpec(lClp, _
                                        lStrikeSelectionMode, _
                                        lParameter, _
                                        lOperator, _
                                        lUnderlyingExchange)
End If

If lContractSpec Is Nothing Then
    mCurrentGroup.ClearCurrentContractProcessor
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
                                    lStrikeSelectionMode, _
                                    lParameter, _
                                    lOperator, _
                                    lUnderlyingExchange

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description, ErrorCountIncrementNo
    Resume Next
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Sub

Private Sub processEndOrdersCommand()
Const ProcName As String = "processEndOrdersCommand"
On Error GoTo Err

If mBlockingErrorCount <> 0 Then
    gWriteLineToConsole gErrorCount & " errors have been found - no orders will be placed"
    mBlockingErrorCount = 0
    mErrorCount = 0
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation
    Exit Sub
End If

If getNumberOfUnprocessedOrders = 0 Then
    gWriteLineToConsole "No orders have been defined"
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
    gWriteErrorLine "Invalid group name: you can't create a group called '" & lGroupName & "'", ErrorCountIncrementNo
    Exit Sub
ElseIf Not isGroupValid(lGroupName) Then
    gWriteErrorLine "Invalid group name: first character must be letter or digit; remaining characters must be letter, digit, hyphen or underscore", ErrorCountIncrementNo
    Exit Sub
End If

If Not mGroups.TryItem(lGroupName, mCurrentGroup) Then
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
ElseIf UCase$(Left$(pParams, Len(OrdersSubcommand))) = OrdersSubcommand Then
    listOrders Trim$(Right$(pParams, Len(pParams) - Len(OrdersSubcommand)))
ElseIf UCase$(pParams) = PositionsSubcommand Then
    listPositions
ElseIf UCase$(pParams) = TradesSubcommand Then
    listTrades
Else
    gWriteErrorLine gCommands.ListCommand.Name & " parameter must be one of " & GroupsSubcommand & ", " & PositionsSubcommand & " or " & TradesSubcommand, ErrorCountIncrementNo
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processModifyCommand(ByVal pParams As String)
Const ProcName As String = "processModifyCommand"
On Error GoTo Err

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pParams, " ")

Dim lKey As String
lKey = lClp.Arg(0)

If lKey = "" Then
    gWriteErrorLine "The order key is missing", ErrorCountIncrementNo
    Exit Sub
End If

Dim lCommandIndex As Long
lCommandIndex = InStr(1, pParams, lKey) + Len(lKey)

Dim lEntry As LiveOrderEntry
If IsNumeric(lKey) Then
    If gLiveOrders.TryItemAtIndex(lKey, lEntry) Then
        lKey = lEntry.Key
    End If
Else
    gLiveOrders.TryItem lKey, lEntry
End If

If lEntry Is Nothing Then
    gWriteErrorLine "Order " & lKey & " cannot be modified", ErrorCountIncrementNo
    Exit Sub
End If

processGroupCommand lEntry.GroupName, mCurrentGroup.CurrentContractProcessor

Dim lCommand As String
lCommand = Trim$(Right$(pParams, Len(pParams) - lCommandIndex))

Set lClp = CreateCommandLineParser(lCommand, " ")

Dim lVerb As String: lVerb = UCase$(lClp.Arg(0))
Dim lArg1 As String: lArg1 = lClp.Arg(1)

Dim lContractSpec As IContractSpecifier
Dim lFirstArgIsContractSpec As Boolean
If (lVerb = gCommands.BuyCommand.Name Or _
        lVerb = gCommands.SellCommand.Name) And _
    Not IsNumeric(Left$(lArg1, 1)) _
Then
    lFirstArgIsContractSpec = True
ElseIf lVerb = gCommands.BracketCommand.Name And _
    lClp.NumberOfArgs = 4 _
Then
    lFirstArgIsContractSpec = True
End If
    
If lFirstArgIsContractSpec Then
    Set lContractSpec = CreateContractSpecifierFromString(lArg1)
    pParams = Trim$(Right$(lCommand, Len(lCommand) - InStr(Len(lVerb), lCommand, lArg1) - Len(lArg1)))
Else
    pParams = Trim$(Right$(lCommand, Len(lCommand) - Len(lVerb)))
End If

If Not lContractSpec Is Nothing Then
ElseIf lEntry.Order Is Nothing Then
    Set lContractSpec = lEntry.BracketOrderSpec.Contract.Specifier
Else
    Set lContractSpec = lEntry.Order.Contract.Specifier
End If

Dim lContractProcessor As ContractProcessor
Set lContractProcessor = mCurrentGroup.AddContractProcessor(lContractSpec, _
                                    mBatchOrders, _
                                    mStageOrders, _
                                    OptionStrikeSelectionModeNone, _
                                    0, _
                                    OptionStrikeSelectionOperatorNone, _
                                    "")
Select Case UCase$(lVerb)
Case gCommands.BracketCommand.Name
    lContractProcessor.ProcessBracketCommand pParams, lKey, True
Case gCommands.BuyCommand.Name
    processBuyOrSellCommand OrderActionBuy, pParams, lContractProcessor, lKey, True
Case gCommands.SellCommand.Name
    processBuyOrSellCommand OrderActionSell, pParams, lContractProcessor, lKey, True
Case Else
    gWriteErrorLine "Invalid command '" & lVerb & "'", ErrorCountIncrementNo
End Select

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
        gWriteErrorLine "No such group", ErrorCountIncrementNo
        Exit Sub
    End If
    gWriteLineToConsole "Purging " & lRes.GroupName
    lRes.Purge
    mGroups.Remove pParams
    Set mCurrentGroup = mGroups.Add(DefaultGroupName)
Else
    gWriteErrorLine "Invalid group name", ErrorCountIncrementNo
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
        Dim lStrikeSelectionMode As OptionStrikeSelectionModes
        Dim lParameter As Long
        Dim lOperator As OptionStrikeSelectionOperators
        Dim lUnderlyingExchangeName As String
        Set lContractSpec = parseContractSpec(lClp, _
                                            lStrikeSelectionMode, _
                                            lParameter, _
                                            lOperator, _
                                            lUnderlyingExchangeName)
    End If
    If lContractSpec Is Nothing Then Exit Sub
    
    Dim lQuoteFetcher As New QuoteFetcher
    lQuoteFetcher.FetchQuote lContractSpec, mContractStore, mMarketDataManager
End If

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description, ErrorCountIncrementNo
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Sub

Private Sub processResetCommand()
mStageOrders = mStageOrdersDefault
mBatchOrders = mBatchOrdersDefault
mBlockingErrorCount = 0
mErrorCount = 0
gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation
End Sub

Private Sub processStageOrdersCommand( _
                ByVal pParams As String)
Select Case UCase$(pParams)
Case Default
    mStageOrders = mStageOrdersDefault
Case Yes
    If Not (mOrderSubmitterFactory.Capabilities And OrderSubmitterCapabilityCanStageOrders) = OrderSubmitterCapabilityCanStageOrders Then
        gWriteErrorLine gCommands.StageOrdersCommand.Name & " parameter cannot be YES with current configuration", ErrorCountIncrementNo
        Exit Sub
    End If
    mStageOrders = True
Case No
    mStageOrders = False
Case Else
    gWriteErrorLine gCommands.StageOrdersCommand.Name & " parameter must be either YES or NO or DEFAULT", ErrorCountIncrementNo
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
    gWriteErrorLine "The /" & BatchOrdersSwitch & " switch has an invalid value: it must be either YES or NO (not case-sensitive)", ErrorCountIncrementNo
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
    gWriteErrorLine "The /" & MonitorSwitch & " switch has an invalid value: it must be either YES or NO (not case-sensitive, default is YES)", ErrorCountIncrementNo
    setMonitor = False
End If
End Function

Private Sub setOrderRecovery()
Const ProcName As String = "setOrderRecovery"
On Error GoTo Err

If mScopeName = "" Then Exit Sub

mRecoveryFileDir = ApplicationSettingsFolder
If mClp.SwitchValue(RecoveryFileDirSwitch) <> "" Then mRecoveryFileDir = mClp.SwitchValue(RecoveryFileDirSwitch)

If mOrderPersistenceDataStore Is Nothing Then Set mOrderPersistenceDataStore = CreateOrderPersistenceDataStore(mRecoveryFileDir)

Dim lOrderRecoverer As New OrderRecoverer
lOrderRecoverer.RecoverOrders mOrderManager, _
                            mScopeName, _
                            mOrderPersistenceDataStore, _
                            mOrderRecoveryAgent, _
                            mMarketDataManager, _
                            mOrderSubmitterFactory, _
                            mGroups, _
                            mMoneyManager, _
                            mAccountDataProvider, _
                            mCurrencyConverter

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
        gWriteErrorLine "The /" & StageOrdersSwitch & " switch has an invalid value: it cannot be YES with the current configuration", ErrorCountIncrementNo
        setStageOrders = False
        Exit Function
    End If
    mStageOrdersDefault = True
ElseIf UCase$(mClp.SwitchValue(StageOrdersSwitch)) = No Then
    mStageOrdersDefault = False
Else
    gWriteErrorLine "The /" & StageOrdersSwitch & " switch has an invalid value: it must be either YES or NO (not case-sensitive)", ErrorCountIncrementNo
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
                gCommands.CancelCommand, _
                gCommands.EndOrdersCommand, _
                gCommands.ModifyCommand, _
                gCommands.Modify1Command, _
                gCommands.Modify2Command, _
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
                gCommands.GroupCommand, _
                gCommands.SetBalanceCommand, _
                gCommands.SetFundsCommand, _
                gCommands.SetGroupFundsCommand, _
                gCommands.SetGroupRolloverCommand, _
                gCommands.SetRolloverCommand, _
                gCommands.ShowBalanceCommand
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
GetLogger("tradebuild.log.orderutils.contractresolution").AddLogListener lLogfile
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
                        "-BracketOrders" & _
                        ".log", _
                    includeTimestamp:=False, _
                    includeLogLevel:=False)
    GetLogger("position.bracketorderprofilestring").AddLogListener lLogfile

    Set lLogfile = CreateFileLogListener( _
                    lResultsPath & "Orders\" & _
                        ProjectName & _
                        "(" & mClientId & ")" & _
                        "-" & lFilenameSuffix & _
                        "-Rollovers" & _
                        ".log", _
                    includeTimestamp:=False, _
                    includeLogLevel:=False)
    GetLogger("position.bracketorderrollover").AddLogListener lLogfile
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
                ByVal pLogApiMessageStats As Boolean, _
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
    gWriteErrorLine "port must be an integer > 1024 and <= 65535", ErrorCountIncrementNo
    setupTwsApi = False
End If
    
Dim clientId As String
clientId = lClp.Arg(2)
If clientId = "" Then
    clientId = DefaultClientId
ElseIf Not IsInteger(clientId, 0, 999999999) Then
    gWriteErrorLine "clientId must be an integer >= 0 and <= 999999999", ErrorCountIncrementNo
    setupTwsApi = False
End If
pClientId = CLng(clientId)

Dim connectionRetryInterval As String
connectionRetryInterval = lClp.Arg(3)
If connectionRetryInterval = "" Then
ElseIf Not IsInteger(connectionRetryInterval, 0, 3600) Then
    gWriteErrorLine "Error: connection retry interval must be an integer >= 0 and <= 3600", ErrorCountIncrementNo
    setupTwsApi = False
End If

If Not setupTwsApi Then Exit Function

Dim lListener As New TwsConnectionListener

If connectionRetryInterval = "" Then
    Set mTwsClient = GetClient(server, _
                            CLng(port), _
                            pClientId, _
                            pLogApiMessages:=pLogApiMessages, _
                            pLogRawApiMessages:=pLogRawApiMessages, _
                            pLogApiMessageStats:=pLogApiMessageStats, _
                            pConnectionStateListener:=lListener)
Else
    Set mTwsClient = GetClient(server, _
                            CLng(port), _
                            pClientId, _
                            pConnectionRetryIntervalSecs:=CLng(connectionRetryInterval), _
                            pLogApiMessages:=pLogApiMessages, _
                            pLogRawApiMessages:=pLogRawApiMessages, _
                            pLogApiMessageStats:=pLogApiMessageStats, _
                            pConnectionStateListener:=lListener)
End If

Set mContractStore = mTwsClient.GetContractStore
Set mMarketDataManager = CreateRealtimeDataManager(mTwsClient.GetMarketDataFactory, mTwsClient.GetContractStore)
Set mAccountDataProvider = mTwsClient.GetAccountDataProvider
mAccountDataProvider.Load True
Set mCurrencyConverter = CreateCurrencyConverter(mMarketDataManager, mContractStore)

mMarketDataManager.LoadFromConfig gGetMarketDataSourcesConfig(mConfigStore)

Set mOrderRecoveryAgent = mTwsClient

If pSimulateOrders Then
    Set mOrderSubmitterFactory = New SimOrderSubmitterFactory
Else
    Set mOrderSubmitterFactory = mTwsClient
End If
    
Set mOrderManager = New OrderManager
mOrderManager.ContractStorePrimary = mContractStore
mOrderManager.MarketDataManager = mMarketDataManager
mOrderManager.OrderSubmitterFactory = mOrderSubmitterFactory

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

s = Mid$(pApiMessageLogging, 2, 1)
If s = None Then
    pLogRawApiMessages = ApiMessageLoggingOptionNone
ElseIf s = Default Then
    pLogRawApiMessages = ApiMessageLoggingOptionDefault
ElseIf s = Always Then
    pLogRawApiMessages = ApiMessageLoggingOptionAlways
Else
    Exit Function
End If

s = Mid$(pApiMessageLogging, 3, 1)
If s = No Then
    pLogApiMessageStats = False
ElseIf s = Yes Then
    pLogApiMessageStats = True
Else
    Exit Function
End If

validateApiMessageLogging = True
End Function

