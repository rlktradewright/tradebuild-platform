VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IErrorListener

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

Private Const ModuleName                            As String = "ContractProcessor"

Private Const OrderPendingMessage                   As String = "Order is pending contract resolution"

'@================================================================================
' Member variables
'@================================================================================

Private mIdent                                         As String
Private mName                                       As String

Private mInitialContractSpec                        As IContractSpecifier
Private mContract                                   As IContract
Private mDataSource                                 As IMarketDataSource

Private mInitialStrikeSelectionMode                 As OptionStrikeSelectionModes
Private mInitialStrikeParameter                     As Long
Private mInitialStrikeOperator                      As OptionStrikeSelectionOperators

Private mBracketOrders                              As New EnumerableCollection

Private mBOSpec                                     As BracketOrderSpecification

Private mPositionManager                            As PositionManager
Private mAccountDataProvider                        As IAccountDataProvider
Private mCurrencyConverter                          As ICurrencyConverter

Private mFixedAccountBalance                        As Double

Private WithEvents mOrderContext                    As OrderContext
Attribute mOrderContext.VB_VarHelpID = -1

Private mContractResolver                           As ContractResolver
Private mGroupResources                             As GroupResources

Private mGroupName                                  As String

Private mLatestBuyCommandParams                     As String
Private mLatestSellCommandParams                    As String

Private mSelectionMode                              As OptionStrikeSelectionModes
Private mParameter                                  As Long
Private mOperator                                   As OptionStrikeSelectionOperators
Private mUnderlyingExchangeName                     As String

Private mPendingCommands                            As SortedDictionary
Private mStageOrders                                As Boolean

Private mDefaultRolloverSpec                        As RolloverSpecification

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mPendingCommands = CreateSortedDictionary(KeyTypeDate)
End Sub

'@================================================================================
' IErrorListener Interface Members
'@================================================================================

Private Sub IErrorListener_Notify(ev As ErrorEventData)
Const ProcName As String = "IErrorListener_Notify"
On Error GoTo Err

gWriteErrorLine "Datasource error for " & gGetContractName(mContract) & _
                ": (" & ev.ErrorCode & ") " & ev.ErrorMessage
mGroupResources.NotifyError Me
Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mOrderContext Event Handlers
'@================================================================================

Private Sub mOrderContext_Change(ev As ChangeEventData)
Const ProcName As String = "mOrderContext_Change"
On Error GoTo Err

Dim lChangeType As OrderContextChangeTypes: lChangeType = ev.ChangeType

If ev.ChangeType <> OrderContextChangeTypes.OrderContextReadyStateChanged Then Exit Sub

If mOrderContext.IsReady Then processPendingCommands
    
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get BracketOrders() As EnumerableCollection
Set BracketOrders = mBracketOrders
End Property

Friend Property Get Contract() As IContract
Set Contract = mContract
End Property

Friend Property Get ContractName() As String
If Not mContract Is Nothing Then
    ContractName = gGetContractName(mContract)
Else
    ContractName = mContractResolver.ContractName
End If
End Property

Friend Property Get DataSource() As IMarketDataSource
Set DataSource = mDataSource
End Property

Friend Property Get GroupName() As String
GroupName = mGroupName
End Property

Friend Property Let FixedAccountBalance(ByVal Value As Double)
mFixedAccountBalance = Value
If Not mPositionManager Is Nothing Then mPositionManager.FixedAccountBalance = mFixedAccountBalance
End Property

Friend Property Get ID() As String
ID = mIdent
End Property

Friend Property Get InitialContractSpec() As IContractSpecifier
Set InitialContractSpec = mInitialContractSpec
End Property

Friend Property Get IsReady()
IsReady = mOrderContext.IsReady
End Property

Friend Property Get LatestBuyCommandParams() As String
LatestBuyCommandParams = mLatestBuyCommandParams
End Property

Friend Property Get LatestSellCommandParams() As String
LatestSellCommandParams = mLatestSellCommandParams
End Property

Friend Property Get Name() As String
Name = mName
End Property

Friend Property Get OrderContext() As OrderContext
Set OrderContext = mOrderContext
End Property

Friend Property Get PendingCommandsEnumerator() As Enumerator
Set PendingCommandsEnumerator = mPendingCommands.EntryEnumerator
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Activate( _
                ByVal pName As String, _
                ByVal pcontract As IContract, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pOrderManager As OrderManager, _
                ByVal pScopeName As String, _
                ByVal pGroupName As String, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory, _
                ByVal pMoneyManager As MoneyManager, _
                ByVal pGroupResources As GroupResources)
Const ProcName As String = "Activate"
On Error GoTo Err

Dim failpoint As String

failpoint = "010"
mName = pName
Set mGroupResources = pGroupResources

failpoint = "020"
Set mContract = pcontract
mGroupName = pGroupName

failpoint = "030"
mSelectionMode = mContractResolver.StrikeSelectionMode
failpoint = "040"
mOperator = mContractResolver.StrikeSelectionOperator
failpoint = "050"
mParameter = mContractResolver.StrikeSelectionParameter

failpoint = "060"
mUnderlyingExchangeName = mContractResolver.UnderlyingExchangeName

failpoint = "070"
Set mDataSource = pMarketDataManager.CreateMarketDataSource(CreateFuture(pcontract), False)
failpoint = "080"
mDataSource.AddErrorListener Me
mDataSource.StartMarketData

If pScopeName = "" Then
failpoint = "090"
    Set mPositionManager = pOrderManager.CreatePositionManager( _
                                pName, _
                                mDataSource, _
                                pOrderSubmitterFactory, _
                                pGroupName, _
                                pMoneyManager:=pMoneyManager, _
                                pAccountDataProvider:=mAccountDataProvider, _
                                pCurrencyConverter:=mCurrencyConverter)
Else
failpoint = "100"
    Set mPositionManager = pOrderManager.CreateRecoverablePositionManager( _
                                pName, _
                                mDataSource, _
                                pOrderSubmitterFactory, _
                                pScopeName, _
                                mGroupName, _
                                pMoneyManager:=pMoneyManager, _
                                pAccountDataProvider:=mAccountDataProvider, _
                                pCurrencyConverter:=mCurrencyConverter)
End If
failpoint = "110"
If mFixedAccountBalance <> 0 Then mPositionManager.FixedAccountBalance = mFixedAccountBalance
failpoint = "120"
Set mOrderContext = mPositionManager.OrderContexts.DefaultOrderContext

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, failpoint
End Sub

Friend Sub AddBracketOrderSpecs( _
                ByVal pContractProcessor As ContractProcessor)
Const ProcName As String = "AddBracketOrderSpecs"
On Error GoTo Err

Dim lBOS As BracketOrderSpecification
For Each lBOS In pContractProcessor.BracketOrders
    mBracketOrders.Add lBOS
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddPendingCommands( _
                ByVal pContractProcessor As ContractProcessor)
Const ProcName As String = "AddPendingCommands"
On Error GoTo Err

Dim en As Enumerator
Set en = pContractProcessor.PendingCommandsEnumerator

Do While en.MoveNext
    Dim lEntry As DictionaryEntry: Set lEntry = en.Current
    Dim lTimestamp As Date: lTimestamp = lEntry.Key
    Dim lPendingCommand As PendingCommand: Set lPendingCommand = lEntry.Data
    addPendingCommand lTimestamp, _
                        lPendingCommand.Command, _
                        lPendingCommand.Params, _
                        lPendingCommand.OrderID
Loop

If IsReady Then processPendingCommands

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ClearOrders()
mBracketOrders.Clear
End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Set mGroupResources = Nothing

If Not mContractResolver Is Nothing Then
    mContractResolver.Finish
    Set mContractResolver = Nothing
End If

mBracketOrders.Clear
If Not mPositionManager Is Nothing Then
    mPositionManager.Finish
    If Not mDataSource Is Nothing Then mDataSource.Finish
End If

mPendingCommands.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pInitialContractSpec As IContractSpecifier, _
                ByVal pInitialStrikeSelectionMode As OptionStrikeSelectionModes, _
                ByVal pInitialStrikeParameter As Long, _
                ByVal pInitialStrikeOperator As OptionStrikeSelectionOperators, _
                ByVal pID As String, _
                ByVal pContractResolver As ContractResolver, _
                ByVal pAccountDataProvider As IAccountDataProvider, _
                ByVal pCurrencyConverter As ICurrencyConverter, _
                ByVal pStageOrders As Boolean)
Set mInitialContractSpec = pInitialContractSpec
mInitialStrikeSelectionMode = pInitialStrikeSelectionMode
mInitialStrikeParameter = pInitialStrikeParameter
mInitialStrikeOperator = pInitialStrikeOperator
mIdent = pID
Set mContractResolver = pContractResolver
Set mAccountDataProvider = pAccountDataProvider
Set mCurrencyConverter = pCurrencyConverter
mStageOrders = pStageOrders
End Sub

Friend Sub ProcessBracketCommand( _
                ByVal pParams As String, _
                ByVal pID As String)
Const ProcName As String = "ProcessBracketCommand"
On Error GoTo Err

If mOrderContext Is Nothing Then
    addPendingCommand GetTimestamp, gCommands.BracketCommand, pParams, pID
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification
    Exit Sub
End If

Set mBOSpec = New BracketOrderSpecification
Set mBOSpec.Contract = mContract
mBOSpec.ID = pID
mBOSpec.EntryTIF = OrderTIFDay
mBOSpec.StopLossTIF = OrderTIFGoodTillCancelled
mBOSpec.TargetTIF = OrderTIFGoodTillCancelled

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pParams, " ")

If lClp.NumberOfArgs <> 2 Then gWriteErrorLine "There must be only two arguments, action (BUY or SELL) and size. Additionally there can be several attributes beginnig with ""/"""

mBOSpec.Action = getOrderAction(lClp.Arg(0))
If mBOSpec.Action = OrderActionNone Then gWriteErrorLine "Order action must be BUY or SELL"

If Not setQuantity(lClp.Arg(1)) Then
    gWriteErrorLine "Invalid quantity", True
End If

Dim lSwitchName As String
lSwitchName = CancelAfterSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.CancelAfter = CInt(lClp.SwitchValue(lSwitchName))
If mBOSpec.CancelAfter < 0 Then gWriteErrorLine "CancelAfter must be an integer >= 0"

lSwitchName = CancelPriceSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.CancelPrice = CDbl(lClp.SwitchValue(lSwitchName))

lSwitchName = DescriptionSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.Description = lClp.SwitchValue(lSwitchName)

lSwitchName = GoodAfterTimeSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.GoodAfterTime = CDate(lClp.SwitchValue(lSwitchName))

lSwitchName = GoodTillDateSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.GoodTillDate = CDate(lClp.SwitchValue(lSwitchName))

lSwitchName = TimezoneSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.TimezoneName = lClp.SwitchValue(lSwitchName)

gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrTypeMismatch Then
    gWriteErrorLine lSwitchName & ": invalid value"
    Resume Next
ElseIf Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description
    Resume Next
End If
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function ProcessBuyCommand( _
                ByVal pParams As String, _
                ByVal pID As String) As Boolean
Const ProcName As String = "ProcessBuyCommand"
On Error GoTo Err

If mOrderContext Is Nothing Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation
    addPendingCommand GetTimestamp, _
                            gCommands.BuyCommand, _
                            pParams, _
                            pID, _
                            OrderPendingMessage
Else
    ProcessBuyCommand = processBuyOrSellCommand(OrderActionBuy, pParams, pID)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function ProcessSellCommand( _
                ByVal pParams As String, _
                ByVal pID As String) As Boolean
Const ProcName As String = "ProcessSellCommand"
On Error GoTo Err

If mOrderContext Is Nothing Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation
    addPendingCommand GetTimestamp, _
                            gCommands.SellCommand, _
                            pParams, _
                            pID, _
                            OrderPendingMessage
Else
    ProcessSellCommand = processBuyOrSellCommand(OrderActionSell, pParams, pID)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ProcessEndBracketCommand()
Const ProcName As String = "ProcessEndBracketCommand"
On Error GoTo Err

If mOrderContext Is Nothing Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation
    addPendingCommand GetTimestamp, _
                            gCommands.EndBracketCommand, _
                            "", _
                            "", _
                            OrderPendingMessage
    Exit Sub
End If

If Not mBOSpec.RolloverSpec Is Nothing Then
    mBOSpec.RolloverSpec.setStopLossOrderType(mBOSpec.StopLossType). _
                    setStopLossLimitPriceSpec(mBOSpec.StopLossLimitPriceSpec). _
                    setStopLossTriggerPriceSpec(mBOSpec.StopLossTriggerPriceSpec). _
                    setTargetOrderType(mBOSpec.TargetType). _
                    setTargetLimitPriceSpec(mBOSpec.TargetLimitPriceSpec). _
                    setTargetTriggerPriceSpec mBOSpec.TargetTriggerPriceSpec
End If

If gErrorCount = 0 Then mBracketOrders.Add mBOSpec

gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ProcessEntryCommand( _
                ByVal pParams As String)
Const ProcName As String = "ProcessEntryCommand"
On Error GoTo Err

If mOrderContext Is Nothing Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification, gCommandListOrderCompletion
    addPendingCommand GetTimestamp, gCommands.EntryCommand, pParams
    Exit Sub
End If

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pParams, " ")

If lClp.NumberOfArgs <> 1 Then gWriteErrorLine "There must be only one argument, the entry order type. Additionally there can be several attributes beginnig with ""/"""

Dim lOrderTypeStr As String: lOrderTypeStr = lClp.Arg(0)
If Not parsePseudoOrder(lOrderTypeStr, lClp.SwitchValue(OffsetSwitch)) Then
    mBOSpec.EntryType = getEntryOrderType(lOrderTypeStr)

    If Not IsEntryOrderType(mBOSpec.EntryType) Then
        mBOSpec.EntryType = OrderTypeNone
        gWriteErrorLine "Invalid entry order type", True
    Else
        Dim lPriceStr As String: lPriceStr = getPriceString(lClp)
        Dim lTriggerPriceStr As String: lTriggerPriceStr = getTriggerPriceString(lClp)
        
        Dim lLimitPriceSpec As PriceSpecifier
        Dim lTriggerPriceSpec As PriceSpecifier
        Dim lMessage As String
        If Not gParseOrderPrices(IIf(lPriceStr <> "", lPriceStr, lTriggerPriceStr), _
                        IIf(lPriceStr <> "", lTriggerPriceStr, ""), _
                        mBOSpec.EntryType, _
                        lLimitPriceSpec, _
                        lTriggerPriceSpec, _
                        lMessage) _
        Then
            gWriteErrorLine lMessage
        End If
        Set mBOSpec.EntryLimitPriceSpec = lLimitPriceSpec
        Set mBOSpec.EntryTriggerPriceSpec = lTriggerPriceSpec
        
        If mBOSpec.EntryLimitPriceSpec.PriceType = PriceValueTypeEntry Or _
            mBOSpec.EntryTriggerPriceSpec.PriceType = PriceValueTypeEntry Then
            gWriteErrorLine "Invalid use of 'ENTRY'", True
        End If
    End If
End If

Dim lSwitchName  As String
lSwitchName = TIFSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.EntryTIF = OrderTIFFromString(lClp.SwitchValue(lSwitchName))

lSwitchName = ReasonSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.EntryReason = lClp.SwitchValue(lSwitchName)

lSwitchName = IgnoreRTHSwitch
mBOSpec.EntryIgnoreRTH = lClp.Switch(lSwitchName)

If mBOSpec.EntryType = OrderTypeNone And gErrorCount = 0 Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification
Else
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification, gCommandListOrderCompletion
End If

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrTypeMismatch Then
    gWriteErrorLine lSwitchName & ": invalid value"
    Resume Next
ElseIf Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description
    Resume Next
End If
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ProcessRolloverCommand( _
                ByVal pParams As String)
Const ProcName As String = "ProcessRolloverCommand"
On Error GoTo Err

If mOrderContext Is Nothing Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification, gCommandListOrderCompletion
    addPendingCommand GetTimestamp, gCommands.RolloverCommand, pParams
    Exit Sub
End If

If pParams <> "" Then
    Set mBOSpec.RolloverSpec = gProcessRolloverCommand( _
                                    pParams, _
                                    mContract.Specifier.SecType)
ElseIf Not mDefaultRolloverSpec Is Nothing Then
    Set mBOSpec.RolloverSpec = mDefaultRolloverSpec
Else
    gWriteErrorLine "No default rollover specification has been defined"
End If

If mBOSpec.EntryType = OrderTypeNone And gErrorCount = 0 Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification
Else
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification, gCommandListOrderCompletion
End If

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Sub

Friend Sub ProcessQuitCommand()
Const ProcName As String = "ProcessQuitCommand"
On Error GoTo Err

If mOrderContext Is Nothing Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation
    addPendingCommand GetTimestamp, _
                            gCommands.QuitCommand, _
                            "", _
                            "", _
                            "Bracket order definition terminated"
    Exit Sub
End If

Set mBOSpec = Nothing
gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation
gWriteLineToConsole "Command definition terminated"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ProcessStopLossCommand( _
                ByVal pParams As String)
Const ProcName As String = "ProcessStopLossCommand"
On Error GoTo Err

If mOrderContext Is Nothing Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification, gCommandListOrderCompletion
    addPendingCommand GetTimestamp, gCommands.StopLossCommand, pParams
    Exit Sub
End If

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pParams, " ")

If lClp.NumberOfArgs <> 1 Then gWriteErrorLine "There must be only one argument, the stop-loss order type. Additionally there can be several attributes beginnig with ""/"""

Dim lOrderTypeStr As String: lOrderTypeStr = lClp.Arg(0)
If Not parsePseudoOrder(lOrderTypeStr, lClp.SwitchValue(OffsetSwitch)) Then
    mBOSpec.StopLossType = getStopLossOrderType(lClp.Arg(0))

    If Not IsStopLossOrderType(mBOSpec.StopLossType) Then
        mBOSpec.StopLossType = OrderTypeNone
        gWriteErrorLine "Invalid stop-loss order type"
    Else
        Dim lPriceStr As String: lPriceStr = getPriceString(lClp)
        Dim lTriggerPriceStr As String: lTriggerPriceStr = getTriggerPriceString(lClp)
        
        Dim lLimitPriceSpec As PriceSpecifier
        Dim lTriggerPriceSpec As PriceSpecifier
        Dim lMessage As String
        If Not gParseOrderPrices(IIf(lPriceStr <> "", lPriceStr, lTriggerPriceStr), _
                        IIf(lPriceStr <> "", lTriggerPriceStr, ""), _
                        mBOSpec.StopLossType, _
                        lLimitPriceSpec, _
                        lTriggerPriceSpec, _
                        lMessage) _
        Then
            gWriteErrorLine lMessage
        End If
        Set mBOSpec.StopLossLimitPriceSpec = lLimitPriceSpec
        Set mBOSpec.StopLossTriggerPriceSpec = lTriggerPriceSpec
    End If
End If

Dim lSwitchName As String
lSwitchName = TIFSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.StopLossTIF = OrderTIFFromString(lClp.SwitchValue(lSwitchName))

lSwitchName = ReasonSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.StopLossReason = lClp.SwitchValue(lSwitchName)

lSwitchName = IgnoreRTHSwitch
mBOSpec.StopLossIgnoreRTH = lClp.Switch(lSwitchName)

If mBOSpec.EntryType = OrderTypeNone And gErrorCount = 0 Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification
Else
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification, gCommandListOrderCompletion
End If

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrTypeMismatch Then
    gWriteErrorLine lSwitchName & ": invalid value"
    Resume Next
ElseIf Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description
    Resume Next
End If
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ProcessTargetCommand( _
                ByVal pParams As String)
Const ProcName As String = "ProcessTargetCommand"
On Error GoTo Err

If mOrderContext Is Nothing Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification, gCommandListOrderCompletion
    addPendingCommand GetTimestamp, gCommands.TargetCommand, pParams
    Exit Sub
End If

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pParams, " ")

If lClp.NumberOfArgs <> 1 Then gWriteErrorLine "There must be only one argument, the target order type. Additionally there can be several attributes beginnig with ""/"""

Dim lOrderTypeStr As String: lOrderTypeStr = lClp.Arg(0)
If Not parsePseudoOrder(lOrderTypeStr, lClp.SwitchValue(OffsetSwitch)) Then
    mBOSpec.TargetType = getTargetOrderType(lOrderTypeStr)
    
    If Not IsTargetOrderType(mBOSpec.TargetType) Then
        mBOSpec.TargetType = OrderTypeNone
        gWriteErrorLine "Invalid target order type"
    Else
        Dim lPriceStr As String: lPriceStr = getPriceString(lClp)
        Dim lTriggerPriceStr As String: lTriggerPriceStr = getTriggerPriceString(lClp)
        
        Dim lLimitPriceSpec As PriceSpecifier
        Dim lTriggerPriceSpec As PriceSpecifier
        Dim lMessage As String
        If Not gParseOrderPrices(IIf(lPriceStr <> "", lPriceStr, lTriggerPriceStr), _
                        IIf(lPriceStr <> "", lTriggerPriceStr, ""), _
                        mBOSpec.TargetType, _
                        lLimitPriceSpec, _
                        lTriggerPriceSpec, _
                        lMessage) _
        Then
            gWriteErrorLine lMessage
        End If
        Set mBOSpec.TargetLimitPriceSpec = lLimitPriceSpec
        Set mBOSpec.TargetTriggerPriceSpec = lTriggerPriceSpec
    End If
End If

Dim lSwitchName As String
lSwitchName = TIFSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.TargetTIF = OrderTIFFromString(lClp.SwitchValue(lSwitchName))

lSwitchName = ReasonSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.TargetReason = lClp.SwitchValue(lSwitchName)

lSwitchName = IgnoreRTHSwitch
mBOSpec.TargetIgnoreRTH = lClp.Switch(lSwitchName)

If mBOSpec.EntryType = OrderTypeNone And gErrorCount = 0 Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification
Else
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification, gCommandListOrderCompletion
End If

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrTypeMismatch Then
    gWriteErrorLine lSwitchName & ": invalid value"
    Resume Next
ElseIf Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description
    Resume Next
End If
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Recover( _
                ByVal pPositionManager As PositionManager)
Const ProcName As String = "Recover"
On Error GoTo Err

Set mPositionManager = pPositionManager
Set mContract = mPositionManager.ContractFuture.Value
mGroupName = mPositionManager.GroupName

Set mDataSource = mPositionManager.DataSource
Set mOrderContext = mPositionManager.OrderContexts.DefaultOrderContext

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetDefaultRolloverSpec(ByVal pSpec As RolloverSpecification)
Const ProcName As String = "SetDefaultRolloverSpec"
On Error GoTo Err

If pSpec Is Nothing Then Exit Sub
Set mDefaultRolloverSpec = pSpec.Clone

mDefaultRolloverSpec.setInitialStrikeSelectionMode(mInitialStrikeSelectionMode). _
                    setInitialStrikeParameter(mInitialStrikeParameter). _
                    setInitialStrikeOperator mInitialStrikeOperator

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addPendingCommand( _
                ByVal pTimestamp As Date, _
                ByVal pCommand As Command, _
                ByVal pParams As String, _
                Optional ByVal pOrderId As String, _
                Optional ByVal pMessage As String)
Const ProcName As String = "addPendingCommand"
On Error GoTo Err

Dim lPendingCommand As New PendingCommand
Set lPendingCommand.Command = pCommand
lPendingCommand.Params = pParams
lPendingCommand.OrderID = pOrderId
mPendingCommands.Add lPendingCommand, pTimestamp
If pMessage <> "" Then gWriteLineToConsole pMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getEntryOrderType(ByVal pInput As String) As OrderTypes
Const ProcName As String = "getEntryOrderType"
On Error GoTo Err

If pInput = "" Then
    getEntryOrderType = OrderTypeNone
Else
    getEntryOrderType = OrderTypeFromString(pInput)
End If
AssertArgument IsEntryOrderType(getEntryOrderType)
If Not mOrderContext Is Nothing Then AssertArgument (mOrderContext.PermittedOrderTypes And getEntryOrderType) <> 0

Exit Function

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine "Invalid entry order type: " & pInput
    getEntryOrderType = OrderTypeNone
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Function

Private Function getOrderAction(ByVal pInput As String) As OrderActions
On Error Resume Next
If pInput = "" Then
    getOrderAction = OrderActionNone
Else
    getOrderAction = OrderActionFromString(pInput)
End If
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then getOrderAction = OrderActionNone
End Function

Private Function getPriceString(ByVal pClp As CommandLineParser) As String
Const ProcName As String = "getPriceString"
On Error GoTo Err

getPriceString = pClp.SwitchValue(PriceSwitch)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getStopLossOrderType(ByVal pInput As String) As OrderTypes
Const ProcName As String = "getStopLossOrderType"
On Error GoTo Err

If pInput = "" Then
    getStopLossOrderType = OrderTypeNone
ElseIf pInput = AutoPseudoOrderType Then
    getStopLossOrderType = OrderTypeStop
Else
    getStopLossOrderType = OrderTypeFromString(pInput)
End If
If (mOrderContext.PermittedOrderTypes And getStopLossOrderType) = 0 Then
    gWriteErrorLine "Stop-loss order type not supported: " & pInput
End If

Exit Function

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine "Invalid stop-loss order type: " & pInput
    getStopLossOrderType = OrderTypeNone
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Function

Private Function getTargetOrderType(ByVal pInput As String) As OrderTypes
Const ProcName As String = "getTargetOrderType"
On Error GoTo Err

If pInput = "" Then
    getTargetOrderType = OrderTypeNone
ElseIf pInput = AutoPseudoOrderType Then
    getTargetOrderType = OrderTypeLimit
Else
    getTargetOrderType = OrderTypeFromString(pInput)
End If
If (mOrderContext.PermittedOrderTypes And getTargetOrderType) = 0 Then
    gWriteErrorLine "Target order type not supported: " & pInput
End If

Exit Function

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine "Invalid target order type: " & pInput
    getTargetOrderType = OrderTypeNone
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Function

Private Function getTriggerPriceString(ByVal pClp As CommandLineParser) As String
Const ProcName As String = "getTriggerPriceString"
On Error GoTo Err

Dim s As String
If pClp.Switch(TriggerPriceSwitch) Then
    s = pClp.SwitchValue(TriggerPriceSwitch)
ElseIf pClp.Switch(TriggerPriceSwitch1) Then
    s = pClp.SwitchValue(TriggerPriceSwitch1)
End If
getTriggerPriceString = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub parseOffset( _
                ByVal pValue As String, _
                ByVal pPriceSpec As PriceSpecifier)
pValue = UCase$(pValue)
If pValue = "" Then
    pPriceSpec.Offset = 0#
    pPriceSpec.OffsetType = PriceOffsetTypeNone
ElseIf Right$(Trim$(pValue), 1) = TickOffsetDesignator Then
    Dim numTicks As String
    numTicks = Left$(pValue, Len(pValue) - 1)
    If Not IsInteger(numTicks) Then Err.Raise ErrorCodes.ErrIllegalArgumentException
    pPriceSpec.Offset = CLng(numTicks)
    pPriceSpec.OffsetType = PriceOffsetTypeNumberOfTicks
ElseIf Right$(Trim$(pValue), 1) = PercentOffsetDesignator Then
    ' Note that for legacy purposes, "%" was used for the bid/ask spread
    ' percent designator
    Dim numPercent As String
    numPercent = Left$(pValue, Len(pValue) - 1)
    If Not IsInteger(numPercent) Then Err.Raise ErrorCodes.ErrIllegalArgumentException
    pPriceSpec.Offset = CLng(numPercent)
    pPriceSpec.OffsetType = PriceOffsetTypeBidAskPercent
Else
    pPriceSpec.Offset = CDbl(pValue)
    pPriceSpec.OffsetType = PriceOffsetTypeIncrement
End If
End Sub

Private Function parsePseudoOrder( _
                ByVal pOrderTypeStr, _
                ByVal pOffset As String) As Boolean
Const ProcName As String = "parsePseudoOrder"
On Error GoTo Err

pOrderTypeStr = UCase$(pOrderTypeStr)
If pOrderTypeStr = AskPseudoOrderType Then
    mBOSpec.EntryType = OrderTypeLimit
    mBOSpec.EntryLimitPriceSpec.PriceType = PriceValueTypeAsk
    parseOffset pOffset, mBOSpec.EntryLimitPriceSpec
    parsePseudoOrder = True
ElseIf pOrderTypeStr = BidPseudoOrderType Then
    mBOSpec.EntryType = OrderTypeLimit
    mBOSpec.EntryLimitPriceSpec.PriceType = PriceValueTypeBid
    parseOffset pOffset, mBOSpec.EntryLimitPriceSpec
    parsePseudoOrder = True
ElseIf pOrderTypeStr = LastPseudoOrderType Then
    mBOSpec.EntryType = OrderTypeLimit
    mBOSpec.EntryLimitPriceSpec.PriceType = PriceValueTypeLast
    parseOffset pOffset, mBOSpec.EntryLimitPriceSpec
    parsePseudoOrder = True
Else
    parsePseudoOrder = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function processBuyOrSellCommand( _
                ByVal pAction As OrderActions, _
                ByVal pParams As String, _
                ByVal pID As String) As Boolean
Const ProcName As String = "processBuyOrSellCommand"
On Error GoTo Err

processBuyOrSellCommand = True

Set mBOSpec = New BracketOrderSpecification
Set mBOSpec.Contract = mContract
mBOSpec.ID = pID
mBOSpec.Action = pAction
mBOSpec.EntryTIF = OrderTIFDay
mBOSpec.StopLossTIF = OrderTIFGoodTillCancelled
mBOSpec.TargetTIF = OrderTIFGoodTillCancelled

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pParams)

If UCase$(lClp.Arg(lClp.NumberOfArgs - 1)) = "ROLLOVER" Then
    Set mBOSpec.RolloverSpec = mDefaultRolloverSpec
    pParams = Left$(pParams, Len(pParams) - Len("ROLLOVER"))
    Set lClp = CreateCommandLineParser(pParams)
End If

If Not setQuantity(lClp.Arg(0)) Then
    gWriteErrorLine "Invalid quantity", True
End If

Dim lOrderTypeStr As String: lOrderTypeStr = lClp.Arg(1)
Dim lPriceOrOffset As String: lPriceOrOffset = lClp.Arg(2)
Dim lTriggerPrice As String: lTriggerPrice = lClp.Arg(3)

If Not parsePseudoOrder(lOrderTypeStr, lPriceOrOffset) Then
    mBOSpec.EntryType = getEntryOrderType(lOrderTypeStr)
    
    If Not IsEntryOrderType(mBOSpec.EntryType) Then
        mBOSpec.EntryType = OrderTypeNone
        gWriteErrorLine "Invalid entry order type", True
        processBuyOrSellCommand = False
    Else
        Dim lLimitPriceSpec As PriceSpecifier
        Dim lTriggerPriceSpec As PriceSpecifier
        Dim lMessage As String
        processBuyOrSellCommand = gParseOrderPrices( _
                                        lPriceOrOffset, _
                                        lTriggerPrice, _
                                        mBOSpec.EntryType, _
                                        lLimitPriceSpec, _
                                        lTriggerPriceSpec, _
                                        lMessage)
        If Not processBuyOrSellCommand Then
            gWriteErrorLine lMessage
        End If
        Set mBOSpec.EntryLimitPriceSpec = lLimitPriceSpec
        Set mBOSpec.EntryTriggerPriceSpec = lTriggerPriceSpec
        
        If mBOSpec.EntryLimitPriceSpec.PriceType = PriceValueTypeEntry Or _
            mBOSpec.EntryTriggerPriceSpec.PriceType = PriceValueTypeEntry Then
            processBuyOrSellCommand = False
            gWriteErrorLine "Invalid use of 'ENTRY'", True
        End If
    End If
End If

Dim lSwitchName As String
lSwitchName = TIFSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.EntryTIF = OrderTIFFromString(lClp.SwitchValue(lSwitchName))

lSwitchName = IgnoreRTHSwitch
mBOSpec.EntryIgnoreRTH = lClp.Switch(lSwitchName)

If processBuyOrSellCommand Then
    mBracketOrders.Add mBOSpec
    
    If pAction = OrderActionBuy Then
        mLatestBuyCommandParams = pParams
    Else
        mLatestSellCommandParams = pParams
    End If
End If



gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrTypeMismatch Then
    gWriteErrorLine lSwitchName & ": invalid value", True
    processBuyOrSellCommand = False
    Resume Next
ElseIf Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description, True
    processBuyOrSellCommand = False
    Resume Next
End If
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processPendingCommands()
Const ProcName As String = "processPendingCommands"
On Error GoTo Err

Dim c As PendingCommand
For Each c In mPendingCommands
    If c.Command Is gCommands.BracketCommand Then
        ProcessBracketCommand c.Params, c.OrderID
    ElseIf c.Command Is gCommands.BuyCommand Then
        ProcessBuyCommand c.Params, c.OrderID
    ElseIf c.Command Is gCommands.EndBracketCommand Then
        ProcessEndBracketCommand
    ElseIf c.Command Is gCommands.EntryCommand Then
        ProcessEntryCommand c.Params
    ElseIf c.Command Is gCommands.QuitCommand Then
        ProcessQuitCommand
    ElseIf c.Command Is gCommands.RolloverCommand Then
        ProcessRolloverCommand c.Params
    ElseIf c.Command Is gCommands.SellCommand Then
        ProcessSellCommand c.Params, c.OrderID
    ElseIf c.Command Is gCommands.StopLossCommand Then
        ProcessStopLossCommand c.Params
    ElseIf c.Command Is gCommands.TargetCommand Then
        ProcessTargetCommand c.Params
    End If
Next

gPlaceOrdersTask.AddContractProcessor Me, mStageOrders
mPendingCommands.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setQuantity( _
                ByVal pValue As String) As Boolean
Const ProcName As String = "setQuantity"
On Error GoTo Err

Const QuantityFormat As String = "^(?:(?:(?:([1-9]\d*(?:\.\d+)?)(\%))|(?:(0?\.\d+(?:\.\d+)?)(\%))|(?:(?:([1-9]\d*)(\$?))))(?:\(([1-9]\d*)\))?)$"

gRegExp.Pattern = QuantityFormat

Dim lMatches As MatchCollection
Set lMatches = gRegExp.Execute(Trim$(pValue))

If lMatches.Count <> 1 Then Exit Function

Dim lMatch As Match: Set lMatch = lMatches(0)

If lMatch.SubMatches(1) = AccountPercentSuffix Then
    mBOSpec.Quantity = 1
    mBOSpec.AutosizeType = BracketAutosizePercentOfAccount
    mBOSpec.AutosizeValue = lMatch.SubMatches(0)
ElseIf lMatch.SubMatches(3) = AccountPercentSuffix Then
    mBOSpec.Quantity = 1
    mBOSpec.AutosizeType = BracketAutosizePercentOfAccount
    mBOSpec.AutosizeValue = lMatch.SubMatches(2)
ElseIf lMatch.SubMatches(5) = MaxOrderCostSuffix Then
    mBOSpec.Quantity = 1
    mBOSpec.AutosizeType = BracketAutosizeMonetaryAmount
    mBOSpec.AutosizeValue = lMatch.SubMatches(4)
Else
    mBOSpec.Quantity = lMatch.SubMatches(4)
    mBOSpec.AutosizeType = BracketAutosizeNone
    mBOSpec.AutosizeValue = 0
End If

If lMatch.SubMatches(6) <> "" Then
    mBOSpec.LotSize = CLng(lMatch.SubMatches(6))
Else
    mBOSpec.LotSize = 1
End If

setQuantity = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub validateRolloverCloseOrEntry( _
                ByVal pValue As String, _
                ByRef pRolloverCloseAtMarket As Boolean, _
                ByRef pRolloverCloseSpreadfactor As Long)
Const ProcName As String = "validateRolloverCloseOrEntry"
On Error GoTo Err

gRegExp.Global = False
gRegExp.IgnoreCase = True
gRegExp.Pattern = _
        "^" & _
            "(mkt)|" & _
            "(?:" & _
                "(lmt)" & _
                "(?:" & _
                    ":(-)?(\d{1,3})" & _
                ")?" & _
            ")" & _
        "$"

Dim lMatches As MatchCollection
Set lMatches = gRegExp.Execute(Trim$(pValue))

If lMatches.Count <> 1 Then
    gWriteErrorLine "Invalid command: syntax error"
Else
    Dim lMatch As Match: Set lMatch = lMatches(0)
    
    If lMatch.SubMatches(0) <> "" Then
        pRolloverCloseAtMarket = True
    Else
        pRolloverCloseSpreadfactor = IIf(lMatch.SubMatches(2) <> "", -1, 1) * _
                                            CLng(lMatch.SubMatches(3))
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


