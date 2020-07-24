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

Private mId                                         As String

Private mContract                                   As IContract
Private mDataSource                                 As IMarketDataSource

Private mBracketOrders                              As New EnumerableCollection

Private mBOSpec                                     As BracketOrderSpecification

Private mPositionManager                            As PositionManager

Private WithEvents mOrderContext                    As OrderContext
Attribute mOrderContext.VB_VarHelpID = -1

Private mContractResolver                           As ContractResolver

Private mGroupName                                  As String

Private mLatestBuyCommandParams                     As String
Private mLatestSellCommandParams                    As String

Private mSelectionMode                              As OptionStrikeSelectionModes
Private mParameter                                  As Long
Private mOperator                                   As OptionStrikeSelectionOperators
Private mUnderlyingExchangeName                     As String

Private mPendingCommands                            As SortedDictionary
Private mStageOrders                                As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mPendingCommands = CreateSortedDictionary(KeyTypeDate)
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

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

Friend Property Get Id() As String
Id = mId
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
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory)
Const ProcName As String = "Activate"
On Error GoTo Err

Set mContract = pcontract
mGroupName = pGroupName

mSelectionMode = mContractResolver.StrikeSelectionMode
mOperator = mContractResolver.StrikeSelectionOperator
mParameter = mContractResolver.StrikeSelectionParameter

mUnderlyingExchangeName = mContractResolver.UnderlyingExchangeName

Set mDataSource = pMarketDataManager.CreateMarketDataSource(CreateFuture(pcontract), False)
mDataSource.StartMarketData

If pScopeName = "" Then
    Set mPositionManager = pOrderManager.CreatePositionManager( _
                                pName, _
                                mDataSource, _
                                pOrderSubmitterFactory, _
                                pGroupName)
Else
    Set mPositionManager = pOrderManager.CreateRecoverablePositionManager( _
                                pName, _
                                mDataSource, _
                                pOrderSubmitterFactory, _
                                pScopeName, _
                                mGroupName)
End If
Set mOrderContext = mPositionManager.OrderContexts.DefaultOrderContext

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
                        lPendingCommand.Params
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
                ByVal pId As String, _
                ByVal pContractResolver As ContractResolver, _
                ByVal pStageOrders As Boolean)
mId = pId
Set mContractResolver = pContractResolver
mStageOrders = pStageOrders
End Sub

Friend Sub ProcessBracketCommand( _
                ByVal pParams As String)
Const ProcName As String = "ProcessBracketCommand"
On Error GoTo Err

If mOrderContext Is Nothing Then
    addPendingCommand GetTimestamp, gCommands.BracketCommand, pParams
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification
    Exit Sub
End If

Set mBOSpec = New BracketOrderSpecification
Set mBOSpec.Contract = mContract
mBOSpec.EntryTIF = OrderTIFDay
mBOSpec.StopLossTIF = OrderTIFGoodTillCancelled
mBOSpec.TargetTIF = OrderTIFGoodTillCancelled

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pParams, " ")

If lClp.NumberOfArgs <> 2 Then gWriteErrorLine "There must be only two arguments, action (BUY or SELL) and size. Additionally there can be several attributes beginnig with ""/"""

mBOSpec.Action = getOrderAction(lClp.Arg(0))
If mBOSpec.Action = OrderActionNone Then gWriteErrorLine "Order action must be BUY or SELL"

Dim lQuantity As String: lQuantity = lClp.Arg(1)
If lQuantity = "" Then
    gWriteErrorLine ("Quantity is missing")
ElseIf Not IsInteger(lQuantity, 1) Then
    gWriteErrorLine ("Quantity must be an integer > 0")
Else
    mBOSpec.Quantity = CInt(lQuantity)
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
                ByVal pParams As String) As Boolean
Const ProcName As String = "ProcessBuyCommand"
On Error GoTo Err

If mOrderContext Is Nothing Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation
    addPendingCommand GetTimestamp, _
                            gCommands.BuyCommand, _
                            pParams, _
                            OrderPendingMessage
Else
    ProcessBuyCommand = processBuyOrSellCommand(OrderActionBuy, pParams)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function ProcessSellCommand( _
                ByVal pParams As String) As Boolean
Const ProcName As String = "ProcessSellCommand"
On Error GoTo Err

If mOrderContext Is Nothing Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation
    addPendingCommand GetTimestamp, _
                            gCommands.SellCommand, _
                            pParams, _
                            OrderPendingMessage
Else
    ProcessSellCommand = processBuyOrSellCommand(OrderActionSell, pParams)
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
                            OrderPendingMessage
    Exit Sub
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
        parseOrderPrices IIf(lPriceStr <> "", lPriceStr, lTriggerPriceStr), _
                        IIf(lPriceStr <> "", lTriggerPriceStr, ""), _
                        mBOSpec.EntryType, _
                        lLimitPriceSpec, _
                        lTriggerPriceSpec
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

If Not (mContract.Specifier.SecType = SecTypeFuture Or _
        mContract.Specifier.SecType = SecTypeFuturesOption Or _
        mContract.Specifier.SecType = SecTypeOption) Then
    gWriteErrorLine "Rollover only applies to options, futures and futures options"
    Exit Sub
End If

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pParams, " ")

If lClp.NumberOfArgs <> 0 Then
    gWriteErrorLine "Only attributes beginning with ""/"" are permitted"
    Exit Sub
End If

Dim lDays As Long
If lClp.Switch(DaysSwitch) Then lDays = processRolloverDays(lClp.SwitchValue(DaysSwitch))

Dim lTime As Date
If lClp.Switch(TimeSwitch) Then lTime = processRolloverTime(lClp.SwitchValue(TimeSwitch))

Dim lCloseType As OrderTypes
Dim lCloseLimitPriceSpec As New PriceSpecifier
Dim lCloseTriggerPriceSpec As New PriceSpecifier
Dim lCloseTimeout As Long
If lClp.Switch(CloseSwitch) Then processRolloverClose _
                                        lClp.SwitchValue(CloseSwitch), _
                                        lCloseType, _
                                        lCloseLimitPriceSpec, _
                                        lCloseTriggerPriceSpec, _
                                        lCloseTimeout

Dim lEntryType As OrderTypes
Dim lEntryLimitPriceSpec As New PriceSpecifier
Dim lEntryTriggerPriceSpec As New PriceSpecifier
Dim lEntryTimeout As Long
If lClp.Switch(EntrySwitch) Then processRolloverEntry _
                                        lClp.SwitchValue(EntrySwitch), _
                                        lEntryType, _
                                        lEntryLimitPriceSpec, _
                                        lEntryTriggerPriceSpec, _
                                        lEntryTimeout

Dim lStrikeMode As RolloverStrikeModes: lStrikeMode = RolloverStrikeModeCurrentValue
Dim lStrikeValue As Long: lStrikeValue = 0

If lClp.Switch(StrikeSwitch) Then processRolloverStrike _
                                        lClp.SwitchValue(StrikeSwitch), _
                                        lStrikeMode, _
                                        lStrikeValue

If mContract.Specifier.SecType = SecTypeOption Or _
    mContract.Specifier.SecType = SecTypeFuturesOption _
Then
    Set mBOSpec.RolloverSpec = CreateOptionRolloverSpecification( _
                                lDays, _
                                lTime, _
                                lStrikeMode, _
                                lStrikeValue, _
                                mUnderlyingExchangeName, _
                                lCloseType, _
                                lCloseTimeout, _
                                lCloseLimitPriceSpec, _
                                lCloseTriggerPriceSpec, _
                                lEntryType, _
                                lEntryTimeout, _
                                lEntryLimitPriceSpec, _
                                lEntryTriggerPriceSpec)
Else
    Set mBOSpec.RolloverSpec = CreateRolloverSpecification( _
                                lDays, _
                                lTime, _
                                lCloseType, _
                                lCloseTimeout, _
                                lCloseLimitPriceSpec, _
                                lCloseTriggerPriceSpec, _
                                lEntryType, _
                                lEntryTimeout, _
                                lEntryLimitPriceSpec, _
                                lEntryTriggerPriceSpec)
End If

If mBOSpec.EntryType = OrderTypeNone And gErrorCount = 0 Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification
Else
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderSpecification, gCommandListOrderCompletion
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ProcessQuitCommand()
Const ProcName As String = "ProcessQuitCommand"
On Error GoTo Err

If mOrderContext Is Nothing Then
    gSetValidNextCommands gCommandListAlways, gCommandListGeneral, gCommandListOrderCreation
    addPendingCommand GetTimestamp, _
                            gCommands.QuitCommand, _
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
        gWriteErrorLine "Invalid stop-loss order type", True
    Else
        Dim lPriceStr As String: lPriceStr = getPriceString(lClp)
        Dim lTriggerPriceStr As String: lTriggerPriceStr = getTriggerPriceString(lClp)
        
        Dim lLimitPriceSpec As PriceSpecifier
        Dim lTriggerPriceSpec As PriceSpecifier
        parseOrderPrices IIf(lPriceStr <> "", lPriceStr, lTriggerPriceStr), _
                        IIf(lPriceStr <> "", lTriggerPriceStr, ""), _
                        mBOSpec.StopLossType, _
                        lLimitPriceSpec, _
                        lTriggerPriceSpec
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
        gWriteErrorLine "Invalid target order type", True
    Else
        Dim lPriceStr As String: lPriceStr = getPriceString(lClp)
        Dim lTriggerPriceStr As String: lTriggerPriceStr = getTriggerPriceString(lClp)
        
        Dim lLimitPriceSpec As PriceSpecifier
        Dim lTriggerPriceSpec As PriceSpecifier
        parseOrderPrices IIf(lPriceStr <> "", lPriceStr, lTriggerPriceStr), _
                        IIf(lPriceStr <> "", lTriggerPriceStr, ""), _
                        mBOSpec.TargetType, _
                        lLimitPriceSpec, _
                        lTriggerPriceSpec
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

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addPendingCommand( _
                ByVal pTimestamp As Date, _
                ByVal pCommand As Command, _
                ByVal pParams As String, _
                Optional ByVal pMessage As String)
Const ProcName As String = "addPendingCommand"
On Error GoTo Err

Dim lPendingCommand As New PendingCommand
Set lPendingCommand.Command = pCommand
lPendingCommand.Params = pParams
mPendingCommands.Add lPendingCommand, pTimestamp
If pMessage <> "" Then gWriteLineToConsole pMessage, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getOrderAction(ByVal pInput As String) As OrderActions
On Error Resume Next
If pInput = "" Then
    getOrderAction = OrderActionNone
Else
    getOrderAction = OrderActionFromString(pInput)
End If
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then getOrderAction = OrderActionNone
End Function

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

Private Function getPriceString(ByVal pClp As CommandLineParser) As String
Const ProcName As String = "getPriceString"
On Error GoTo Err

getPriceString = pClp.SwitchValue(PriceSwitch)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Private Function parseOrderPrices( _
                ByVal pPrice1 As String, _
                ByVal pPrice2 As String, _
                ByVal pOrderType As OrderTypes, _
                ByRef pLimitPriceSpec As PriceSpecifier, _
                ByRef pTriggerPriceSpec As PriceSpecifier) As Boolean
Const ProcName As String = "parseOrderPrices"
On Error GoTo Err

parseOrderPrices = True

Set pLimitPriceSpec = NewPriceSpecifier
Set pTriggerPriceSpec = NewPriceSpecifier

Select Case pOrderType
Case OrderTypeMarket, _
        OrderTypeMarketOnClose, _
        OrderTypeMarketOnOpen, _
        OrderTypeMarketToLimit
    ' nothing to do
Case OrderTypeMarketIfTouched, _
        OrderTypeStop, _
        OrderTypeTrail
    If Not parseOrderPrice(pPrice1, "Trigger", pTriggerPriceSpec) Then parseOrderPrices = False
Case OrderTypeLimit, _
        OrderTypeLimitOnOpen, _
        OrderTypeLimitOnClose
    If Not parseOrderPrice(pPrice1, "Limit", pLimitPriceSpec) Then parseOrderPrices = False
Case OrderTypeLimitIfTouched, _
        OrderTypeStopLimit, _
        OrderTypeTrailLimit
    If Not parseOrderPrice(pPrice1, "Trigger", pTriggerPriceSpec) Then parseOrderPrices = False
    If Not parseOrderPrice(pPrice2, "Limit", pLimitPriceSpec) Then parseOrderPrices = False
End Select
    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub parseOrderSpec( _
                ByVal pValue As String, _
                ByRef pOrderType As OrderTypes, _
                ByRef pLimitPriceSpec As PriceSpecifier, _
                ByRef pTriggerPriceSpec As PriceSpecifier, _
                ByRef pTimeoutSecs As Long)
Const ProcName As String = "parseOrderSpec"
On Error GoTo Err

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pValue, ";")

Dim lOrderTypeStr As String: lOrderTypeStr = lClp.Arg(0)

If lOrderTypeStr = "" Then
    pOrderType = OrderTypeNone
Else
    pOrderType = OrderTypeFromString(lOrderTypeStr)
End If
If Not mOrderContext Is Nothing Then AssertArgument (mOrderContext.PermittedOrderTypes And pOrderType) <> 0

Dim lMaxNumberOfArgsRequired As Long
Dim lPrice1Spec As String
Dim lPrice2Spec As String
Dim lTimeoutSpec As String

Select Case pOrderType
Case OrderTypeMarket, _
        OrderTypeMarketOnClose, _
        OrderTypeMarketOnOpen, _
        OrderTypeMarketToLimit
    lMaxNumberOfArgsRequired = 1
Case OrderTypeMarketIfTouched, _
        OrderTypeStop, _
        OrderTypeTrail, _
        OrderTypeLimit, _
        OrderTypeLimitOnOpen, _
        OrderTypeLimitOnClose, _
        OrderTypeLimitIfTouched, _
        OrderTypeStopLimit, _
        OrderTypeTrailLimit
    lMaxNumberOfArgsRequired = 3
    lPrice1Spec = lClp.Arg(1)
    lTimeoutSpec = lClp.Arg(2)
Case Else
    lMaxNumberOfArgsRequired = 4
    lPrice1Spec = lClp.Arg(1)
    lPrice2Spec = lClp.Arg(2)
    lTimeoutSpec = lClp.Arg(3)
End Select

If lClp.NumberOfArgs > lMaxNumberOfArgsRequired Then
    gWriteErrorLine "Too many elements in order spec"
End If

parseOrderPrices lPrice1Spec, _
                    lPrice2Spec, _
                    pOrderType, _
                    pLimitPriceSpec, _
                    pTriggerPriceSpec

pTimeoutSecs = parseOrderTimeoutSecs(lTimeoutSpec)

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine "Invalid order type: " & lOrderTypeStr
    pOrderType = OrderTypeNone
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Sub

Private Function parseOrderTimeoutSecs( _
                ByVal pTimeoutSpec As String) As Double
Const SecondsDesignator As String = "S"
Const MinutesDesignator As String = "M"

If pTimeoutSpec = "" Then
    parseOrderTimeoutSecs = 0
    Exit Function
End If

Dim lUnits As String
Dim lTimeout As String
If UCase$(Right$(pTimeoutSpec, 1)) = SecondsDesignator Then
    lUnits = SecondsDesignator
    lTimeout = Left$(pTimeoutSpec, Len(pTimeoutSpec) - 1)
ElseIf UCase$(Right$(pTimeoutSpec, 1)) = MinutesDesignator Then
    lUnits = MinutesDesignator
    lTimeout = Left$(pTimeoutSpec, Len(pTimeoutSpec) - 1)
Else
    gWriteErrorLine "Timeout spec invalid - must be a positive integer " & _
                        "followed by ""M"" for minutes or ""S"" for seconds " & _
                        "(not case-sensitive)"
    parseOrderTimeoutSecs = 0
    Exit Function
End If

If IsInteger(lTimeout, 0) Then
    parseOrderTimeoutSecs = CLng(lTimeout)
    If lUnits = MinutesDesignator Then parseOrderTimeoutSecs = parseOrderTimeoutSecs * 60
Else
    gWriteErrorLine "Invalid timeout: " & pTimeoutSpec
    parseOrderTimeoutSecs = 0
End If
End Function

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

Private Function parseOrderPrice( _
                ByVal pPriceStr As String, _
                ByVal pPriceName As String, _
                ByRef pTriggerPriceSpec As PriceSpecifier) As Boolean
Const ProcName As String = "parseOrderPrice"
On Error GoTo Err

If mOrderContext.ParsePriceAndOffset(pTriggerPriceSpec, pPriceStr) And _
    pTriggerPriceSpec.IsValid _
Then
    parseOrderPrice = True
Else
    parseOrderPrice = False
    gWriteErrorLine pPriceName & " price must be specified for this order type"
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function processBuyOrSellCommand( _
                ByVal pAction As OrderActions, _
                ByVal pParams As String) As Boolean
Const ProcName As String = "processBuyOrSellCommand"
On Error GoTo Err

processBuyOrSellCommand = True

Set mBOSpec = New BracketOrderSpecification
Set mBOSpec.Contract = mContract
mBOSpec.Action = pAction
mBOSpec.EntryTIF = OrderTIFDay
mBOSpec.StopLossTIF = OrderTIFGoodTillCancelled
mBOSpec.TargetTIF = OrderTIFGoodTillCancelled

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pParams)

Dim lQuantity As String: lQuantity = lClp.Arg(0)
If lQuantity = "" Then
    processBuyOrSellCommand = False
    gWriteErrorLine "Quantity is missing", True
ElseIf Not IsInteger(lQuantity, 1) Then
    processBuyOrSellCommand = False
    gWriteErrorLine "Quantity must be an integer > 0", True
Else
    mBOSpec.Quantity = CInt(lQuantity)
End If

Dim lOrderTypeStr As String: lOrderTypeStr = lClp.Arg(1)
Dim lPriceOrOffset As String: lPriceOrOffset = lClp.Arg(2)

If Not parsePseudoOrder(lOrderTypeStr, lPriceOrOffset) Then
    mBOSpec.EntryType = getEntryOrderType(lClp.Arg(1))
    
    If Not IsEntryOrderType(mBOSpec.EntryType) Then
        mBOSpec.EntryType = OrderTypeNone
        gWriteErrorLine "Invalid entry order type", True
        processBuyOrSellCommand = False
    Else
        Dim lLimitPriceSpec As PriceSpecifier
        Dim lTriggerPriceSpec As PriceSpecifier
        
        processBuyOrSellCommand = parseOrderPrices( _
                                        lPriceOrOffset, _
                                        lClp.Arg(3), _
                                        mBOSpec.EntryType, _
                                        lLimitPriceSpec, _
                                        lTriggerPriceSpec)
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
    gWriteLineToConsole c.Command.Name & " " & c.Params, True
    If c.Command Is gCommands.BracketCommand Then
        ProcessBracketCommand c.Params
    ElseIf c.Command Is gCommands.BuyCommand Then
        ProcessBuyCommand c.Params
    ElseIf c.Command Is gCommands.EndBracketCommand Then
        ProcessEndBracketCommand
    ElseIf c.Command Is gCommands.EntryCommand Then
        ProcessEntryCommand c.Params
    ElseIf c.Command Is gCommands.QuitCommand Then
        ProcessQuitCommand
    ElseIf c.Command Is gCommands.RolloverCommand Then
        ProcessRolloverCommand c.Params
    ElseIf c.Command Is gCommands.SellCommand Then
        ProcessSellCommand c.Params
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

Private Sub processRolloverClose( _
                ByVal pValue As String, _
                ByRef pCloseType As OrderTypes, _
                ByRef pCloseLimitPriceSpec As PriceSpecifier, _
                ByRef pCloseTriggerPriceSpec As PriceSpecifier, _
                ByRef pCloseTimeoutSecs As Long)
Const ProcName As String = "processRolloverClose"
On Error GoTo Err

parseOrderSpec pValue, pCloseType, pCloseLimitPriceSpec, pCloseTriggerPriceSpec, pCloseTimeoutSecs

If Not IsEntryOrderType(pCloseType) Then
    pCloseType = OrderTypeNone
    Set pCloseLimitPriceSpec = New PriceSpecifier
    Set pCloseTriggerPriceSpec = New PriceSpecifier
    pCloseTimeoutSecs = 0
    gWriteErrorLine "Invalid close order type", True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processRolloverDays(ByVal pValue As String) As Long
If Not IsInteger(pValue, 0, 30) Then
    gWriteErrorLine "Days must be an integer from 0 to 30"
Else
    processRolloverDays = CLng(pValue)
End If
End Function

Private Sub processRolloverEntry( _
                ByVal pValue As String, _
                ByRef pEntryType As OrderTypes, _
                ByRef pEntryLimitPriceSpec As PriceSpecifier, _
                ByRef pEntryTriggerPriceSpec As PriceSpecifier, _
                ByRef pEntryTimeoutSecs As Long)
Const ProcName As String = "processRolloverEntry"
On Error GoTo Err

parseOrderSpec pValue, pEntryType, pEntryLimitPriceSpec, pEntryTriggerPriceSpec, pEntryTimeoutSecs

If Not IsEntryOrderType(pEntryType) Then
    pEntryType = OrderTypeNone
    Set pEntryLimitPriceSpec = New PriceSpecifier
    Set pEntryTriggerPriceSpec = New PriceSpecifier
    pEntryTimeoutSecs = 0
    gWriteErrorLine "Invalid entry order type", True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processRolloverStrike( _
                ByVal pValue As String, _
                ByRef pStrikeMode As RolloverStrikeModes, _
                ByRef pStrikeValue As Long)
Const ProcName As String = "processRolloverStrike"
On Error GoTo Err

Const StrikeFormat As String = "^(?:((?:[1-9]\d{1,6})|(\d\d?\d?)(?:%)|<|)(\$))|(\+|-)([1-9]\d{0,1})$"
gRegExp.Pattern = StrikeFormat

Dim lMatches As MatchCollection
Set lMatches = gRegExp.Execute(pValue)

If lMatches.Count <> 1 Then
    gWriteErrorLine "Invalid " & StrikeSwitch & " syntax"
Else
    Dim lMatch As Match: Set lMatch = lMatches(0)

    Dim lAmount As String: lAmount = lMatch.SubMatches(0)
    Dim lProfitPercentage As Long: lProfitPercentage = lMatch.SubMatches(1)
    Dim lDollar As String: lDollar = lMatch.SubMatches(2)
    Dim lSign As String: lSign = lMatch.SubMatches(3)
    Dim lIncrement As String: lIncrement = lMatch.SubMatches(4)
    
    If lDollar = "$" Then
        If lAmount = "" Then
            pStrikeMode = RolloverStrikeModeCurrentValue
            pStrikeValue = 0
        ElseIf lAmount = "<" Then
            pStrikeMode = RolloverStrikeModePreviousAmount
            pStrikeValue = mParameter
        ElseIf lProfitPercentage <> 0 Then
            pStrikeMode = RolloverStrikeModeCurrentProfit
            pStrikeValue = CLng(lProfitPercentage)
        End If
    ElseIf lProfitPercentage <> "" Then
        pStrikeMode = RolloverStrikeModeCurrentProfit
        pStrikeValue = CLng(lProfitPercentage)
    Else
        pStrikeMode = RolloverStrikeModeIncrement
        pStrikeValue = lIncrement
        If lSign = "-" Then pStrikeValue = -pStrikeValue
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processRolloverTime(ByVal pValue As String) As Date
If Not IsDate(pValue) Then
    gWriteErrorLine "Time must be hh:mm[:ss]"
Else
    processRolloverTime = CDate(pValue)
    If CDbl(processRolloverTime) >= 1 Then gWriteErrorLine "Time must be hh:mm[:ss]"
End If
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


