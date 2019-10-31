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

Event IsReady()

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

'@================================================================================
' Member variables
'@================================================================================

Private mMarketDataManager                          As IMarketDataManager

Private mContract                                   As IContract
Private mDataSource                                 As IMarketDataSource

Private mBracketOrders                              As New EnumerableCollection

Private mBOSpec                                     As BracketOrderSpecification

Private mOrderManager                               As OrderManager

Private mOrderSubmitterFactory                      As IOrderSubmitterFactory
Private mPositionManager                            As PositionManager

Private WithEvents mOrderContext                    As OrderContext
Attribute mOrderContext.VB_VarHelpID = -1

Private mScopeName                                  As String
Private mGroupName                                  As String

Private mLatestBuyCommandParams                     As String
Private mLatestSellCommandParams                    As String

'@================================================================================
' Class Event Handlers
'@================================================================================

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

If ev.ChangeType = OrderContextChangeTypes.OrderContextReadyStateChanged And _
    mOrderContext.IsReady _
Then
    RaiseEvent IsReady
End If
    
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
ContractName = mContract.Specifier.LocalSymbol & "@" & mContract.Specifier.Exchange
End Property

Friend Property Get DataSource() As IMarketDataSource
Set DataSource = mDataSource
End Property

Friend Property Get GroupName() As String
GroupName = mGroupName
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

'@================================================================================
' Methods
'@================================================================================

Friend Sub ClearOrders()
mBracketOrders.Clear
End Sub

Friend Sub Initialise( _
                ByVal pName As String, _
                ByVal pContractFuture As IFuture, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pOrderManager As OrderManager, _
                ByVal pScopeName As String, _
                ByVal pGroupName As String, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContract = pContractFuture.Value
Set mMarketDataManager = pMarketDataManager
Set mOrderManager = pOrderManager
mScopeName = pScopeName
mGroupName = pGroupName
Set mOrderSubmitterFactory = pOrderSubmitterFactory

Set mDataSource = mMarketDataManager.CreateMarketDataSource(pContractFuture, False)
mDataSource.StartMarketData

If mScopeName = "" Then
    Set mPositionManager = mOrderManager.CreatePositionManager( _
                                pName, _
                                mDataSource, _
                                mOrderSubmitterFactory)
Else
    Set mPositionManager = mOrderManager.CreateRecoverablePositionManager( _
                                pName, _
                                mDataSource, _
                                mOrderSubmitterFactory, _
                                mScopeName, _
                                mGroupName)
End If
Set mOrderContext = mPositionManager.OrderContexts.DefaultOrderContext

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ProcessBracketCommand( _
                ByVal pParams As String)
Const ProcName As String = "ProcessBracketCommand"
On Error GoTo Err

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

gSetValidNextCommands QuoteCommand, EntryCommand, StopLossCommand, TargetCommand, RolloverCommand, BracketCommand, ResetCommand, CloseoutCommand

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

ProcessBuyCommand = processBuyOrSellCommand(OrderActionBuy, pParams)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function ProcessSellCommand( _
                ByVal pParams As String) As Boolean
Const ProcName As String = "ProcessSellCommand"
On Error GoTo Err

ProcessSellCommand = processBuyOrSellCommand(OrderActionSell, pParams)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ProcessEndBracketCommand()
Const ProcName As String = "ProcessEndBracketCommand"
On Error GoTo Err

If gErrorCount = 0 Then mBracketOrders.Add mBOSpec

gSetValidNextCommands ListCommand, GroupCommand, QuoteCommand, BracketCommand, BuyCommand, SellCommand, SellAgainCommand, ContractCommand, EndOrdersCommand, ResetCommand, CloseoutCommand

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ProcessEntryCommand( _
                ByVal pParams As String)
Const ProcName As String = "ProcessEntryCommand"
On Error GoTo Err

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
    gSetValidNextCommands QuoteCommand, EntryCommand, StopLossCommand, TargetCommand, ResetCommand
Else
    gSetValidNextCommands QuoteCommand, EntryCommand, StopLossCommand, TargetCommand, RolloverCommand, EndBracketCommand, ResetCommand
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

If Not (mContract.Specifier.SecType = SecTypeFuture Or _
        mContract.Specifier.SecType = SecTypeFuturesOption Or _
        mContract.Specifier.SecType = SecTypeOption) Then
    gWriteErrorLine "Rollover only applies to options, futures and futures options"
End If

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pParams, " ")

If lClp.NumberOfArgs <> 0 Then gWriteErrorLine "Only attributes beginning with ""/"" are permitted"

Dim lSwitchName As String
lSwitchName = DaysSwitch
If lClp.Switch(lSwitchName) Then
    If Not IsInteger(lClp.SwitchValue(lSwitchName), 0, 10) Then
        gWriteErrorLine "Days must be an integer from 0 to 10"
    Else
        mBOSpec.RolloverDays = CLng(lClp.SwitchValue(lSwitchName))
    End If
End If

lSwitchName = TimeSwitch
If lClp.Switch(lSwitchName) Then
    If Not IsDate(lClp.SwitchValue(lSwitchName)) Then
        gWriteErrorLine "Time must be hh:mm[:ss]"
    Else
        Dim lDate As Date
        lDate = CDate(lClp.SwitchValue(lSwitchName))
        If CDbl(lDate) >= 1 Then
            gWriteErrorLine "Time must be hh:mm[:ss]"
        Else
            mBOSpec.RolloverTime = lDate
        End If
    End If
End If

lSwitchName = CloseSwitch
If lClp.Switch(lSwitchName) Then
    Dim lRolloverCloseAtMarket As Boolean
    Dim lRolloverCloseSpreadfactor As Long
    validateRolloverCloseOrEntry lClp.SwitchValue(lSwitchName), _
                                lRolloverCloseAtMarket, _
                                lRolloverCloseSpreadfactor
    mBOSpec.RolloverCloseAtMarket = lRolloverCloseAtMarket
    mBOSpec.RolloverCloseSpreadfactor = lRolloverCloseSpreadfactor
End If

lSwitchName = EntrySwitch
If lClp.Switch(lSwitchName) Then
    Dim lRolloverEntryAtMarket As Boolean
    Dim lRolloverEntrySpreadfactor As Long
    validateRolloverCloseOrEntry lClp.SwitchValue(lSwitchName), _
                                lRolloverCloseAtMarket, _
                                lRolloverCloseSpreadfactor
    mBOSpec.RolloverEntryAtMarket = lRolloverEntryAtMarket
    mBOSpec.RolloverEntrySpreadfactor = lRolloverEntrySpreadfactor
End If

mBOSpec.Rollover = True

If mBOSpec.EntryType = OrderTypeNone And gErrorCount = 0 Then
    gSetValidNextCommands QuoteCommand, EntryCommand, StopLossCommand, TargetCommand, RolloverCommand, ResetCommand
Else
    gSetValidNextCommands QuoteCommand, EntryCommand, StopLossCommand, TargetCommand, RolloverCommand, EndBracketCommand, ResetCommand
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


Friend Sub ProcessStopLossCommand( _
                ByVal pParams As String)
Const ProcName As String = "ProcessStopLossCommand"
On Error GoTo Err

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
    gSetValidNextCommands QuoteCommand, EntryCommand, StopLossCommand, TargetCommand, ResetCommand
Else
    gSetValidNextCommands QuoteCommand, EntryCommand, StopLossCommand, TargetCommand, RolloverCommand, EndBracketCommand, ResetCommand
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
    gSetValidNextCommands QuoteCommand, EntryCommand, StopLossCommand, TargetCommand, ResetCommand
Else
    gSetValidNextCommands QuoteCommand, EntryCommand, StopLossCommand, TargetCommand, RolloverCommand, EndBracketCommand, ResetCommand
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

'@================================================================================
' Helper Functions
'@================================================================================

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
AssertArgument (mOrderContext.PermittedOrderTypes And getEntryOrderType) <> 0

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
                ByVal pArg1 As String, _
                ByVal pArg2 As String, _
                ByVal pOrderType As OrderTypes, _
                ByRef pLimitPriceSpec As PriceSpecifier, _
                ByRef pTriggerPriceSpec As PriceSpecifier) As Boolean
Const ProcName As String = "parseOrderPrices"
On Error GoTo Err

parseOrderPrices = True

Set pLimitPriceSpec = NewPriceSpecifier
Set pTriggerPriceSpec = NewPriceSpecifier

Select Case pOrderType
Case OrderTypeMarketIfTouched, _
        OrderTypeStop, _
        OrderTypeTrail
    Set pTriggerPriceSpec = parsePrice(pArg1)
    If pTriggerPriceSpec.PriceType = PriceValueTypeNone Then
        parseOrderPrices = False
        gWriteErrorLine "Trigger price must be specified for this order type", True
    End If
Case OrderTypeMarketToLimit, _
        OrderTypeLimit, _
        OrderTypeLimitOnOpen, _
        OrderTypeLimitOnClose, _
        OrderTypeLimitIfTouched, _
        OrderTypeStopLimit, _
        OrderTypeTrailLimit
    Set pLimitPriceSpec = parsePrice(pArg1)
    If pLimitPriceSpec.PriceType = PriceValueTypeNone Then
        parseOrderPrices = False
        gWriteErrorLine "Limit price must be specified for this order type", True
    End If
Case Else
    If pArg1 <> "" Then
        parseOrderPrices = False
        gWriteErrorLine "Price must not be specified for this order type", True
    End If
End Select
    
Select Case pOrderType
Case OrderTypeLimitIfTouched, _
        OrderTypeStopLimit, _
        OrderTypeTrailLimit
    Set pTriggerPriceSpec = parsePrice(pArg2)
    If pTriggerPriceSpec.PriceType = PriceValueTypeNone Then
        parseOrderPrices = False
        gWriteErrorLine "Trigger price must be specified for this order type", True
    End If
Case Else
    If pArg2 <> "" Then
        parseOrderPrices = False
        gWriteErrorLine "Trigger price must not be specified for this order type", True
    End If
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Private Function parsePrice( _
                ByVal pValue As String) As PriceSpecifier
Const ProcName As String = "parsePrice"
On Error GoTo Err

gRegExp.Global = False
gRegExp.IgnoreCase = True

Dim p As String
p = "(?:^(?:(ASK|BID|LAST|ENTRY|(?:[-+]?\d{1,6}(?:.\d{1,6})?)))(?:\[(?:([-+]?\d{1,6}(?:.\d{1,6})?))([T%S]?)\])?$)"

gRegExp.Pattern = p

Dim lMatches As MatchCollection
Set lMatches = gRegExp.Execute(Trim$(pValue))

If lMatches.Count <> 1 Then
    gWriteErrorLine "Invalid command: syntax error in " & pValue, True
    Set parsePrice = NewPriceSpecifier()
    Exit Function
End If

Dim lMatch As Match: Set lMatch = lMatches(0)

Dim lPrice As Double
Dim lPriceType As PriceValueTypes
Dim lOffset As Double
Dim lOffsetType As PriceOffsetTypes

Dim lPricePart As String: lPricePart = UCase$(lMatch.SubMatches(0))
Select Case lPricePart
Case ""
    lPriceType = PriceValueTypeNone
Case AskPriceDesignator
    lPriceType = PriceValueTypeAsk
Case BidPriceDesignator
    lPriceType = PriceValueTypeBid
Case LastPriceDesignator
    lPriceType = PriceValueTypeLast
Case EntryPriceDesignator
    lPriceType = PriceValueTypeEntry
Case Else
    lPriceType = PriceValueTypeValue
    lPrice = CDbl(lPricePart)
End Select

Dim lOffsetPart As String: lOffsetPart = lMatch.SubMatches(1)
If lOffsetPart <> "" Then lOffset = CDbl(lOffsetPart)

Dim lOffsetDesignator As String: lOffsetDesignator = UCase$(lMatch.SubMatches(2))
Select Case lOffsetDesignator
Case ""
    lOffsetType = PriceOffsetTypeIncrement
Case TickOffsetDesignator
    lOffsetType = PriceOffsetTypeNumberOfTicks
Case PercentOffsetDesignator
    lOffsetType = PriceOffsetTypePercent
Case BidAskSpreadPercentOffsetDesignator
    lOffsetType = PriceOffsetTypeBidAskPercent
End Select

Set parsePrice = NewPriceSpecifier(lPrice, lPriceType, lOffset, lOffsetType)

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

gSetValidNextCommands ContractCommand, QuoteCommand, BuyCommand, SellCommand, SellAgainCommand, BracketCommand, ResetCommand, CloseoutCommand

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


