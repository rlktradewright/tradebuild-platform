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

Private mTimezone                                   As TimeZone

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

Friend Property Get IsReadyToPlaceOrders() As Boolean
Const ProcName As String = "IsReadyToPlaceOrders"
On Error GoTo Err

IsReadyToPlaceOrders = mOrderContext.IsReady

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
Set mTimezone = GetTimeZone(mContract.TimezoneName)
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

Friend Sub PlaceOrder(ByVal pBOS As BracketOrderSpecification, pStage As Boolean)
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

gWriteLineToStdOut "TIME " & FormatTimestamp(mTimezone.ConvertDateLocalToTz(GetTimestamp), TimestampDateAndTimeISO8601)
gWriteLineToStdOut pBOS.ToString

mOrderContext.OrderSubmitter.OrdersAreStaged = pStage

Dim lBO As IBracketOrder
With pBOS
    If .Action = OrderActionBuy Then
        Set lBO = mOrderContext.CreateBuyBracketOrder(.Quantity, _
                          .EntryType, _
                          .EntryPrice, _
                          .EntryOffset, _
                          .EntryTriggerPrice, _
                          .StopLossType, _
                          .StopLossTriggerPrice, _
                          .StopLossOffset, _
                          .StopLossPrice, _
                          .TargetType, _
                          .TargetPrice, _
                          .TargetOffset, _
                          .TargetTriggerPrice)
    Else
        Set lBO = mOrderContext.CreateSellBracketOrder(.Quantity, _
                          .EntryType, _
                          .EntryPrice, _
                          .EntryOffset, _
                          .EntryTriggerPrice, _
                          .StopLossType, _
                          .StopLossTriggerPrice, _
                          .StopLossOffset, _
                          .StopLossPrice, _
                          .TargetType, _
                          .TargetPrice, _
                          .TargetOffset, _
                          .TargetTriggerPrice)
    End If
    
    lBO.CancelAfter = .CancelAfter
    lBO.CancelPrice = .CancelPrice
    lBO.EntryOrder.TimeInForce = .EntryTIF
    If Not lBO.StopLossOrder Is Nothing Then lBO.StopLossOrder.TimeInForce = .StopLossTIF
    If Not lBO.TargetOrder Is Nothing Then lBO.TargetOrder.TimeInForce = .TargetTIF
    If .GoodAfterTime <> 0 Then
        lBO.EntryOrder.GoodAfterTime = .GoodAfterTime
        lBO.EntryOrder.GoodAfterTimeTZ = .TimezoneName
    End If
    If .GoodTillDate <> 0 Then
        lBO.EntryOrder.GoodTillDate = .GoodTillDate
        lBO.EntryOrder.GoodTillDateTZ = .TimezoneName
    End If
End With

mOrderContext.ExecuteBracketOrder lBO
CreateBracketProfitCalculator lBO, mDataSource

gNumberOfOrdersPlaced = gNumberOfOrdersPlaced + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ProcessBracketCommand( _
                ByVal pParams As String)
Const ProcName As String = "ProcessBracketCommand"
On Error GoTo Err

Set mBOSpec = New BracketOrderSpecification
mBOSpec.EntryTIF = OrderTIFDay
mBOSpec.StopLossTIF = OrderTIFGoodTillCancelled
mBOSpec.TargetTIF = OrderTIFGoodTillCancelled

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pParams, " ")

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

lSwitchName = GoodAfterTimeSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.GoodAfterTime = CDate(lClp.SwitchValue(lSwitchName))

lSwitchName = GoodTillDateSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.GoodTillDate = CDate(lClp.SwitchValue(lSwitchName))

lSwitchName = TimezoneSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.TimezoneName = lClp.SwitchValue(lSwitchName)

gSetValidNextCommands EntryCommand, StopLossCommand, TargetCommand, BracketCommand, ResetCommand, CloseoutCommand

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

mBracketOrders.Add mBOSpec

gSetValidNextCommands ListCommand, GroupCommand, BracketCommand, BuyCommand, SellCommand, ContractCommand, EndOrdersCommand, ResetCommand

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

mBOSpec.EntryType = getEntryOrderType(lClp.Arg(0))

Dim lSwitchName As String
lSwitchName = PriceSwitch
If lClp.Switch(PriceSwitch) Then mBOSpec.EntryPrice = CDbl(lClp.SwitchValue(PriceSwitch))

lSwitchName = TriggerPriceSwitch
If lClp.Switch(TriggerPriceSwitch) Then
    mBOSpec.EntryTriggerPrice = CDbl(lClp.SwitchValue(TriggerPriceSwitch))
ElseIf lClp.Switch(TriggerPriceSwitch1) Then
    mBOSpec.EntryTriggerPrice = CDbl(lClp.SwitchValue(TriggerPriceSwitch1))
End If

lSwitchName = OffsetSwitch
If lClp.Switch(OffsetSwitch) Then mBOSpec.EntryOffset = parseOffset(lClp.SwitchValue(OffsetSwitch))

lSwitchName = TIFSwitch
If lClp.Switch(TIFSwitch) Then mBOSpec.EntryTIF = OrderTIFFromString(lClp.SwitchValue(TIFSwitch))

If mBOSpec.EntryType = BracketEntryTypeNone Then
    gSetValidNextCommands EntryCommand, StopLossCommand, TargetCommand, ResetCommand
Else
    gSetValidNextCommands EntryCommand, StopLossCommand, TargetCommand, EndBracketCommand, ResetCommand
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

Friend Sub ProcessStopLossCommand( _
                ByVal pParams As String)
Const ProcName As String = "ProcessStopLossCommand"
On Error GoTo Err

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pParams, " ")

mBOSpec.StopLossType = getStopLossOrderType(lClp.Arg(0))

Dim lSwitchName As String
lSwitchName = PriceSwitch
If lClp.Switch(PriceSwitch) Then mBOSpec.StopLossPrice = CDbl(lClp.SwitchValue(PriceSwitch))

lSwitchName = TriggerPriceSwitch
If lClp.Switch(TriggerPriceSwitch) Then
    mBOSpec.StopLossTriggerPrice = CDbl(lClp.SwitchValue(TriggerPriceSwitch))
ElseIf lClp.Switch(TriggerPriceSwitch1) Then
    mBOSpec.StopLossTriggerPrice = CDbl(lClp.SwitchValue(TriggerPriceSwitch1))
End If

lSwitchName = OffsetSwitch
If lClp.Switch(OffsetSwitch) Then mBOSpec.StopLossOffset = parseOffset(lClp.SwitchValue(OffsetSwitch))

lSwitchName = TIFSwitch
If lClp.Switch(TIFSwitch) Then mBOSpec.StopLossTIF = OrderTIFFromString(lClp.SwitchValue(TIFSwitch))

If mBOSpec.EntryType = BracketEntryTypeNone Then
    gSetValidNextCommands EntryCommand, StopLossCommand, TargetCommand, ResetCommand
Else
    gSetValidNextCommands EntryCommand, StopLossCommand, TargetCommand, EndBracketCommand, ResetCommand
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

mBOSpec.TargetType = getTargetOrderType(lClp.Arg(0))

Dim lSwitchName As String
lSwitchName = PriceSwitch
If lClp.Switch(PriceSwitch) Then mBOSpec.TargetPrice = CDbl(lClp.SwitchValue(PriceSwitch))

lSwitchName = TriggerPriceSwitch
If lClp.Switch(TriggerPriceSwitch) Then
    mBOSpec.TargetTriggerPrice = CDbl(lClp.SwitchValue(TriggerPriceSwitch))
ElseIf lClp.Switch(TriggerPriceSwitch1) Then
    mBOSpec.TargetTriggerPrice = CDbl(lClp.SwitchValue(TriggerPriceSwitch1))
End If

lSwitchName = OffsetSwitch
If lClp.Switch(OffsetSwitch) Then mBOSpec.TargetOffset = parseOffset(lClp.SwitchValue(OffsetSwitch))

lSwitchName = TIFSwitch
If lClp.Switch(TIFSwitch) Then mBOSpec.TargetTIF = OrderTIFFromString(lClp.SwitchValue(TIFSwitch))

If mBOSpec.EntryType = BracketEntryTypeNone Then
    gSetValidNextCommands EntryCommand, StopLossCommand, TargetCommand, ResetCommand
Else
    gSetValidNextCommands EntryCommand, StopLossCommand, TargetCommand, EndBracketCommand, ResetCommand
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

Friend Sub WriteContractSpecToConsole()
Const ProcName As String = "WriteContractSpecToConsole"
On Error GoTo Err

Dim s As String
s = ContractCommand & " "
s = s & gGenerateSwitch(SymbolSwitch, mContract.Specifier.Symbol)
s = s & gGenerateSwitch(LocalSymbolSwitch, mContract.Specifier.LocalSymbol)
s = s & gGenerateSwitch(SecTypeSwitch, SecTypeToShortString(mContract.Specifier.SecType))
s = s & gGenerateSwitch(ExchangeSwitch, mContract.Specifier.Exchange)
If mContract.Specifier.Expiry <> "" Then s = s & gGenerateSwitch(ExpirySwitch, mContract.Specifier.Expiry)
s = s & gGenerateSwitch(CurrencySwitch, mContract.Specifier.CurrencyCode)
If mContract.Specifier.Multiplier <> 1 Then s = s & gGenerateSwitch(MultiplierSwitch, mContract.Specifier.Multiplier)
If mContract.Specifier.Right <> OptNone Then
    s = s & gGenerateSwitch(RightSwitch, OptionRightToString(mContract.Specifier.Right))
    s = s & gGenerateSwitch(StrikeSwitch, mContract.Specifier.Strike)
End If
gWriteLineToStdOut s

Exit Sub

Err:
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

Private Function getEntryOrderType(ByVal pInput As String) As BracketEntryTypes
On Error Resume Next
If pInput = "" Then
    getEntryOrderType = BracketEntryTypeNone
Else
    getEntryOrderType = BracketEntryTypeFromString(pInput)
End If
If (mOrderContext.PermittedOrderTypes And BracketEntryTypeToOrderType(getEntryOrderType)) = 0 Then
    gWriteErrorLine "Entry order type not supported: " & pInput
End If
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine "Invalid entry order type: " & pInput
    getEntryOrderType = BracketEntryTypeNone
End If
End Function

Private Function getStopLossOrderType(ByVal pInput As String) As BracketStopLossTypes
On Error Resume Next
If pInput = "" Then
    getStopLossOrderType = BracketStopLossTypeNone
Else
    getStopLossOrderType = BracketStopLossTypeFromString(pInput)
End If
If (mOrderContext.PermittedOrderTypes And BracketStopLossTypeToOrderType(getStopLossOrderType)) = 0 Then
    gWriteErrorLine "Stop-loss order type not supported: " & pInput
End If
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine "Invalid stop-loss order type: " & pInput
    getStopLossOrderType = BracketStopLossTypeNone
End If
End Function

Private Function getTargetOrderType(ByVal pInput As String) As BracketTargetTypes
On Error Resume Next
If pInput = "" Then
    getTargetOrderType = BracketTargetTypeNone
Else
    getTargetOrderType = BracketTargetTypeFromString(pInput)
End If
If (mOrderContext.PermittedOrderTypes And BracketTargetTypeToOrderType(getTargetOrderType)) = 0 Then
    gWriteErrorLine "Target order type not supported: " & pInput
End If
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine "Invalid target order type: " & pInput
    getTargetOrderType = BracketTargetTypeNone
End If
End Function

Private Function parseOffset(ByVal pValue As String) As Long
If UCase$(Right$(Trim$(pValue), 1)) = TickDesignator Then
    Dim numTicks As String
    numTicks = Left$(pValue, Len(pValue) - 1)
    If Not IsInteger(numTicks) Then Err.Raise ErrorCodes.ErrIllegalArgumentException
    parseOffset = CLng(numTicks)
Else
    parseOffset = CLng(CDbl(pValue) / mContract.TickSize)
End If
End Function

Private Function processBuyOrSellCommand( _
                ByVal pAction As OrderActions, _
                ByVal pParams As String) As Boolean
Const ProcName As String = "processBuyOrSellCommand"
On Error GoTo Err

processBuyOrSellCommand = True

Set mBOSpec = New BracketOrderSpecification
mBOSpec.Action = pAction
mBOSpec.EntryTIF = OrderTIFDay
mBOSpec.StopLossTIF = OrderTIFGoodTillCancelled
mBOSpec.TargetTIF = OrderTIFGoodTillCancelled

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pParams)

Dim lQuantity As String: lQuantity = lClp.Arg(0)
If lQuantity = "" Then
    processBuyOrSellCommand = False
    gWriteErrorLine ("Quantity is missing")
ElseIf Not IsInteger(lQuantity, 1) Then
    processBuyOrSellCommand = False
    gWriteErrorLine ("Quantity must be an integer > 0")
Else
    mBOSpec.Quantity = CInt(lQuantity)
End If

mBOSpec.EntryType = getEntryOrderType(lClp.Arg(1))
If mBOSpec.EntryType = BracketEntryTypeNone Then
    processBuyOrSellCommand = False
    Exit Function
End If

Dim lPriceOrOffset As String: lPriceOrOffset = lClp.Arg(2)
If lPriceOrOffset <> "" Then
    Select Case mBOSpec.EntryType
    Case BracketEntryTypeMarketIfTouched, _
            BracketEntryTypeStop
        mBOSpec.EntryTriggerPrice = CDbl(lPriceOrOffset)
    Case BracketEntryTypeMarketToLimit, _
            BracketEntryTypeLimit, _
            BracketEntryTypeLimitOnOpen, _
            BracketEntryTypeLimitOnClose, _
            BracketEntryTypeLimitIfTouched, _
            BracketEntryTypeStopLimit
        mBOSpec.EntryPrice = CDbl(lPriceOrOffset)
    Case BracketEntryTypeBid, _
            BracketEntryTypeAsk, _
            BracketEntryTypeLast
        mBOSpec.EntryOffset = parseOffset(lPriceOrOffset)
    Case Else
        processBuyOrSellCommand = False
        gWriteErrorLine "Price must not be specified for this order type"
    End Select
End If
    
Dim lTriggerPrice As String: lTriggerPrice = lClp.Arg(3)
If lTriggerPrice <> "" Then
    Select Case mBOSpec.EntryType
    Case BracketEntryTypeLimitIfTouched, _
            BracketEntryTypeStopLimit
        mBOSpec.EntryTriggerPrice = CDbl(lTriggerPrice)
    Case Else
        processBuyOrSellCommand = False
        gWriteErrorLine "Trigger price must not be specified for this order type"
    End Select
End If

Dim lSwitchName As String
lSwitchName = TIFSwitch
If lClp.Switch(TIFSwitch) Then mBOSpec.EntryTIF = OrderTIFFromString(lClp.SwitchValue(TIFSwitch))

mBracketOrders.Add mBOSpec

gSetValidNextCommands ContractCommand, BuyCommand, SellCommand, BracketCommand, ResetCommand, CloseoutCommand

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrTypeMismatch Then
    gWriteErrorLine lSwitchName & ": invalid value"
    Resume Next
ElseIf Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description
    Resume Next
End If
gHandleUnexpectedError ProcName, ModuleName
End Function




