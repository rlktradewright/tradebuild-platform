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

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mContractStore                              As IContractStore
Private mMarketDataManager                          As IMarketDataManager

Private mContract                                   As IContract
Private mDataSource                                 As IMarketDataSource

Private mStageOrders                                As Boolean

Private mBracketOrders                              As New EnumerableCollection

Private mBOSpec                                     As BracketOrderSpecification

Private mOrderManager                               As OrderManager

Private mOrderSubmitterFactory                      As IOrderSubmitterFactory
Private mPositionManager                            As PositionManager

Private mOrderContext                               As OrderContext

Private mScopeName                                  As String
Private mGroupName                                  As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

gInputPaused = False

If ev.Future.IsCancelled Then
    gWriteErrorLine "contract fetch was cancelled"
ElseIf ev.Future.IsFaulted Then
    gWriteErrorLine ev.Future.ErrorMessage
ElseIf TypeOf ev.Future.Value Is IContract Then
    Set mContract = ev.Future.Value
    
    If IsContractExpired(mContract) Then
        gWriteErrorLine "contract has expired"
    Else
        If mScopeName = "" Then
            Set mPositionManager = mOrderManager.CreatePositionManager(mContract.Specifier.Key, mDataSource, mOrderSubmitterFactory)
        Else
            Set mPositionManager = mOrderManager.CreateRecoverablePositionManager(mGroupName & "$$" & mContract.Specifier.Key, mDataSource, mOrderSubmitterFactory, mScopeName, mGroupName)
        End If
        Set mOrderContext = mPositionManager.OrderContexts.DefaultOrderContext
        mOrderContext.OrderSubmitter.OrdersAreStaged = mStageOrders
        gSetValidNextCommands ListCommand, GroupCommand, ContractCommand, BracketCommand, EndOrdersCommand, ResetCommand, CloseoutCommand
    End If
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

Friend Property Let StageOrders(ByVal Value As Boolean)
mStageOrders = Value
End Property

Friend Property Get StageOrders() As Boolean
StageOrders = mStageOrders
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractStore As IContractStore, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pOrderManager As OrderManager, _
                ByVal pScopeName As String, _
                ByVal pGroupName As String, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory)
Set mContractStore = pContractStore
Set mMarketDataManager = pMarketDataManager
Set mOrderManager = pOrderManager
mScopeName = pScopeName
mGroupName = pGroupName
Set mOrderSubmitterFactory = pOrderSubmitterFactory
End Sub

Friend Sub PlaceOrder(ByVal pBOS As BracketOrderSpecification)
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

gWriteLineToStdOut "TIME " & FormatTimestamp(GetTimestamp, TimestampDateAndTimeISO8601)
gWriteLineToStdOut pBOS.ToString
    
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

Friend Function processContractCommand( _
                ByVal pParams As String) As Boolean
Const ProcName As String = "ProcessContractCommand"
On Error GoTo Err

If Trim$(pParams) = "" Then Exit Function

If Trim$(pParams) = HelpCommand Or Trim$(pParams) = Help1Command Then
    showContractHelp
    Exit Function
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
    gSetValidNextCommands ListCommand, ContractCommand, EndOrdersCommand, ResetCommand
    Exit Function
End If

Dim lContractFuture As IFuture
Set lContractFuture = FetchContract(lContractSpec, mContractStore)
mFutureWaiter.Add lContractFuture
Set mDataSource = mMarketDataManager.CreateMarketDataSource(lContractFuture, False)
mDataSource.StartMarketData

processContractCommand = True

Exit Function

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description
    processContractCommand = False
    Exit Function
End If

gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ProcessEndBracketCommand()
Const ProcName As String = "ProcessEndBracketCommand"
On Error GoTo Err

mBracketOrders.Add mBOSpec

gSetValidNextCommands ListCommand, GroupCommand, BracketCommand, ContractCommand, EndOrdersCommand, ResetCommand

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




