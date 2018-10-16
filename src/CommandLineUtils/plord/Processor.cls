VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Processor"
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

Private Const ModuleName                            As String = "Processor"

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

Private mOrderSubmitterFactory                      As IOrderSubmitterFactory
Private mPositionManager                            As PositionManager

Private mOrderContext                               As OrderContext

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

If ev.Future.IsCancelled Then
    gWriteErrorLine "Error: contract fetch was cancelled"
    Set mContract = Nothing
    Set mDataSource = Nothing
ElseIf ev.Future.IsFaulted Then
    gWriteErrorLine "Error: " & ev.Future.ErrorMessage
    Set mContract = Nothing
    Set mDataSource = Nothing
ElseIf TypeOf ev.Future.Value Is IContract Then
    Set mContract = ev.Future.Value
    
    If IsContractExpired(mContract) Then
        gWriteErrorLine "Error: contract has expired"
        Set mContract = Nothing
        Set mDataSource = Nothing
    Else
        Dim lOrderManager As New OrderManager
        Set mPositionManager = lOrderManager.CreatePositionManager(mContract.Specifier.Key, mDataSource, mOrderSubmitterFactory)
        Set mOrderContext = mPositionManager.OrderContexts.DefaultOrderContext
        mOrderContext.OrderSubmitter.OrdersAreStaged = mStageOrders
    End If
End If

gSetValidNextCommands BracketCommand, EndOrdersCommand
gInputPaused = False

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
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory)
Set mContractStore = pContractStore
Set mMarketDataManager = pMarketDataManager
Set mOrderSubmitterFactory = pOrderSubmitterFactory
End Sub

Friend Sub PlaceOrder(ByVal pBOS As BracketOrderSpecification)
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

gWriteLineToStdOut "TIME " & FormatTimestamp(GetTimestamp, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
gWriteLineToStdOut pBOS.ToString
    
With pBOS
    Dim lBO As BracketOrder
    If .Action = OrderActionBuy Then
        Set lBO = mOrderContext.Buy(.Quantity, _
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
                          .TargetTriggerPrice, _
                          .EntryTIF, _
                          .StopLossTIF, _
                          .TargetTIF, _
                          .CancelPrice, _
                          .CancelAfter)
    Else
        Set lBO = mOrderContext.Sell(.Quantity, _
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
                          .TargetTriggerPrice, _
                          .EntryTIF, _
                          .StopLossTIF, _
                          .TargetTIF, _
                          .CancelPrice, _
                          .CancelAfter)
    End If
End With

If pBOS.GoodAfterTime <> 0 Then lBO.EntryOrder.GoodAfterTime = pBOS.GoodAfterTime
If pBOS.GoodTillDate <> 0 Then lBO.EntryOrder.GoodTillDate = pBOS.GoodTillDate
lBO.EntryOrder.GoodAfterTimeTZ = pBOS.TimezoneName
lBO.EntryOrder.GoodTillDateTZ = pBOS.TimezoneName

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
If lQuantity = "" Then gWriteErrorLine ("Quantity is missing")

If Not IsInteger(lQuantity, 1) Then gWriteErrorLine ("Quantity must be an integer > 0")
mBOSpec.Quantity = CInt(lQuantity)

Dim lSwitchName As String
lSwitchName = CancelAfterSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.CancelAfter = CInt(lClp.switchValue(lSwitchName))
If mBOSpec.CancelAfter < 0 Then gWriteErrorLine "CancelAfter must be an integer >= 0"

lSwitchName = CancelPriceSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.CancelPrice = CDbl(lClp.switchValue(lSwitchName))

lSwitchName = GoodAfterTimeSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.GoodAfterTime = CDate(lClp.switchValue(lSwitchName))

lSwitchName = GoodTillDateSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.GoodTillDate = CDate(lClp.switchValue(lSwitchName))

lSwitchName = TimezoneSwitch
If lClp.Switch(lSwitchName) Then mBOSpec.TimezoneName = lClp.switchValue(lSwitchName)

gSetValidNextCommands EntryCommand

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

Dim lContractSpec As IContractSpecifier
Set lContractSpec = parseContractSpec(pParams)

If lContractSpec Is Nothing Then
    gSetValidNextCommands ContractCommand, EndOrdersCommand
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
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ProcessEndBracketCommand()
Const ProcName As String = "ProcessEndBracketCommand"
On Error GoTo Err

mBracketOrders.Add mBOSpec

gSetValidNextCommands BracketCommand, ContractCommand, EndOrdersCommand

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
If lClp.Switch(PriceSwitch) Then mBOSpec.EntryPrice = CDbl(lClp.switchValue(PriceSwitch))

lSwitchName = TriggerPriceSwitch
If lClp.Switch(TriggerPriceSwitch) Then
    mBOSpec.EntryTriggerPrice = CDbl(lClp.switchValue(TriggerPriceSwitch))
ElseIf lClp.Switch(TriggerPriceSwitch1) Then
    mBOSpec.EntryTriggerPrice = CDbl(lClp.switchValue(TriggerPriceSwitch1))
End If

lSwitchName = OffsetSwitch
If lClp.Switch(OffsetSwitch) Then mBOSpec.EntryOffset = parseOffset(lClp.switchValue(OffsetSwitch))

lSwitchName = TIFSwitch
If lClp.Switch(TIFSwitch) Then mBOSpec.EntryTIF = OrderTIFFromString(lClp.switchValue(TIFSwitch))

gSetValidNextCommands StopLossCommand, TargetCommand, EndBracketCommand

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
If lClp.Switch(PriceSwitch) Then mBOSpec.StopLossPrice = CDbl(lClp.switchValue(PriceSwitch))

lSwitchName = TriggerPriceSwitch
If lClp.Switch(TriggerPriceSwitch) Then
    mBOSpec.StopLossTriggerPrice = CDbl(lClp.switchValue(TriggerPriceSwitch))
ElseIf lClp.Switch(TriggerPriceSwitch1) Then
    mBOSpec.StopLossTriggerPrice = CDbl(lClp.switchValue(TriggerPriceSwitch1))
End If

lSwitchName = OffsetSwitch
If lClp.Switch(OffsetSwitch) Then mBOSpec.StopLossOffset = parseOffset(lClp.switchValue(OffsetSwitch))

lSwitchName = TIFSwitch
If lClp.Switch(TIFSwitch) Then mBOSpec.StopLossTIF = OrderTIFFromString(lClp.switchValue(TIFSwitch))

If mBOSpec.TargetType = BracketTargetTypeNone Then
    gSetValidNextCommands TargetCommand, EndBracketCommand
Else
    gSetValidNextCommands EndBracketCommand
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
If lClp.Switch(PriceSwitch) Then mBOSpec.TargetPrice = CDbl(lClp.switchValue(PriceSwitch))

lSwitchName = TriggerPriceSwitch
If lClp.Switch(TriggerPriceSwitch) Then
    mBOSpec.TargetTriggerPrice = CDbl(lClp.switchValue(TriggerPriceSwitch))
ElseIf lClp.Switch(TriggerPriceSwitch1) Then
    mBOSpec.TargetTriggerPrice = CDbl(lClp.switchValue(TriggerPriceSwitch1))
End If

lSwitchName = OffsetSwitch
If lClp.Switch(OffsetSwitch) Then mBOSpec.TargetOffset = parseOffset(lClp.switchValue(OffsetSwitch))

lSwitchName = TIFSwitch
If lClp.Switch(TIFSwitch) Then mBOSpec.TargetTIF = OrderTIFFromString(lClp.switchValue(TIFSwitch))

If mBOSpec.StopLossType = BracketStopLossTypeNone Then
    gSetValidNextCommands StopLossCommand, EndBracketCommand
Else
    gSetValidNextCommands EndBracketCommand
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
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine "Invalid target order type: " & pInput
    getTargetOrderType = BracketTargetTypeNone
End If
End Function

Private Function parseContractSpec( _
                ByVal params As String) As IContractSpecifier
Const ProcName As String = "parseContractSpec"
On Error GoTo Err

If Trim$(params) = "" Then
    showContractHelp
    Exit Function
End If

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(params, " ")

If lClp.Arg(0) = "?" Or _
    lClp.Switch("?") Or _
    lClp.NumberOfSwitches = 0 _
Then
    showContractHelp
    Exit Function
End If

Dim validParams As Boolean
validParams = True

Dim lSectypeStr As String: lSectypeStr = lClp.switchValue(SecTypeSwitch)
If lSectypeStr = "" Then lSectypeStr = lClp.switchValue(SecTypeSwitch1)

Dim lExchange As String: lExchange = lClp.switchValue(ExchangeSwitch)
If lExchange = "" Then lExchange = lClp.switchValue(ExchangeSwitch1)

Dim lLocalSymbol As String: lLocalSymbol = lClp.switchValue(LocalSymbolSwitch)
If lLocalSymbol = "" Then lLocalSymbol = lClp.switchValue(LocalSymbolSwitch1)

Dim lSymbol As String: lSymbol = lClp.switchValue(SymbolSwitch)
If lSymbol = "" Then lSymbol = lClp.switchValue(SymbolSwitch1)

Dim lCurrency As String: lCurrency = lClp.switchValue(CurrencySwitch)
If lCurrency = "" Then lCurrency = lClp.switchValue(CurrencySwitch1)

Dim lExpiry As String: lExpiry = lClp.switchValue(ExpirySwitch)
If lExpiry = "" Then lExpiry = lClp.switchValue(ExpirySwitch1)

Dim lMultiplier As String: lMultiplier = lClp.switchValue(MultiplierSwitch)
If lMultiplier = "" Then lMultiplier = lClp.switchValue(MultiplierSwitch1)
If lMultiplier = "" Then lMultiplier = "1.0"

Dim lStrike As String: lStrike = lClp.switchValue(StrikeSwitch)
If lStrike = "" Then lStrike = lClp.switchValue(StrikeSwitch1)
If lStrike = "" Then lStrike = "0.0"

Dim lRight As String: lRight = lClp.switchValue(RightSwitch)

Dim lSectype As SecurityTypes
lSectype = SecTypeFromString(lSectypeStr)
If lSectypeStr <> "" And lSectype = SecTypeNone Then
    gWriteErrorLine "Invalid Sectype '" & lSectypeStr & "'"
    validParams = False
End If

If lExpiry <> "" Then
    If IsDate(lExpiry) Then
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




