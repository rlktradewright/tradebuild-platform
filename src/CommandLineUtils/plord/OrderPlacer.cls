VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPlacer"
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

Private Const ModuleName                            As String = "OrderPlacer"

'@================================================================================
' Member variables
'@================================================================================

Private mBracketOrderSpec                           As BracketOrderSpecification
Private mContract                                   As IContract
Private mOrderContext                               As OrderContext
Attribute mOrderContext.VB_VarHelpID = -1
Private WithEvents mOrderContextEv                  As OrderContext
Attribute mOrderContextEv.VB_VarHelpID = -1
Private mStageOrders                                As Boolean

Private mTimezone                                   As TimeZone

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mOrderContext Event Handlers
'@================================================================================

Private Sub mOrderContextEv_Change(ev As ChangeEventData)
Const ProcName As String = "mOrderContextEv_Change"
On Error GoTo Err

If ev.ChangeType <> OrderContextChangeTypes.OrderContextReadyStateChanged Then Exit Sub

If Not mOrderContext.IsReadyForBracketOrder(mBracketOrderSpec.EntryType, _
                    mBracketOrderSpec.StopLossType, _
                    mBracketOrderSpec.TargetType) Then
    Exit Sub
End If

Set mOrderContextEv = Nothing
gPlaceOrdersTask.AddOrderPlacer Me

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pBracketOrderSpec As BracketOrderSpecification, _
                ByVal pContract As IContract, _
                ByVal pOrderContext As OrderContext, _
                ByVal pStageOrders As Boolean)
Set mBracketOrderSpec = pBracketOrderSpec
Set mContract = pContract
Set mOrderContext = pOrderContext
mStageOrders = pStageOrders

Set mTimezone = GetTimeZone(mContract.TimezoneName)

Set mSelfRef = Me
End Sub

Friend Function PlaceOrder() As Boolean
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

Static sAttempts As Long

If sAttempts = 0 Then
    sAttempts = sAttempts + 1
    
    gWriteLineToStdOut "TIME " & FormatTimestamp(mTimezone.ConvertDateLocalToTz(GetTimestamp), TimestampDateAndTimeISO8601)
    writeContractSpecToStdOut
    gWriteLineToStdOut mBracketOrderSpec.ToString
End If

If Not mOrderContext.IsReadyForBracketOrder(mBracketOrderSpec.EntryType, _
                                            mBracketOrderSpec.StopLossType, _
                                            mBracketOrderSpec.TargetType) Then
    Set mOrderContextEv = mOrderContext
    gWriteLineToStdOut "Order is pending, waiting for the provider to be ready"
    Exit Function
End If

If sAttempts > 1 Then
    gWriteLineToStdOut "Provider is now ready"
    gWriteLineToStdOut "TIME " & FormatTimestamp(mTimezone.ConvertDateLocalToTz(GetTimestamp), TimestampDateAndTimeISO8601)
    writeContractSpecToStdOut
    gWriteLineToStdOut mBracketOrderSpec.ToString
End If

mOrderContext.OrderSubmitter.OrdersAreStaged = mStageOrders

Dim lBo As IBracketOrder
With mBracketOrderSpec
    If .Action = OrderActionBuy Then
        Set lBo = mOrderContext.CreateBuyBracketOrder(.Quantity, _
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
        Set lBo = mOrderContext.CreateSellBracketOrder(.Quantity, _
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
    
    lBo.CancelAfter = .CancelAfter
    lBo.CancelPrice = .CancelPrice
    lBo.Description = .Description
    lBo.EntryOrder.TimeInForce = .EntryTIF
    lBo.EntryReason = .EntryReason
    If Not lBo.StopLossOrder Is Nothing Then
        lBo.StopLossOrder.TimeInForce = .StopLossTIF
        lBo.StopReason = .StopLossReason
    End If
    If Not lBo.TargetOrder Is Nothing Then
        lBo.TargetOrder.TimeInForce = .TargetTIF
        lBo.TargetReason = .TargetReason
    End If
    If .GoodAfterTime <> 0 Then
        lBo.EntryOrder.GoodAfterTime = .GoodAfterTime
        lBo.EntryOrder.GoodAfterTimeTZ = .TimezoneName
    End If
    If .GoodTillDate <> 0 Then
        lBo.EntryOrder.GoodTillDate = .GoodTillDate
        lBo.EntryOrder.GoodTillDateTZ = .TimezoneName
    End If
End With

gWriteLineToStdOut lBo.Key & ": bracket order placed"

mOrderContext.ExecuteBracketOrder lBo
CreateBracketProfitCalculator lBo, mOrderContext.DataSource
gBracketOrderListener.Add lBo

gNumberOfOrdersPlaced = gNumberOfOrdersPlaced + 1

Set mSelfRef = Nothing

PlaceOrder = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub writeContractSpecToStdOut()
Const ProcName As String = "writeContractSpecToStdOut"
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




