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

If Not isReadyForOrder Then Exit Sub

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
End Sub

Friend Sub Finish()
Set mOrderContext = Nothing
Set mOrderContextEv = Nothing
End Sub

Friend Function PlaceOrder() As Boolean
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

Static sAttempts As Long

If sAttempts = 0 Then
    gWriteLineToStdOut "TIME " & FormatTimestamp(mTimezone.ConvertDateLocalToTz(GetTimestamp), TimestampDateAndTimeISO8601)
    writeContractSpecToStdOut
    gWriteLineToStdOut mBracketOrderSpec.ToString
End If
sAttempts = sAttempts + 1

If Not isReadyForOrder Then
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
    Dim lEntryOrder As IOrder
    Set lEntryOrder = mOrderContext.CreateEntryOrder( _
                            .EntryType, _
                            .EntryLimitPriceSpec, _
                            .EntryTriggerPriceSpec, _
                            .EntryIgnoreRTH)
    
    Dim lStopLossOrder As IOrder
    If .StopLossType <> OrderTypeNone Then
        Set lStopLossOrder = mOrderContext.CreateStopLossOrder( _
                                .StopLossType, _
                                .StopLossLimitPriceSpec, _
                                .StopLossTriggerPriceSpec, _
                                .StopLossIgnoreRTH)
    End If

    Dim lTargetOrder As IOrder
    If .TargetType <> OrderTypeNone Then
        Set lTargetOrder = mOrderContext.CreateTargetOrder( _
                                .TargetType, _
                                .TargetLimitPriceSpec, _
                                .TargetTriggerPriceSpec, _
                                .TargetIgnoreRTH)
    End If


    Set lBo = mOrderContext.CreateBracketOrder( _
                    .Action, _
                    .Quantity, _
                    lEntryOrder, _
                    lStopLossOrder, _
                    lTargetOrder _
                    )
    
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
    If .Rollover Then
        lBo.RolloverSpecification = CreateRolloverSpecification( _
                                            .RolloverDays, _
                                            .RolloverTime, _
                                            .RolloverCloseType, _
                                            .RolloverCloseLimitPriceSpec, _
                                            .RolloverCloseTriggerPriceSpec, _
                                            .RolloverEntryType, _
                                            .RolloverEntryLimitPriceSpec, _
                                            .RolloverEntryTriggerPriceSpec)
    End If
End With

gWriteLineToStdOut lBo.Key & ": bracket order placed"

mOrderContext.ExecuteBracketOrder lBo
CreateBracketProfitCalculator lBo, mOrderContext.DataSource
gBracketOrderListener.Add lBo

gNumberOfOrdersPlaced = gNumberOfOrdersPlaced + 1

PlaceOrder = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function isReadyForOrder() As Boolean
With mOrderContext
    If Not .IsReady Then Exit Function
    
    If Not .IsReadyTickSize(mContract.Specifier.SecType) Then Exit Function
    
    If Not .IsReadyForPriceSpecifier(mBracketOrderSpec.EntryLimitPriceSpec) Then Exit Function
    If Not .IsReadyForPriceSpecifier(mBracketOrderSpec.EntryTriggerPriceSpec) Then Exit Function
    If Not .IsReadyForPriceSpecifier(mBracketOrderSpec.StopLossLimitPriceSpec) Then Exit Function
    If Not .IsReadyForPriceSpecifier(mBracketOrderSpec.StopLossTriggerPriceSpec) Then Exit Function
    If Not .IsReadyForPriceSpecifier(mBracketOrderSpec.TargetLimitPriceSpec) Then Exit Function
    If Not .IsReadyForPriceSpecifier(mBracketOrderSpec.TargetTriggerPriceSpec) Then Exit Function
End With
isReadyForOrder = True
End Function

Private Sub writeContractSpecToStdOut()
Const ProcName As String = "writeContractSpecToStdOut"
On Error GoTo Err

Dim s As String
s = gCommands.ContractCommand.Name & " "
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




