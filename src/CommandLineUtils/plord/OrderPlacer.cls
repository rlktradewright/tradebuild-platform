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

Private mMoneyManager                               As MoneyManager

Private mGroupResources                             As GroupResources

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

Dim lChangeType As OrderContextChangeTypes: lChangeType = ev.ChangeType

Select Case lChangeType
Case OrderContextChangeTypes.OrderContextActiveStateChanged
    Exit Sub
Case OrderContextChangeTypes.OrderContextReadyStateChanged
    Static sPendingMessageLogged As Boolean
    If Not isReadyForOrder Then
        If Not sPendingMessageLogged Then
            sPendingMessageLogged = True
            LogMessage mBracketOrderSpec.ID & ": Order is pending: " & mOrderContextEv.NotReadyMessage
        End If
        Exit Sub
    End If
    
    Set mOrderContextEv = Nothing
    gPlaceOrdersTask.AddOrderPlacer Me
    
Case OrderContextChangeTypes.OrderContextFinished
    Cancel mOrderContextEv.NotReadyMessage
End Select


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

Friend Sub Cancel( _
                ByVal pReason As String)
If mGroupResources.OrderPlacers.Contains(Me) Then mGroupResources.OrderPlacers.Remove Me
Set mGroupResources = Nothing
gPlaceOrdersTask.RemoveOrderPlacer Me
Set mOrderContext = Nothing
Set mOrderContextEv = Nothing

gWriteLineToConsole mBracketOrderSpec.ID & ": Order cancelled: " & pReason
End Sub

Friend Sub Initialise( _
                ByVal pBracketOrderSpec As BracketOrderSpecification, _
                ByVal pcontract As IContract, _
                ByVal pOrderContext As OrderContext, _
                ByVal pStageOrders As Boolean, _
                ByVal pMoneyManager As MoneyManager, _
                ByVal pGroupResources As GroupResources)
Set mBracketOrderSpec = pBracketOrderSpec
Set mContract = pcontract
Set mOrderContext = pOrderContext
mStageOrders = pStageOrders
Set mMoneyManager = pMoneyManager
Set mGroupResources = pGroupResources

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

Dim s As String
If sAttempts = 0 Then
    s = contractSpecToString & vbCrLf & mBracketOrderSpec.ToString
    LogMessage s
End If
sAttempts = sAttempts + 1

If Not isReadyForOrder Then
    Set mOrderContextEv = mOrderContext
    LogMessage mBracketOrderSpec.ID & ": Order is pending: " & mOrderContext.NotReadyMessage
    Exit Function
End If

If sAttempts > 1 Then LogMessage mBracketOrderSpec.ID & ": Provider is now ready"
    
mOrderContext.OrderSubmitter.OrdersAreStaged = mStageOrders

Dim lBO As BracketOrder

If mBracketOrderSpec.Modify Then
    Dim lMsg As String
    Set lBO = modifyOrder(lMsg)
    If lBO Is Nothing Then
        gWriteLineToConsole mBracketOrderSpec.ID & ": " & lMsg
        PlaceOrder = False
    End If
Else
    Set lBO = placeNewOrder
    If lBO.State = BracketOrderStateSubmitted Then CreateBracketProfitCalculator lBO, mOrderContext.DataSource
    gLiveOrders.Item(lBO.Key).Order = lBO
    PlaceOrder = True
End If

mGroupResources.OrderPlacers.Remove mBracketOrderSpec.ID
Set mGroupResources = Nothing
gPlaceOrdersTask.RemoveOrderPlacer Me

Exit Function

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description, ErrorCountIncrementNo
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function isReadyForOrder() As Boolean
isReadyForOrder = False
With mOrderContext
    If Not .IsReady Then
    ElseIf Not .IsReadyTickSize() Then
    ElseIf Not .IsReadyForAutoSize(mBracketOrderSpec.AutosizeType, _
                                    mBracketOrderSpec.Contract.Specifier.CurrencyCode) Then
    ElseIf Not .IsReadyForOrderType(mBracketOrderSpec.EntryType) Then
    ElseIf Not .IsReadyForPriceSpecifier(mBracketOrderSpec.EntryLimitPriceSpec) Then
    ElseIf Not .IsReadyForPriceSpecifier(mBracketOrderSpec.EntryTriggerPriceSpec) Then
    ElseIf Not .IsReadyForOrderType(mBracketOrderSpec.StopLossType) Then
    ElseIf Not .IsReadyForPriceSpecifier(mBracketOrderSpec.StopLossLimitPriceSpec) Then
    ElseIf Not .IsReadyForPriceSpecifier(mBracketOrderSpec.StopLossTriggerPriceSpec) Then
    ElseIf Not .IsReadyForOrderType(mBracketOrderSpec.TargetType) Then
    ElseIf Not .IsReadyForPriceSpecifier(mBracketOrderSpec.TargetLimitPriceSpec) Then
    ElseIf Not .IsReadyForPriceSpecifier(mBracketOrderSpec.TargetTriggerPriceSpec) Then
    Else
        isReadyForOrder = True
    End If
End With
End Function

Private Function contractSpecToString() As String
Const ProcName As String = "contractSpecToString"
On Error GoTo Err

Dim s As String
s = gCommands.ContractCommand.Name & " "
s = s & mContract.Specifier.ToString
contractSpecToString = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function modifyOrder(ByRef pMessage As String) As BracketOrder
Const ProcName As String = "modifyOrder"
On Error GoTo Err

Dim s As String
s = mBracketOrderSpec.ID & ": Modifying existing order"
gWriteLineToConsole s, True, True

Dim lBO As BracketOrder
Dim lEntry As LiveOrderEntry
If Not gLiveOrders.TryItem(mBracketOrderSpec.ID, lEntry) Then
    pMessage = "Order " & mBracketOrderSpec.ID & " not found"
    Exit Function
End If

Set lBO = lEntry.Order

With mBracketOrderSpec
    If lBO.EntryOrder.Status <> OrderStatusFilled And _
        lBO.EntryOrder.Status <> OrderStatusCancelling And _
        lBO.EntryOrder.Status <> OrderStatusCancelled _
    Then
        If lBO.EntryOrder.OrderType <> .EntryType Then
            pMessage = "Can't change the entry order type"
            Exit Function
        End If
        lBO.SetNewEntryQuantity .Quantity
        If .EntryLimitPriceSpec.IsValid Then lBO.SetNewEntryLimitPrice .EntryLimitPriceSpec
        If .EntryTriggerPriceSpec.IsValid Then lBO.SetNewEntryTriggerPrice .EntryTriggerPriceSpec
    End If
    
    If lBO.StopLossOrder Is Nothing Then
        If .StopLossType <> OrderTypeNone Or _
            .StopLossLimitPriceSpec.IsValid Or _
            .StopLossTriggerPriceSpec.IsValid _
        Then
            pMessage = "Can't add a stop-loss order"
            Exit Function
        End If
    ElseIf lBO.StopLossOrder.Status <> OrderStatusCreated And _
            lBO.StopLossOrder.Status <> OrderStatusSubmitted And _
            lBO.StopLossOrder.Status <> OrderStatusCancelling _
    Then
        ' ignore if not suitable state
    ElseIf .StopLossType <> lBO.StopLossOrder.OrderType Then
        pMessage = "Can't change stop-loss order type"
        Exit Function
    Else
        lBO.SetNewStopLossQuantity .Quantity
        If .StopLossLimitPriceSpec.IsValid Then lBO.SetNewStopLossLimitPrice .StopLossLimitPriceSpec
        If .StopLossTriggerPriceSpec.IsValid Then lBO.SetNewStopLossTriggerPrice .StopLossTriggerPriceSpec
    End If
        
    If lBO.TargetOrder Is Nothing Then
        If .TargetType <> OrderTypeNone Or _
            .TargetLimitPriceSpec.IsValid Or _
            .TargetTriggerPriceSpec.IsValid _
        Then
            pMessage = "Can't add a target order"
            Exit Function
        End If
    ElseIf lBO.TargetOrder.Status <> OrderStatusCreated And _
            lBO.TargetOrder.Status <> OrderStatusSubmitted And _
            lBO.TargetOrder.Status <> OrderStatusCancelling _
    Then
        ' ignore if not suitable state
    ElseIf .TargetType <> lBO.TargetOrder.OrderType Then
        pMessage = "Can't change stop-loss order type"
        Exit Function
    Else
        lBO.SetNewTargetQuantity .Quantity
        If .TargetLimitPriceSpec.IsValid Then lBO.SetNewTargetLimitPrice .TargetLimitPriceSpec
        If .TargetTriggerPriceSpec.IsValid Then lBO.SetNewTargetTriggerPrice .TargetTriggerPriceSpec
    End If

    On Error Resume Next
    lBO.CancelAfter = .CancelAfter
    lBO.CancelPrice = .CancelPrice
    lBO.Description = .Description
    lBO.EntryOrder.TimeInForce = .EntryTIF
    lBO.EntryReason = .EntryReason
    If Not lBO.StopLossOrder Is Nothing Then
        lBO.StopLossOrder.TimeInForce = .StopLossTIF
        lBO.StopReason = .StopLossReason
    End If
    If Not lBO.TargetOrder Is Nothing Then
        lBO.TargetOrder.TimeInForce = .TargetTIF
        lBO.TargetReason = .TargetReason
    End If
    If .GoodAfterTime <> 0 Then
        lBO.EntryOrder.GoodAfterTime = .GoodAfterTime
        lBO.EntryOrder.GoodAfterTimeTZ = .TimezoneName
    End If
    If .GoodTillDate <> 0 Then
        lBO.EntryOrder.GoodTillDate = .GoodTillDate
        lBO.EntryOrder.GoodTillDateTZ = .TimezoneName
    End If
    
    If Not .RolloverSpec Is Nothing Then lBO.RolloverSpecification = .RolloverSpec
    
    If .AutosizeType <> BracketAutosizeNone Then
        lBO.AutosizeType = .AutosizeType
        lBO.AutosizeValue = .AutosizeValue
        lBO.LotSize = .LotSize
    End If
    
    On Error GoTo Err
End With

LogMessage mBracketOrderSpec.ID & ": order ready to be updated"

Set modifyOrder = lBO
lBO.Update

Exit Function

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    pMessage = Err.Description
    Exit Function
End If
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function placeNewOrder() As BracketOrder
Const ProcName As String = "placeNewOrder"
On Error GoTo Err

Dim s As String
s = mBracketOrderSpec.ID & ": Placing new order"
gWriteLineToConsole s, True, True

Dim lBO As BracketOrder

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


    Set lBO = mOrderContext.CreateBracketOrderWithId( _
                    .Action, _
                    .Quantity, _
                    lEntryOrder, _
                    .ID, _
                    lStopLossOrder, _
                    lTargetOrder _
                    )
    
    gBracketOrderListener.Add lBO
    
    lBO.CancelAfter = .CancelAfter
    lBO.CancelPrice = .CancelPrice
    lBO.Description = .Description
    lBO.EntryOrder.TimeInForce = .EntryTIF
    lBO.EntryReason = .EntryReason
    If Not lBO.StopLossOrder Is Nothing Then
        lBO.StopLossOrder.TimeInForce = .StopLossTIF
        lBO.StopReason = .StopLossReason
    End If
    If Not lBO.TargetOrder Is Nothing Then
        lBO.TargetOrder.TimeInForce = .TargetTIF
        lBO.TargetReason = .TargetReason
    End If
    If .GoodAfterTime <> 0 Then
        lBO.EntryOrder.GoodAfterTime = .GoodAfterTime
        lBO.EntryOrder.GoodAfterTimeTZ = .TimezoneName
    End If
    If .GoodTillDate <> 0 Then
        lBO.EntryOrder.GoodTillDate = .GoodTillDate
        lBO.EntryOrder.GoodTillDateTZ = .TimezoneName
    End If
    
    If Not .RolloverSpec Is Nothing Then lBO.RolloverSpecification = .RolloverSpec
    
    If .AutosizeType <> BracketAutosizeNone Then
        lBO.AutosizeType = .AutosizeType
        lBO.AutosizeValue = .AutosizeValue
        lBO.LotSize = .LotSize
    End If
End With

LogMessage mBracketOrderSpec.ID & ": order ready to be placed"

mOrderContext.ExecuteBracketOrder lBO, True

Set placeNewOrder = lBO

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


