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
Attribute VB_Creatable = False
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

Implements IOrderInfoConsumer

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

Private Const OrderRecoveryExecutionsRequestId      As Long = 777

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTwsAPI                          As TwsAPI
Attribute mTwsAPI.VB_VarHelpID = -1

Private mKnownOrders                                As New KnownOrders

Private mOrderRecoveryListener                      As IOrderRecoveryListener

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IOrderInfoConsumer Interface Members
'@================================================================================

Private Sub IOrderInfoConsumer_EndExecutions(ByVal pRequestId As Long)
Const ProcName As String = "IOrderInfoConsumer_EndExecutions"
On Error GoTo Err

If Not mOrderRecoveryListener Is Nothing Then mOrderRecoveryListener.ExecutionReportsComplete

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderInfoConsumer_EndOpenOrders()
Const ProcName As String = "IOrderInfoConsumer_EndOpenOrders"
On Error GoTo Err

If Not mOrderRecoveryListener Is Nothing Then mOrderRecoveryListener.OrderReportsComplete

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderInfoConsumer_NotifyError(ByVal pOrderId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IOrderInfoConsumer_NotifyError"
On Error GoTo Err

If Not mKnownOrders.Contains(pOrderId, mTwsAPI.ClientID) Then Exit Sub

Select Case pErrorCode
Case 202
    Dim osp As New OrderStatusReport
    osp.Initialise GetOrder(pOrderId).id, OrderStatuses.OrderStatusCancelled
    GetOrderSubmitter(pOrderId).NotifyOrderStatusReport osp
Case Else
    GetOrderSubmitter(pOrderId).NotifyError GetOrder(pOrderId).BrokerId, pErrorCode, pErrorMsg
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderInfoConsumer_NotifyExecution(ByVal pReqId As Long, ByVal pContract As TwsContract, ByVal pExecution As TwsExecution)
Const ProcName As String = "IOrderInfoConsumer_NotifyExecution"
On Error GoTo Err

Static sExecutions As New EnumerableCollection
Dim lExecReport As New ExecutionReport
Dim lOrder As IOrder
Dim lorderId As String

If sExecutions.Contains(pExecution.ExecId) Then Exit Sub
sExecutions.Add pExecution, pExecution.ExecId

If isKnownOrder(pExecution.OrderId) Then
    Set lOrder = GetOrder(pExecution.OrderId)
    lorderId = lOrder.id
End If

lExecReport.Initialise pContract.LocalSymbol & "." & pContract.Exchange, _
                        pExecution.AcctNumber, _
                        gOrderActionFromString(pExecution.Side), _
                        pExecution.ExecId, _
                        pExecution.Exchange, _
                        pExecution.Liquidation, _
                        pExecution.OrderId, _
                        lorderId, _
                        pExecution.Price, _
                        pExecution.Shares, _
                        pExecution.Time, _
                        GetTimeZone("").DisplayName

If Not lOrder Is Nothing Then
    Dim lOrderSubmitter As OrderSubmitter
    Dim qtyRemaining As Long
    Dim statusRpt As New OrderStatusReport
    
    Set lOrder = GetOrder(pExecution.OrderId)
    Set lOrderSubmitter = GetOrderSubmitter(pExecution.OrderId)
    
    qtyRemaining = lOrder.QuantityRemaining
    
    lOrderSubmitter.NotifyExecutionReport lExecReport
    
    If pReqId <> -1 Then
        ' notify a status update - this is in case this execution occurred during a period when
        ' Tws had lost its connection to the IB servers and this exec details is being notified
        ' in response to the reqExecutions call made when the connection is recovered. In these
        ' circumstances Tws will not itself issue an orderStatus.
        statusRpt.Initialise lOrder.id, IIf(pExecution.Shares >= qtyRemaining, OrderStatuses.OrderStatusFilled, OrderStatuses.OrderStatusSubmitted)
        lOrderSubmitter.NotifyOrderStatusReport statusRpt
    End If
Else
    ' this must be an execution from a previous session
    If Not mOrderRecoveryListener Is Nothing Then mOrderRecoveryListener.NotifyExecutionReport lExecReport
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderInfoConsumer_NotifyOpenOrder(ByVal pOrderId As Long, ByVal pContract As TwsContract, ByVal pOrder As TwsOrder, ByVal pOrderState As TwsOrderState)
Const ProcName As String = "IOrderInfoConsumer_NotifyOpenOrder"
On Error GoTo Err

Dim lOrder As IOrder
Dim lOrderReport As IOrderReport
Dim lorderId As String
Dim lTimezoneName As String

If isKnownOrder(pOrderId) Then
    Set lOrder = GetOrder(pOrderId)
    lorderId = lOrder.id
End If

Set lOrderReport = New OrderReport

With lOrderReport
    .Action = gTwsOrderActionToOrderAction(pOrder.Action)
    .AllOrNone = pOrder.AllOrNone
    .BlockOrder = pOrder.BlockOrder
    .BrokerId = pOrder.OrderId
    .DiscretionaryAmount = pOrder.DiscretionaryAmt
    .DisplaySize = pOrder.DisplaySize
    .ETradeOnly = pOrder.ETradeOnly
    .FirmQuoteOnly = pOrder.FirmQuoteOnly
    .GoodAfterTime = gGetDate(pOrder.GoodAfterTime, lTimezoneName)
    .GoodAfterTimeTZ = lTimezoneName
    .GoodTillDate = gGetDate(pOrder.GoodTillDate, lTimezoneName)
    .GoodTillDateTZ = lTimezoneName
    .Hidden = pOrder.Hidden
    .id = lorderId
    .IgnoreRegularTradingHours = pOrder.OutsideRth
    .LimitPrice = pOrder.LmtPrice
    .MinimumQuantity = pOrder.MinQty
    .NbboPriceCap = pOrder.NbboPriceCap
    .OrderType = gTwsOrderTypeToOrderType(pOrder.OrderType)
    .Origin = pOrder.Origin
    .OriginatorRef = pOrder.OrderRef
    .OverrideConstraints = pOrder.OverridePercentageConstraints
    .Quantity = pOrder.TotalQuantity
    .SettlingFirm = pOrder.SettlingFirm
    .StopTriggerMethod = pOrder.TriggerMethod
    .SweepToFill = pOrder.SweepToFill
    .TimeInForce = pOrder.Tif
    .TriggerPrice = pOrder.AuxPrice
End With

If isKnownOrder(pOrderId) Then
    GetOrderSubmitter(pOrderId).NotifyOrderReport lOrderReport
Else
    ' this must be an open Order from the previous API session, or this is
    ' a duplicate openOrder message for an Order that has already finished
    If Not mOrderRecoveryListener Is Nothing Then mOrderRecoveryListener.NotifyOrderReport lOrderReport
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderInfoConsumer_NotifyOrderStatus( _
                ByVal pOrderId As Long, _
                ByVal pStatus As String, _
                ByVal pFilled As Long, _
                ByVal pRemaining As Long, _
                ByVal pAvgFillPrice As Double, _
                ByVal pPermId As Long, _
                ByVal pParentId As Long, _
                ByVal pLastFillPrice As Double, _
                ByVal pClientId As Long, _
                ByVal pWhyHeld As String)
Const ProcName As String = "IOrderInfoConsumer_NotifyOrderStatus"
On Error GoTo Err

Dim lStatusReport As New OrderStatusReport
Dim lStatus As OrderStatuses

If isKnownOrder(pOrderId) Then
    
    lStatus = gOrderStatusFromString(pStatus)
    lStatusReport.Initialise GetOrder(pOrderId).id, lStatus
    GetOrderSubmitter(pOrderId).NotifyOrderStatusReport lStatusReport
    If lStatus = OrderStatusCancelling And pParentId <> 0 Then
        ' for some reason IB never sends a cancelled status when a target Order
        ' is cancelled after the entry Order has been filled. It just sits there at
        ' 'cancelling' status until Tws is closed (note that it doesn't reappear if
        ' Tws is restarted, so it must have been cancelled).
        ' So we'll treat it as cancelled if we haven't received a fill in say 500 millisecs
        generateDeferredOrderCancelNotification pOrderId
    End If
    
Else
    ' this must be an open Order from the previous API session, but
    ' for which we haven't yet had the openOrder notification
    ' don't notify the service consumer because it won't know about this
    ' Order yet
    '
    ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! need to improve this!!!!!!!!!!!!!!!
End If


Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub CancelOrder(ByVal pOrder As IOrder)
Const ProcName As String = "CancelOrder"
On Error GoTo Err

If pOrder.Status = OrderStatusPendingSubmit Then
    ' if an Order is cancelled immediately after being placed, and before
    ' an orderStatus or openOrder notification has been received from Tws,
    ' Tws sometimes just cancels it without ever sending either of these,
    ' so we need to generate one as a backstop
    generateDeferredOrderCancelNotification pOrder.BrokerId
End If

mTwsAPI.CancelOrder pOrder.BrokerId
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pClient As TwsAPI)
Set mTwsAPI = pClient
End Sub

Friend Sub FinishedOrderRecovery()
Set mOrderRecoveryListener = Nothing
End Sub

Friend Sub PlaceOrder( _
                ByVal pOrderSubmitter As OrderSubmitter, _
                ByVal pOrder As IOrder, _
                ByVal pContract As TwsContract, _
                ByVal pTransmit As Boolean, _
                ByVal pParentId As String, _
                ByVal pOcaGroup As String)
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

Dim mapEntry As OrderIdMapEntry
Dim lOrder As TwsOrder

mKnownOrders.Add pOrder, pOrderSubmitter, mTwsAPI.ClientID

Set lOrder = gOrderToTwsOrder(pOrder)
lOrder.Transmit = pTransmit
lOrder.ParentId = pParentId
lOrder.OcaGroup = pOcaGroup

mTwsAPI.PlaceOrder lOrder, pContract

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub
                    
Friend Sub StartOrderRecovery( _
                ByVal pOrderRecoveryListener As IOrderRecoveryListener, _
                ByVal pStartFrom As Date)
Const ProcName As String = "StartOrderRecovery"
On Error GoTo Err

Dim lExecFilter As New TwsExecutionFilter

Set mOrderRecoveryListener = pOrderRecoveryListener
mTwsAPI.RequestOpenOrders

lExecFilter.Time = pStartFrom
mTwsAPI.RequestExecutions OrderRecoveryExecutionsRequestId, lExecFilter

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub generateDeferredOrderCancelNotification( _
                ByVal pBrokerId As Long)
Const ProcName As String = "generateDeferredOrderCancelNotification"
On Error GoTo Err

Dim lMapEntry As OrderIdMapEntry

If Not isKnownOrder(pBrokerId) Then Exit Sub

Set lMapEntry.Order = mKnownOrders.GetOrder(pBrokerId, mTwsAPI.ClientID)
Set lMapEntry.OrderSubmitter = mKnownOrders.GetOrderSubmitter(pBrokerId, mTwsAPI.ClientID)

DeferAction New DeferredCancelStatus, lMapEntry, 500

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function GetOrder(ByVal pBrokerId) As IOrder
Const ProcName As String = "getOrder"
On Error GoTo Err

Set GetOrder = mKnownOrders.GetOrder(pBrokerId, mTwsAPI.ClientID)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function GetOrderSubmitter(ByVal pBrokerId) As OrderSubmitter
Const ProcName As String = "getOrderSubmitter"
On Error GoTo Err

Set GetOrderSubmitter = mKnownOrders.GetOrderSubmitter(pBrokerId, mTwsAPI.ClientID)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function isKnownOrder(ByVal pBrokerId) As Boolean
Const ProcName As String = "isKnownOrder"
On Error GoTo Err

isKnownOrder = mKnownOrders.Contains(pBrokerId, mTwsAPI.ClientID)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function


