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
Implements ITwsConnectionStateListener

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
Private Const ReconnectionExecutionsRequestId       As Long = 333

'@================================================================================
' Member variables
'@================================================================================

Private mTwsAPI                                     As TwsAPI
Attribute mTwsAPI.VB_VarHelpID = -1

Private mKnownOrders                                As New KnownOrders

Private mOrderRecoveryListener                      As IOrderRecoveryListener
Private mStartOrderRecoveryFrom                     As Date

Private mOrderRecoveryFinished                      As Boolean

Private mQueuedOrders                               As New EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IOrderInfoConsumer Interface Members
'@================================================================================

Private Sub IOrderInfoConsumer_EndExecutions(ByVal pRequestId As Long)
Const ProcName As String = "IOrderInfoConsumer_EndExecutions"
On Error GoTo Err

If Not mOrderRecoveryListener Is Nothing Then mOrderRecoveryListener.EndExecutionReports
If mOrderRecoveryFinished Then
    Set mOrderRecoveryListener = Nothing
    processQueuedOrders
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderInfoConsumer_EndOpenOrders()
Const ProcName As String = "IOrderInfoConsumer_EndOpenOrders"
On Error GoTo Err

If Not mOrderRecoveryListener Is Nothing Then mOrderRecoveryListener.EndOrderReports
If mOrderRecoveryFinished Then
    Set mOrderRecoveryListener = Nothing
    processQueuedOrders
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderInfoConsumer_NotifyError(ByVal pOrderId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IOrderInfoConsumer_NotifyError"
On Error GoTo Err

If Not mKnownOrders.Contains(pOrderId, mTwsAPI.ClientId) Then Exit Sub

Select Case pErrorCode
Case 161
    ' TWS sends this when an attempt is made to cancel an order that is already cancelled.
    ' Since this can be a race situation, where TWS has cancelled an order but the
    ' notification hasn't arrived via the API, and the application also tries to cancel,
    ' we just ignore it.
    gLog "Ignored error 161 from TWS (" & pOrderId & "): " & pErrorMsg, ModuleName, ProcName
Case 202
    notifyStatus pOrderId, OrderStatuses.OrderStatusCancelled
Case 451
    ' TWS sends this when it is unable to transmit the order as requested, typically because
    ' the order value exceeds some restriction
    GetOrderSubmitter(pOrderId).NotifyMessage GetOrder(pOrderId).Id, pErrorMsg
Case 382, 383
    ' indicates that the order size or price exceeds the constraints specified
    ' in TWS.
    GetOrderSubmitter(pOrderId).NotifyError GetOrder(pOrderId).Id, pErrorCode, pErrorMsg
    notifyStatus pOrderId, OrderStatuses.OrderStatusRejected
Case 399
    ' contains information about the order which doesn't invalidate it - for example
    ' order won't be submitted to exchange until start of next trading session
    GetOrderSubmitter(pOrderId).NotifyMessage GetOrder(pOrderId).Id, pErrorMsg
Case 10148  ' Trying to cancel an order that is already filled or cancelled
    ' Just ignore this
Case Else
    GetOrderSubmitter(pOrderId).NotifyError GetOrder(pOrderId).Id, pErrorCode, pErrorMsg
    notifyStatus pOrderId, OrderStatuses.OrderStatusRejected
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderInfoConsumer_NotifyExecution(ByVal pReqId As Long, ByVal pContract As TwsContract, ByVal pExecution As TwsExecution)
Const ProcName As String = "IOrderInfoConsumer_NotifyExecution"
On Error GoTo Err

Static sExecutions As New EnumerableCollection

If pReqId = ReconnectionExecutionsRequestId Then
    If sExecutions.Contains(pExecution.ExecId) Then Exit Sub
    sExecutions.Add pExecution, pExecution.ExecId
End If

Dim lOrder As IOrder
Dim lorderId As String
If isKnownOrder(pExecution.OrderId) Then
    Set lOrder = GetOrder(pExecution.OrderId)
    lorderId = lOrder.Id
End If

Dim lExecReport As New Execution
lExecReport.SecurityName = pContract.LocalSymbol & "." & pContract.Exchange
lExecReport.AccountIdentifier = pExecution.AcctNumber
lExecReport.Action = IIf(pExecution.Side = "BOT", OrderActionBuy, OrderActionSell)
lExecReport.Id = pExecution.ExecId
lExecReport.FillingExchange = pExecution.Exchange
lExecReport.IsLiquidation = pExecution.Liquidation
lExecReport.BrokerId = pExecution.OrderId
lExecReport.OrderId = lorderId
lExecReport.Price = pExecution.Price
lExecReport.Quantity = pExecution.Shares
lExecReport.FillTime = TwsDateStringToDate(pExecution.Time)
lExecReport.TimezoneName = GetTimeZone("").StandardName

If Not lOrder Is Nothing Then
    Dim lOrderSubmitter As OrderSubmitter
    Set lOrderSubmitter = GetOrderSubmitter(pExecution.OrderId)
    
    Dim qtyRemaining As Long
    qtyRemaining = lOrder.QuantityRemaining
    
    lOrderSubmitter.NotifyExecutionReport lExecReport
    
    If pReqId <> -1 Then
        ' notify a status update - this is in case this execution occurred during a period when
        ' Tws had lost its connection to the IB servers and this exec details is being notified
        ' in response to the reqExecutions call made when the connection is recovered. In these
        ' circumstances Tws will not itself issue an orderStatus.
        notifyStatus pExecution.OrderId, IIf(pExecution.Shares >= qtyRemaining, OrderStatuses.OrderStatusFilled, OrderStatuses.OrderStatusSubmitted)
    End If
Else
    ' this must be an execution from a previous session
    If Not mOrderRecoveryListener Is Nothing Then mOrderRecoveryListener.NotifyExecutionReport lExecReport
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderInfoConsumer_NotifyOpenOrder(ByVal pOrderId As Long, ByVal pContract As TwsContract, ByVal pOrder As TwsOrder, ByVal pOrderState As TwsOrderState)
Const ProcName As String = "IOrderInfoConsumer_NotifyOpenOrder"
On Error GoTo Err

If pOrder.ClientId <> mTwsAPI.ClientId Then
    gLog "Ignoring order " & pOrder.OrderId & " for clientID " & pOrder.ClientId, ModuleName, ProcName
    Exit Sub
End If

Dim lorderId As String
If isKnownOrder(pOrderId) Then
    lorderId = GetOrder(pOrderId).Id
End If

Dim lOrderReport As IOrderReport
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
    
    Dim lTimezoneName As String ' returned byref from TwsDateStringToDate
    .GoodAfterTime = TwsDateStringToDate(pOrder.GoodAfterTime, gTwsTimezoneNameToStandardTimeZoneName(lTimezoneName))
    .GoodAfterTimeTZ = lTimezoneName
    .GoodTillDate = TwsDateStringToDate(pOrder.GoodTillDate, lTimezoneName)
    .GoodTillDateTZ = lTimezoneName
    
    .Hidden = pOrder.Hidden
    .Id = lorderId
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
    .StopTriggerMethod = gTwsTriggerMethodToStopTriggerMethod(pOrder.TriggerMethod)
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
gHandleUnexpectedError ProcName, ModuleName
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

Dim lStatus As OrderStatuses

lStatus = gOrderStatusFromString(pStatus)
If isKnownOrder(pOrderId) Then
    notifyStatus pOrderId, lStatus
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITwsConnectionStateListener Interface Members
'@================================================================================

Private Sub ITwsConnectionStateListener_NotifyAPIConnectionStateChange(ByVal pSource As Object, ByVal pState As ApiConnectionStates, ByVal pMessage As String)
Const ProcName As String = "ITwsConnectionStateListener_NotifyAPIConnectionStateChange"
On Error GoTo Err

Select Case pState
Case ApiConnNotConnected

Case ApiConnConnecting

Case ApiConnConnected
    doStartOrderRecovery ReconnectionExecutionsRequestId
Case ApiConnFailed

End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionClosed(ByVal pSource As Object)
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionRecovered(ByVal pSource As Object, ByVal pDataLost As Boolean)
Const ProcName As String = "ITwsConnectionStateListener_NotifyIBServerConnectionRecovered"
On Error GoTo Err

doStartOrderRecovery ReconnectionExecutionsRequestId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

If Not mKnownOrders.Contains(pOrder.BrokerId, mTwsAPI.ClientId) Then
    ' this can happen during order recovery if the order has been cancelled
    ' between sessions
    mKnownOrders.Add pOrder, Nothing, mTwsAPI.ClientId, 0
End If

mKnownOrders.Cancel pOrder.BrokerId, mTwsAPI.ClientId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FinishOrderRecovery()
mOrderRecoveryFinished = True
End Sub

Friend Sub Initialise( _
                ByVal pClient As TwsAPI)
Set mTwsAPI = pClient
End Sub

Friend Sub NotifyRecoveredOrder( _
                ByVal pOrderSubmitter As OrderSubmitter, _
                ByVal pOrder As IOrder)
Const ProcName As String = "NotifyRecoveredOrder"
On Error GoTo Err

mKnownOrders.Add pOrder, pOrderSubmitter, mTwsAPI.ClientId, 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub PlaceOrder( _
                ByVal pOrderSubmitter As OrderSubmitter, _
                ByVal pOrder As IOrder, _
                ByVal pContract As TwsContract, _
                ByVal pTransmit As Boolean, _
                ByVal pParentId As Long, _
                ByVal pOcaGroup As String)
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

Dim mapEntry As OrderIdMapEntry
Dim lOrder As TwsOrder

mKnownOrders.Add pOrder, pOrderSubmitter, mTwsAPI.ClientId, pParentId

Set lOrder = gOrderToTwsOrder(pOrder)
lOrder.Transmit = pTransmit
lOrder.ParentId = pParentId
lOrder.OcaGroup = pOcaGroup

If mTwsAPI.ConnectionState = TwsConnConnected Then
    mTwsAPI.PlaceOrder lOrder, pContract
Else
    queueOrder lOrder, pContract
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                    
Friend Sub StartOrderRecovery( _
                ByVal pOrderRecoveryListener As IOrderRecoveryListener, _
                ByVal pStartFrom As Date)
Const ProcName As String = "StartOrderRecovery"
On Error GoTo Err

Assert mOrderRecoveryListener Is Nothing, "Order recovery is already in progress"

Set mOrderRecoveryListener = pOrderRecoveryListener
mStartOrderRecoveryFrom = pStartFrom

If Not (mTwsAPI.ConnectionState = TwsConnConnected And mTwsAPI.IsTwsConnectedToIBServers) Then
    gLog "Can't start order recovery: not fully connected", ModuleName, ProcName
    Exit Sub
End If

doStartOrderRecovery OrderRecoveryExecutionsRequestId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub doStartOrderRecovery(pReqId As Long)
Const ProcName As String = "doStartOrderRecovery"
On Error GoTo Err

gLog "Requesting open orders and executions", ModuleName, ProcName

gRequestOpenOrders mTwsAPI

gRequestExecutions mTwsAPI, pReqId, mTwsAPI.ClientId, mStartOrderRecoveryFrom

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub generateDeferredOrderCancelNotification( _
                ByVal pBrokerId As String)
Const ProcName As String = "generateDeferredOrderCancelNotification"
On Error GoTo Err

Dim lMapEntry As OrderIdMapEntry

If Not isKnownOrder(pBrokerId) Then Exit Sub

Set lMapEntry.Order = mKnownOrders.GetOrder(pBrokerId, mTwsAPI.ClientId)
Set lMapEntry.OrderSubmitter = mKnownOrders.GetOrderSubmitter(pBrokerId, mTwsAPI.ClientId)

DeferAction New DeferredCancelStatus, lMapEntry, 500

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function GetOrder(ByVal pBrokerId As String) As IOrder
Const ProcName As String = "getOrder"
On Error GoTo Err

Set GetOrder = mKnownOrders.GetOrder(pBrokerId, mTwsAPI.ClientId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function GetOrderSubmitter(ByVal pBrokerId As String) As OrderSubmitter
Const ProcName As String = "getOrderSubmitter"
On Error GoTo Err

Set GetOrderSubmitter = mKnownOrders.GetOrderSubmitter(pBrokerId, mTwsAPI.ClientId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isKnownOrder(ByVal pBrokerId As String) As Boolean
Const ProcName As String = "isKnownOrder"
On Error GoTo Err

isKnownOrder = mKnownOrders.Contains(pBrokerId, mTwsAPI.ClientId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub notifyStatus( _
                ByVal pOrderId As String, _
                ByVal pStatus As OrderStatuses)
Const ProcName As String = "notifyStatus"
On Error GoTo Err

Dim lStatusReport As New OrderStatusReport
Dim lOrder As IOrder

If pStatus = OrderStatuses.OrderStatusCancelled Then mKnownOrders.Cancel pOrderId, mTwsAPI.ClientId

Set lOrder = GetOrder(pOrderId)
lStatusReport.OrderId = lOrder.Id
lStatusReport.BrokerId = lOrder.BrokerId
lStatusReport.Status = pStatus
GetOrderSubmitter(pOrderId).NotifyOrderStatusReport lStatusReport

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processQueuedOrders()
Const ProcName As String = "processQueuedOrders"
On Error GoTo Err

' we could do this with a task to space the order submissions out a bit,
' but since there are only rarely expected to be any queued orders we'll
' just do them immediaately

gLog "Processing queued orders: " & mQueuedOrders.Count & " orders to be placed", ProcName, ModuleName

Dim en As Enumerator
Set en = mQueuedOrders.Enumerator

Do While en.MoveNext
    Dim lEntry As OrderQueueEntry
    Set lEntry = en.Current
    mTwsAPI.PlaceOrder lEntry.Order, lEntry.Contract
    en.Remove
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub queueOrder( _
                ByVal pOrder As IOrder, _
                ByVal pContract As TwsContract)
Const ProcName As String = "queueOrder"
On Error GoTo Err

Dim e As New OrderQueueEntry
e.Order = pOrder
e.Contract = pContract
mQueuedOrders.Add pOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

