VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderSubmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

Implements IOrderSubmitter
Implements ITwsConnectionStateListener
Implements TaskCompletionListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "OrderSubmitter"

Private Const OrderModeEntry            As String = "entry"
Private Const OrderModeStop             As String = "stop loss"
Private Const OrderModeTarget           As String = "target"
Private Const OrderModeCloseout         As String = "closeout"

Private Const EntryOrderIdIncrement     As Long = 0
Private Const StopOrderIdIncrement      As Long = 1
Private Const TargetOrderIdIncrement    As Long = 2

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mOrderPlacer                                As OrderPlacer

Private mTwsAPI                                     As TwsAPI
Attribute mTwsAPI.VB_VarHelpID = -1
Private mContractRequester                          As ContractDetailsRequester
Private mContractCache                              As TwsContractCache

Private mState                                      As OrderSubmitterStates
Private mStateChangeListeners                       As Listeners

Private mOrderSubmissionListeners                   As Listeners

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mStateChangeListeners = New Listeners
Set mOrderSubmissionListeners = New Listeners
End Sub

Private Sub Class_Terminate()
Finish
Debug.Print "OrderSubmitter terminated"
End Sub

'================================================================================
' IOrderSubmitter Interface Members
'================================================================================

Private Sub IOrderSubmitter_AddOrderSubmissionListener(ByVal pListener As IOrderSubmissionListener)
Const ProcName As String = "IOrderSubmitter_AddOrderSubmissionListener"
On Error GoTo Err

AddOrderSubmissionListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_AddStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "IOrderSubmitter_AddStateChangeListener"
On Error GoTo Err

AddStateChangeListener pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IOrderSubmitter_CancelBracketOrder(ByVal pBracketOrder As IBracketOrder) As Boolean
Const ProcName As String = "IOrderSubmitter_cancelBracketOrder"
On Error GoTo Err

IOrderSubmitter_CancelBracketOrder = CancelBracketOrder(pBracketOrder)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IOrderSubmitter_CancelStopOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IOrderSubmitter_cancelStopOrder"
On Error GoTo Err

CancelStopOrder pBracketOrder

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_CancelTargetOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IOrderSubmitter_cancelTargetOrder"
On Error GoTo Err

CancelTargetOrder pBracketOrder

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_ExecuteBracketOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IOrderSubmitter_executeBracketOrder"
On Error GoTo Err

ExecuteBracketOrder pBracketOrder

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IOrderSubmitter_FetchPermittedOrderProperties(ByVal pContractFuture As IFuture, Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "IOrderSubmitter_FetchPermittedOrderProperties"
On Error GoTo Err

Set IOrderSubmitter_FetchPermittedOrderProperties = FetchPermittedOrderProperties(pContractFuture, pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IOrderSubmitter_Finish()
Const ProcName As String = "IOrderSubmitter_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_ModifyOrder(ByVal pOrder As IOrder)
Const ProcName As String = "IOrderSubmitter_ModifyOrder"
On Error GoTo Err

ModifyOrder pOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_ModifyBracketOrder( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal EntryOrderChanged As Boolean, _
                ByVal StopOrderChanged As Boolean, _
                ByVal TargetOrderChanged As Boolean, _
                ByVal CloseoutOrderChanged As Boolean)
Const ProcName As String = "IOrderSubmitter_ModifyBracketOrder"
On Error GoTo Err

ModifyBracketOrder pBracketOrder, EntryOrderChanged, StopOrderChanged, TargetOrderChanged, CloseoutOrderChanged

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_NotifyRecoveredOrder(ByVal pOrder As IOrder)
Const ProcName As String = "IOrderSubmitter_NotifyRecoveredOrder"
On Error GoTo Err

NotifyRecoveredOrder pOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_PlaceOrder(ByVal pOrder As IOrder)
Const ProcName As String = "IOrderSubmitter_PlaceOrder"
On Error GoTo Err

PlaceOrder pOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_RemoveOrderSubmissionListener(ByVal pListener As IOrderSubmissionListener)
Const ProcName As String = "IOrderSubmitter_RemoveOrderSubmissionListener"
On Error GoTo Err

RemoveOrderSubmissionListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_RemoveStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "IOrderSubmitter_RemoveStateChangeListener"
On Error GoTo Err

RemoveStateChangeListener pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_ResubmitStopAndTargetOrders(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IOrderSubmitter_resubmitStopAndTargetOrders"
On Error GoTo Err

ResubmitStopAndTargetOrders pBracketOrder

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_ResubmitStopOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IOrderSubmitter_resubmitStopOrder"
On Error GoTo Err

ResubmitStopOrder pBracketOrder

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_ResubmitTargetOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IOrderSubmitter_resubmitTargetOrder"
On Error GoTo Err

ResubmitTargetOrder pBracketOrder
Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IOrderSubmitter_State() As OrderSubmitterStates
IOrderSubmitter_State = State
End Property

'================================================================================
' ITwsConnectionStateListener Interface Members
'================================================================================

Private Sub ITwsConnectionStateListener_NotifyAPIConnectionStateChange(ByVal pSource As Object, ByVal pState As ApiConnectionStates, ByVal pMessage As String)
Const ProcName As String = "ITwsConnectionStateListener_NotifyAPIConnectionStateChange"
On Error GoTo Err

Select Case pState
Case ApiConnNotConnected
    setState OrderSubmitterStateNotReady
Case ApiConnConnecting

Case ApiConnConnected
    setState OrderSubmitterStateReady
Case ApiConnFailed
    setState OrderSubmitterStateNotReady
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionClosed(ByVal pSource As Object)
Const ProcName As String = "ITwsConnectionStateListener_NotifyIBServerConnectionClosed"
On Error GoTo Err

setState OrderSubmitterStateNotReady

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionRecovered(ByVal pSource As Object, ByVal pDataLost As Boolean)
Const ProcName As String = "ITwsConnectionStateListener_NotifyIBServerConnectionRecovered"
On Error GoTo Err

setState OrderSubmitterStateReady

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "TaskCompletionListener_TaskCompleted"
On Error GoTo Err

Dim lDeferredAction As DeferredOrderAction
Set lDeferredAction = ev.Cookie

If ev.Cancelled Then
ElseIf ev.ErrorNumber <> 0 Then
    If lDeferredAction.TargetIsBracketOrder Then
        Assert False, "Deferred " & IIf(lDeferredAction.ActionIsModify, "modify", "execute") & "bracket order failed: " & ev.ErrorMessage, ev.ErrorNumber
    Else
        Assert False, "Deferred " & IIf(lDeferredAction.ActionIsModify, "modify", "place") & "order failed: " & ev.ErrorMessage, ev.ErrorNumber
    End If
Else
    Dim lContract As TwsContract
    Set lContract = ev.Result
    If lDeferredAction.TargetIsBracketOrder Then
        If lDeferredAction.ActionIsModify Then
            modifyABracketOrder lDeferredAction.BracketOrder, lContract, lDeferredAction.EntryOrderChanged, lDeferredAction.StopOrderChanged, lDeferredAction.TargetOrderChanged, lDeferredAction.CloseoutOrderChanged
        Else
            execBracketOrder lDeferredAction.BracketOrder, lContract
        End If
    Else
        placeAnOrder lDeferredAction.Order, lContract, lDeferredAction.Transmit, lDeferredAction.ParentId, lDeferredAction.OcaGroup, lDeferredAction.OrderMode
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get State() As OrderSubmitterStates
State = mState
End Property

'================================================================================
' Methods
'================================================================================

Public Sub AddOrderSubmissionListener(ByVal pListener As IOrderSubmissionListener)
Const ProcName As String = "AddOrderSubmissionListener"
On Error GoTo Err

mOrderSubmissionListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "AddStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function CancelBracketOrder(ByVal pBracketOrder As IBracketOrder) As Boolean
Const ProcName As String = "CancelBracketOrder"
On Error GoTo Err

Dim lEntryOrder As IOrder
Set lEntryOrder = pBracketOrder.EntryOrder

Select Case lEntryOrder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
Case Else
    CancelOrder lEntryOrder, OrderModeEntry ' should automatically cancel the other orders
                                            ' if they have parentid set
    CancelBracketOrder = True
End Select

Dim lStopOrder As IOrder
Set lStopOrder = pBracketOrder.StopLossOrder
If Not lStopOrder Is Nothing Then
    Select Case lStopOrder.Status
    Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusFilled, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusCancelled
    Case Else
        Select Case lEntryOrder.Status
        Case OrderStatuses.OrderStatusFilled, _
            OrderStatuses.OrderStatusCancelling, _
            OrderStatuses.OrderStatusCancelled, _
            OrderStatuses.OrderStatusRejected
            CancelOrder lStopOrder, OrderModeStop
            CancelBracketOrder = True
        End Select
    End Select
End If

Dim lTargetOrder As IOrder
Set lTargetOrder = pBracketOrder.TargetOrder
If Not lTargetOrder Is Nothing Then
    Select Case lTargetOrder.Status
    Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusFilled, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusCancelled
    Case Else
        Select Case lEntryOrder.Status
        Case OrderStatuses.OrderStatusFilled, _
            OrderStatuses.OrderStatusCancelling, _
            OrderStatuses.OrderStatusCancelled, _
            OrderStatuses.OrderStatusRejected
            CancelOrder lTargetOrder, OrderModeTarget
            CancelBracketOrder = True
        End Select
    End Select
End If
    
' need some stuff here to cancel if it's an oca group

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub CancelStopOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "CancelStopOrder"
On Error GoTo Err

Dim lStopOrder As IOrder
Set lStopOrder = pBracketOrder.StopLossOrder

Assert Not lStopOrder Is Nothing, "Order plex " & pBracketOrder.Key & " has no stop Order"

Select Case lStopOrder.Status
Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusFilled, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusCancelled
    Assert False, "Stop Order state invalid for cancellation"
Case Else
    CancelOrder lStopOrder, OrderModeStop
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub CancelTargetOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "CancelTargetOrder"
On Error GoTo Err

Dim lTargetOrder As IOrder
Set lTargetOrder = pBracketOrder.TargetOrder

Assert Not lTargetOrder Is Nothing, "Order plex " & pBracketOrder.Key & " has no target Order"

Select Case lTargetOrder.Status
Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusCancelled
    Assert False, "Target Order state invalid for cancellation"
Case Else
    CancelOrder lTargetOrder, OrderModeTarget
End Select
Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ExecuteBracketOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "ExecuteBracketOrder"
On Error GoTo Err

If pBracketOrder.Contract.Specifier.Sectype = SecTypeCombo Then
    Dim lTask As ComboContractConversionTask
    Set lTask = New ComboContractConversionTask
    lTask.Initialise mContractRequester, pBracketOrder.Contract.Specifier
    
    Dim lCookie As New DeferredOrderAction
    lCookie.TargetIsBracketOrder = True
    lCookie.BracketOrder = pBracketOrder
    StartTask(lTask, PriorityNormal, , lCookie).AddTaskCompletionListener Me
Else
    execBracketOrder pBracketOrder, gContractSpecToTwsContract(pBracketOrder.Contract.Specifier)
End If

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function FetchPermittedOrderProperties( _
                ByVal pContractFuture As IFuture, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchPermittedOrderAttributes"
On Error GoTo Err

Dim lFetchTask As New PermittedOrderPropsFetchTask

lFetchTask.Initialise pContractFuture, mContractRequester, mContractCache, pCookie
StartTask lFetchTask, PriorityLow, , pCookie
Set FetchPermittedOrderProperties = lFetchTask.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

mStateChangeListeners.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTwsAPI As TwsAPI, _
                ByVal pOrderPlacer As OrderPlacer, _
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractCache As TwsContractCache)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mOrderPlacer = pOrderPlacer
Set mTwsAPI = pTwsAPI
If mTwsAPI.ConnectionState = TwsConnConnected Then
    setState OrderSubmitterStateReady
Else
    setState OrderSubmitterStateNotReady
End If
Set mContractRequester = pContractRequester
Set mContractCache = pContractCache

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ModifyOrder(ByVal pOrder As IOrder)
Const ProcName As String = "ModifyOrder"
On Error GoTo Err

If pOrder.ContractSpecifier.Sectype = SecTypeCombo Then
    Dim lTask As ComboContractConversionTask
    Set lTask = New ComboContractConversionTask
    lTask.Initialise mContractRequester, pOrder.ContractSpecifier
    
    Dim lCookie As New DeferredOrderAction
    lCookie.TargetIsBracketOrder = False
    lCookie.ActionIsModify = True
    lCookie.Order = pOrder
    lCookie.OrderMode = OrderModeEntry
    lCookie.ParentId = "0"
    lCookie.Transmit = True
    StartTask(lTask, PriorityNormal, , lCookie).AddTaskCompletionListener Me
Else
    placeAnOrder pOrder, gContractSpecToTwsContract(pOrder.ContractSpecifier), True, 0, "", OrderModeEntry
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ModifyBracketOrder( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pEntryOrderChanged As Boolean, _
                ByVal pStopOrderChanged As Boolean, _
                ByVal pTargetOrderChanged As Boolean, _
                ByVal pCloseoutOrderChanged As Boolean)
Const ProcName As String = "ModifyBracketOrder"
On Error GoTo Err

If pBracketOrder.Contract.Specifier.Sectype = SecTypeCombo Then
    Dim lTask As ComboContractConversionTask
    Set lTask = New ComboContractConversionTask
    lTask.Initialise mContractRequester, pBracketOrder.Contract.Specifier
    
    Dim lCookie As New DeferredOrderAction
    lCookie.TargetIsBracketOrder = True
    lCookie.ActionIsModify = True
    lCookie.CloseoutOrderChanged = pCloseoutOrderChanged
    lCookie.EntryOrderChanged = pEntryOrderChanged
    lCookie.StopOrderChanged = pStopOrderChanged
    lCookie.TargetOrderChanged = pTargetOrderChanged
    lCookie.BracketOrder = pBracketOrder
    StartTask(lTask, PriorityNormal, , lCookie).AddTaskCompletionListener Me
Else
    modifyABracketOrder pBracketOrder, _
                        gContractSpecToTwsContract(pBracketOrder.Contract.Specifier), _
                        pEntryOrderChanged, _
                        pStopOrderChanged, _
                        pTargetOrderChanged, _
                        pCloseoutOrderChanged
End If

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyError( _
                ByVal pOrderId As String, _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "NotifyError"
On Error GoTo Err

mOrderSubmissionListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mOrderSubmissionListeners.Count
    Dim lListener As IOrderSubmissionListener
    Set lListener = mOrderSubmissionListeners.GetListener(i)
    lListener.NotifyError pOrderId, pErrorCode, pErrorMsg
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyExecutionReport(ByVal pExecutionReport As IExecutionReport)
Const ProcName As String = "NotifyExecutionReport"
On Error GoTo Err

mOrderSubmissionListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mOrderSubmissionListeners.Count
    Dim lListener As IOrderSubmissionListener
    Set lListener = mOrderSubmissionListeners.GetListener(i)
    lListener.NotifyExecutionReport pExecutionReport
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyOrderReport(ByVal pOrderReport As IOrderReport)
Const ProcName As String = "NotifyOrderReport"
On Error GoTo Err

mOrderSubmissionListeners.SetCurrentListeners

Dim i As Long
For i = 1 To mOrderSubmissionListeners.Count
    Dim lListener As IOrderSubmissionListener
    Set lListener = mOrderSubmissionListeners.GetListener(i)
    lListener.NotifyOrderReport pOrderReport
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyOrderStatusReport(ByVal pOrderStatusReport As IOrderStatusReport)
Const ProcName As String = "NotifyOrderStatusReport"
On Error GoTo Err

mOrderSubmissionListeners.SetCurrentListeners

Dim i As Long
For i = 1 To mOrderSubmissionListeners.Count
    Dim lListener As IOrderSubmissionListener
    Set lListener = mOrderSubmissionListeners.GetListener(i)
    lListener.NotifyOrderStatusReport pOrderStatusReport
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyRecoveredOrder(ByVal pOrder As IOrder)
Const ProcName As String = "NotifyRecoveredOrder"
On Error GoTo Err

mOrderPlacer.NotifyRecoveredOrder Me, pOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PlaceOrder(ByVal pOrder As IOrder)
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

placeOrderEx pOrder, True, 0, "", OrderModeEntry

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveOrderSubmissionListener(ByVal pListener As IOrderSubmissionListener)
Const ProcName As String = "RemoveOrderSubmissionListener"
On Error GoTo Err

mOrderSubmissionListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ResubmitStopAndTargetOrders(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "ResubmitStopAndTargetOrders"
On Error GoTo Err

Dim lOCAGroup As String
lOCAGroup = GenerateGUIDString

placeOrderEx pBracketOrder.StopLossOrder, True, "0", lOCAGroup, OrderModeStop
placeOrderEx pBracketOrder.TargetOrder, True, "0", lOCAGroup, OrderModeTarget

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ResubmitStopOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "ResubmitStopOrder"
On Error GoTo Err

If pBracketOrder.TargetOrder Is Nothing Then
    placeOrderEx pBracketOrder.StopLossOrder, True, "0", "", OrderModeStop
Else
    Dim lOCAGroup As String
    lOCAGroup = GenerateGUIDString
    placeOrderEx pBracketOrder.StopLossOrder, False, "0", lOCAGroup, OrderModeStop
    placeOrderEx pBracketOrder.TargetOrder, True, "0", lOCAGroup, OrderModeTarget
End If

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ResubmitTargetOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "ResubmitTargetOrder"
On Error GoTo Err

If pBracketOrder.StopLossOrder Is Nothing Then
    placeOrderEx pBracketOrder.TargetOrder, True, "0", "", OrderModeTarget
Else
    Dim lOCAGroup As String
    lOCAGroup = GenerateGUIDString
    placeOrderEx pBracketOrder.StopLossOrder, False, "0", lOCAGroup, OrderModeStop
    placeOrderEx pBracketOrder.TargetOrder, True, "0", lOCAGroup, OrderModeTarget
End If
Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub CancelOrder( _
                ByVal pOrder As IOrder, _
                ByRef pOrderMode As String)
Const ProcName As String = "CancelOrder"
On Error GoTo Err

logMessage "Cancel " & pOrderMode & " Order: " & _
            "broker id=" & pOrder.BrokerId & _
            "; TradeBuild id=" & pOrder.Id, _
            ProcName

mOrderPlacer.CancelOrder pOrder


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub execBracketOrder( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pContract As TwsContract)
Const ProcName As String = "execBracketOrder"
On Error GoTo Err

Dim lEntryOrder As IOrder
Set lEntryOrder = pBracketOrder.EntryOrder

Dim lStopOrder As IOrder
Set lStopOrder = pBracketOrder.StopLossOrder

Dim lTargetOrder As IOrder
Set lTargetOrder = pBracketOrder.TargetOrder

Dim lPlexId As Long
lPlexId = getNextBracketId

If Not lEntryOrder Is Nothing Then
    lEntryOrder.BrokerId = lPlexId + EntryOrderIdIncrement
    
    Dim Transmit As Boolean
    Transmit = True
    
    If Not lStopOrder Is Nothing Then
        lStopOrder.BrokerId = lPlexId + StopOrderIdIncrement
        If lStopOrder.OrderType <> OrderTypes.OrderTypeAutoStop Then Transmit = False
    End If
    If Not lTargetOrder Is Nothing Then
        lTargetOrder.BrokerId = lPlexId + TargetOrderIdIncrement
        If lTargetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then Transmit = False
    End If
        
    placeAnOrder lEntryOrder, pContract, Transmit, 0, "", OrderModeEntry
    
    If Not lStopOrder Is Nothing Then
        Transmit = True
        If Not lTargetOrder Is Nothing Then
            If lTargetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then Transmit = False
        End If
        
        If lStopOrder.OrderType <> OrderTypes.OrderTypeAutoStop Then
            ' note that AUTOSTP orders will be sent when the entry Order is filled
            placeAnOrder lStopOrder, pContract, Transmit, lEntryOrder.BrokerId, "", OrderModeStop
        End If
    End If
    
    If Not lTargetOrder Is Nothing Then
        If lTargetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then
            placeAnOrder lTargetOrder, pContract, True, lEntryOrder.BrokerId, "", OrderModeTarget
        End If
    End If
    
Else
    ' treat the other orders as an OCA group - still to be implemented
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireStateChange()
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = mState

mStateChangeListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mStateChangeListeners.Count
    Dim lListener As StateChangeListener
    Set lListener = mStateChangeListeners.GetListener(i)
    lListener.Change ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getNextBracketId() As Long
Const ProcName As String = "getNextBracketId"
On Error GoTo Err

getNextBracketId = Int((mTwsAPI.NextOrderId + 4) / 5) * 5

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Sub modifyABracketOrder( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pContract As TwsContract, _
                ByVal pEntryOrderChanged As Boolean, _
                ByVal pStopOrderChanged As Boolean, _
                ByVal pTargetOrderChanged As Boolean, _
                ByVal pCloseoutOrderChanged As Boolean)
Const ProcName As String = "modifyABracketOrder"
On Error GoTo Err

Dim lParentId As Long

Select Case pBracketOrder.EntryOrder.Status
Case OrderStatusFilled, _
        OrderStatusCancelling, _
        OrderStatusCancelled
    lParentId = 0
Case Else
    lParentId = pBracketOrder.EntryOrder.BrokerId
End Select

If pBracketOrder.EntryOrder.QuantityRemaining = 0 Then lParentId = 0

Dim lOCAGroup As String
If Not pBracketOrder.StopLossOrder Is Nothing Then lOCAGroup = pBracketOrder.StopLossOrder.ProviderProperties.GetParameterValue("OCA group")
If lOCAGroup = "" And Not pBracketOrder.TargetOrder Is Nothing Then lOCAGroup = pBracketOrder.TargetOrder.ProviderProperties.GetParameterValue("OCA group")
If lOCAGroup = "" Then lOCAGroup = GenerateGUIDString

If pEntryOrderChanged Then
    placeAnOrder pBracketOrder.EntryOrder, pContract, True, 0, "", OrderModeEntry
End If
If pStopOrderChanged Then
    placeAnOrder pBracketOrder.StopLossOrder, pContract, True, lParentId, lOCAGroup, OrderModeStop
End If
If pTargetOrderChanged Then
    placeAnOrder pBracketOrder.TargetOrder, pContract, True, lParentId, lOCAGroup, OrderModeTarget
End If
If pCloseoutOrderChanged Then
    placeAnOrder pBracketOrder.CloseoutOrder, pContract, True, 0, "", OrderModeCloseout
End If

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyOpenOrder(ByVal pOrder As IOrder)
Const ProcName As String = "notifyOpenOrder"
On Error GoTo Err

Dim lOrderReport As IOrderReport
Set lOrderReport = New OrderReport

With lOrderReport
    .Action = gTwsOrderActionToOrderAction(pOrder.Action)
    .AllOrNone = pOrder.AllOrNone
    .BlockOrder = pOrder.BlockOrder
    .BrokerId = pOrder.BrokerId
    .DiscretionaryAmount = pOrder.DiscretionaryAmount
    .DisplaySize = pOrder.DisplaySize
    .ETradeOnly = pOrder.ETradeOnly
    .FirmQuoteOnly = pOrder.FirmQuoteOnly
    .GoodAfterTime = pOrder.GoodAfterTime
    .GoodAfterTimeTZ = pOrder.GoodAfterTimeTZ
    .GoodTillDate = pOrder.GoodTillDate
    .GoodTillDateTZ = pOrder.GoodTillDateTZ
    .Hidden = pOrder.Hidden
    .Id = pOrder.Id
    .IgnoreRegularTradingHours = pOrder.IgnoreRegularTradingHours
    .LimitPrice = pOrder.LimitPrice
    .MinimumQuantity = pOrder.MinimumQuantity
    .NbboPriceCap = pOrder.NbboPriceCap
    .OrderType = pOrder.OrderType
    .Origin = pOrder.Origin
    .OriginatorRef = pOrder.OriginatorRef
    .OverrideConstraints = pOrder.OverrideConstraints
    .Quantity = pOrder.Quantity
    .SettlingFirm = pOrder.SettlingFirm
    .Status = pOrder.Status
    .StopTriggerMethod = pOrder.StopTriggerMethod
    .SweepToFill = pOrder.SweepToFill
    .TimeInForce = pOrder.TimeInForce
    .TriggerPrice = pOrder.TriggerPrice
End With

NotifyOrderReport lOrderReport

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyStatus(ByVal pOrder As IOrder, ByVal pStatus As OrderStatuses)
Const ProcName As String = "notifyStatus"
On Error GoTo Err

Dim statusRpt As New OrderStatusReport
statusRpt.OrderId = pOrder.Id
statusRpt.BrokerId = pOrder.BrokerId
statusRpt.Status = pStatus
NotifyOrderStatusReport statusRpt

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub placeAnOrder( _
                ByVal pOrder As IOrder, _
                ByVal pContract As TwsContract, _
                ByVal pTransmit As Boolean, _
                ByVal pParentId As Long, _
                ByVal pOcaGroup As String, _
                ByRef pOrderMode As String)
Const ProcName As String = "placeAnOrder"
On Error GoTo Err

logMessage "Place " & pOrderMode & " Order: " & _
                "broker id=" & pOrder.BrokerId & _
                "; TradeBuild id=" & pOrder.Id, _
            ProcName

Assert pOrder.Id <> "", "Order has no id"
Assert pOrder.Action <> OrderActionNone, "Order has no action"
Assert pOrder.OrderType <> OrderTypeNone, "Order has no order type"
Assert pOrder.Quantity <> 0, "Order has zero quantity"

mOrderPlacer.PlaceOrder Me, pOrder, pContract, pTransmit, pParentId, pOcaGroup

If pOrder.Status = OrderStatusCreated Then
    notifyOpenOrder pOrder
    notifyStatus pOrder, OrderStatusPendingSubmit
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub placeOrderEx( _
                ByVal pOrder As IOrder, _
                ByVal pTransmit As Boolean, _
                ByVal pParentId As Long, _
                ByVal pOcaGroup As String, _
                ByRef pOrderMode As String)
Const ProcName As String = "placeOrderEx"
On Error GoTo Err

pOrder.BrokerId = CStr(mTwsAPI.NextOrderId)

If pOrder.ContractSpecifier.Sectype = SecTypeCombo Then
    Dim lTask As ComboContractConversionTask
    Set lTask = New ComboContractConversionTask
    lTask.Initialise mContractRequester, pOrder.ContractSpecifier
    
    Dim lCookie As New DeferredOrderAction
    lCookie.TargetIsBracketOrder = False
    lCookie.Order = pOrder
    lCookie.OcaGroup = pOcaGroup
    lCookie.OrderMode = pOrderMode
    lCookie.ParentId = pParentId
    lCookie.Transmit = pTransmit
    StartTask(lTask, PriorityNormal, , pOrder).AddTaskCompletionListener Me
Else
    placeAnOrder pOrder, gContractSpecToTwsContract(pOrder.ContractSpecifier), pTransmit, pParentId, pOcaGroup, pOrderMode
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setState(ByVal pState As OrderSubmitterStates)
Const ProcName As String = "setState"
On Error GoTo Err

If mState = pState Then Exit Sub
mState = pState
fireStateChange

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




