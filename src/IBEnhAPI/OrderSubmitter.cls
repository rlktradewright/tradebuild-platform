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
Implements ITaskCompletionListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                            As String = "OrderSubmitter"

Private Const OrderModeEntry                        As String = "entry"
Private Const OrderModeStopLoss                     As String = "stop loss"
Private Const OrderModeTarget                       As String = "target"
Private Const OrderModeCloseout                     As String = "closeout"

Private Const EntryOrderIdIncrement                 As Long = 0
Private Const StopLossOrderIdIncrement              As Long = 1
Private Const TargetOrderIdIncrement                As Long = 2
Private Const CloseoutOrderIdIncrement              As Long = 3

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

Private mStageOrders                                As Boolean

Private mIsTwsConnectedToIBServers                  As Boolean

Private mDataSource                                 As IMarketDataSource

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mContract                                   As IContract

Private mOptionTickSizeResolver                     As OptionTickSizeResolver

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
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

Private Sub IOrderSubmitter_AddStateChangeListener(ByVal pListener As IStateChangeListener)
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

Private Sub IOrderSubmitter_CancelOrder(ByVal pOrder As OrderUtils27.IOrder)
Const ProcName As String = "IOrderSubmitter_CancelOrder"
On Error GoTo Err

CancelOrder pOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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

Private Property Get IOrderSubmitter_Capabilities() As OrderSubmitterCapabilities
IOrderSubmitter_Capabilities = Capabilities
End Property

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

Private Function IOrderSubmitter_GetTickSize(ByVal pDummyOrder As IOrder) As IFuture
Const ProcName As String = "IOrderSubmitter_GetTickSize"
On Error GoTo Err

Set IOrderSubmitter_GetTickSize = GetTickSize(pDummyOrder)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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

Private Property Let IOrderSubmitter_OrdersAreStaged(ByVal Value As Boolean)
OrdersAreStaged = Value
End Property

Private Property Get IOrderSubmitter_OrdersAreStaged() As Boolean
IOrderSubmitter_OrdersAreStaged = OrdersAreStaged
End Property

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

Private Sub IOrderSubmitter_RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
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
    ' we'll assume that TWS has a server connection: if not,
    ' it will soon tell us
    mIsTwsConnectedToIBServers = True
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
' ITaskCompletionListener Interface Members
'@================================================================================

Private Sub ITaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "ITaskCompletionListener_TaskCompleted"
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
' mFutureWaiter Event Handlers
'================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsAvailable Then Set mContract = ev.Future.Value

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get Capabilities() As OrderSubmitterCapabilities
Capabilities = OrderSubmitterCapabilities.OrderSubmitterCapabilityCanStageOrders
End Property

Public Property Let OrdersAreStaged(ByVal Value As Boolean)
mStageOrders = Value
End Property

Public Property Get OrdersAreStaged() As Boolean
OrdersAreStaged = mStageOrders
End Property

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

Public Sub AddStateChangeListener(ByVal pListener As IStateChangeListener)
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
    ' should automatically cancel the other orders
    ' if they have parentid set
    cancelAnOrder lEntryOrder, OrderModeEntry
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
            cancelAnOrder lStopOrder, OrderModeStopLoss
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
            cancelAnOrder lTargetOrder, OrderModeTarget
            CancelBracketOrder = True
        End Select
    End Select
End If
    
' need some stuff here to cancel if it's an oca group

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub CancelOrder(ByVal pOrder As IOrder)
Const ProcName As String = "CancelOrder"
On Error GoTo Err

cancelAnOrder pOrder, OrderModeEntry

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
    cancelAnOrder lStopOrder, OrderModeStopLoss
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
    cancelAnOrder lTargetOrder, OrderModeTarget
End Select
Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ExecuteBracketOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "ExecuteBracketOrder"
On Error GoTo Err

If pBracketOrder.Contract.Specifier.SecType = SecTypeCombo Then
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

Public Function GetTickSize(ByVal pDummyOrder As IOrder) As IFuture
Const ProcName As String = "GetTickSize"
On Error GoTo Err

Dim lTickSize As Double
Dim lTickType As TickTypes
Dim lPrice As Double

If mContract.Specifier.SecType = SecTypeOption Or _
    mContract.Specifier.SecType = SecTypeFuturesOption _
Then
    AssertArgument Not pDummyOrder Is Nothing, "No dummy order supplied"
    If Not mOptionTickSizeResolver Is Nothing Then
        Set GetTickSize = mOptionTickSizeResolver.TickSizeFuture
    Else
        lPrice = getAPrice(lTickType)
        Set mOptionTickSizeResolver = New OptionTickSizeResolver
        Set GetTickSize = GetTickSizeForOptions( _
                                mOptionTickSizeResolver, _
                                lPrice, _
                                lTickType, _
                                pDummyOrder)
    End If
    Exit Function
End If

If mContract.Specifier.SecType = SecTypeStock Then
    lPrice = getAPrice(lTickType)
    lTickSize = GetTickSizeForStocks(mContract.Specifier.Exchange, lPrice, mContract.TickSize)
End If

If lTickSize = 0# Then lTickSize = mContract.TickSize

Set GetTickSize = CreateFuture(CreateBoxedValue(lTickSize))

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
                ByVal pContractCache As TwsContractCache, _
                ByVal pDataSource As IMarketDataSource)
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
Set mDataSource = pDataSource
mFutureWaiter.Add mDataSource.ContractFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ModifyOrder(ByVal pOrder As IOrder)
Const ProcName As String = "ModifyOrder"
On Error GoTo Err

If pOrder.ContractSpecifier.SecType = SecTypeCombo Then
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

If pBracketOrder.Contract.Specifier.SecType = SecTypeCombo Then
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
                ByVal pOrder As IOrder, _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "NotifyError"
On Error GoTo Err

If Not mOptionTickSizeResolver Is Nothing Then
    If mOptionTickSizeResolver.HandleError(pErrorCode) Then Exit Sub
End If

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mOrderSubmissionListeners.Valid Then
    sInit = True
    sSomeListeners = mOrderSubmissionListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IOrderSubmissionListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyError pOrder.Id, pErrorCode, pErrorMsg
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyExecutionReport(ByVal pExecutionReport As IExecutionReport)
Const ProcName As String = "NotifyExecutionReport"
On Error GoTo Err

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mOrderSubmissionListeners.Valid Then
    sInit = True
    sSomeListeners = mOrderSubmissionListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IOrderSubmissionListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyExecutionReport pExecutionReport
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyMessage( _
                ByVal pOrder As IOrder, _
                ByVal pMessage As String)
Const ProcName As String = "NotifyMessage"
On Error GoTo Err

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mOrderSubmissionListeners.Valid Then
    sInit = True
    sSomeListeners = mOrderSubmissionListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IOrderSubmissionListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyMessage pOrder.Id, pMessage
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyOrderReport(ByVal pOrderReport As IOrderReport)
Const ProcName As String = "NotifyOrderReport"
On Error GoTo Err

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mOrderSubmissionListeners.Valid Then
    sInit = True
    sSomeListeners = mOrderSubmissionListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IOrderSubmissionListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyOrderReport pOrderReport
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyOrderStatusReport(ByVal pOrderStatusReport As IOrderStatusReport)
Const ProcName As String = "NotifyOrderStatusReport"
On Error GoTo Err

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mOrderSubmissionListeners.Valid Then
    sInit = True
    sSomeListeners = mOrderSubmissionListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IOrderSubmissionListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyOrderStatusReport pOrderStatusReport
    Next
End If

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

Public Sub RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
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

placeOrderEx pBracketOrder.StopLossOrder, False, "0", lOCAGroup, OrderModeStopLoss
placeOrderEx pBracketOrder.TargetOrder, True, "0", lOCAGroup, OrderModeTarget

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ResubmitStopOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "ResubmitStopOrder"
On Error GoTo Err

If pBracketOrder.TargetOrder Is Nothing Then
    placeOrderEx pBracketOrder.StopLossOrder, True, "0", "", OrderModeStopLoss
Else
    Dim lOCAGroup As String
    lOCAGroup = GenerateGUIDString
    placeOrderEx pBracketOrder.StopLossOrder, False, "0", lOCAGroup, OrderModeStopLoss
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
    placeOrderEx pBracketOrder.StopLossOrder, False, "0", lOCAGroup, OrderModeStopLoss
    placeOrderEx pBracketOrder.TargetOrder, True, "0", lOCAGroup, OrderModeTarget
End If
Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub cancelAnOrder( _
                ByVal pOrder As IOrder, _
                ByVal pOrderMode As String)
Const ProcName As String = "cancelAnOrder"
On Error GoTo Err

logMessage "Cancel " & pOrderMode & " Order: " & _
            "broker id=" & pOrder.BrokerId & _
            "; TradeBuild id=" & pOrder.Id, _
            ProcName

mOrderPlacer.CancelOrder Me, pOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub execBracketOrder( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pContract As TwsContract)
Const ProcName As String = "execBracketOrder"
On Error GoTo Err

Dim lEntryOrder As IOrder: Set lEntryOrder = pBracketOrder.EntryOrder

Dim lStopLossOrder As IOrder: Set lStopLossOrder = pBracketOrder.StopLossOrder
Dim lStopLossOrderCanBePlaced As Boolean: lStopLossOrderCanBePlaced = orderCanBePlaced(lStopLossOrder)

Dim lTargetOrder As IOrder: Set lTargetOrder = pBracketOrder.TargetOrder
Dim lTargetOrderCanBePlaced As Boolean: lTargetOrderCanBePlaced = orderCanBePlaced(lTargetOrder)

Dim lPlexId As Long
lPlexId = getNextBracketId

If Not lEntryOrder Is Nothing Then
    Dim Transmit As Boolean
    Transmit = True
    
    If lStopLossOrderCanBePlaced Then Transmit = False
    If lTargetOrderCanBePlaced Then Transmit = False
    lEntryOrder.BrokerId = lPlexId + EntryOrderIdIncrement
    placeAnOrder lEntryOrder, pContract, Transmit, 0, "", OrderModeEntry
    
    If lStopLossOrderCanBePlaced Then
        Transmit = True
        If lTargetOrderCanBePlaced Then Transmit = False
        lStopLossOrder.BrokerId = lPlexId + StopLossOrderIdIncrement
        placeAnOrder lStopLossOrder, pContract, Transmit, lEntryOrder.BrokerId, "", OrderModeStopLoss
    End If
    
    If lTargetOrderCanBePlaced Then
        lTargetOrder.BrokerId = lPlexId + TargetOrderIdIncrement
        placeAnOrder lTargetOrder, pContract, True, lEntryOrder.BrokerId, "", OrderModeTarget
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

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mStateChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mStateChangeListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IStateChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Change ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getAPrice(ByRef pTickType As TickTypes) As Double
Const ProcName As String = "getCurrentPrice"
On Error GoTo Err

Assert mDataSource.State = MarketDataSourceStateRunning, "Data source is not running"
If mDataSource.HasCurrentTick(TickTypeTrade) Then
    pTickType = TickTypeTrade
    getAPrice = mDataSource.CurrentTick(TickTypeTrade).Price
ElseIf mDataSource.HasCurrentTick(TickTypeAsk) Then
    pTickType = TickTypeAsk
    getAPrice = mDataSource.CurrentTick(TickTypeAsk).Price
ElseIf mDataSource.HasCurrentTick(TickTypeBid) Then
    pTickType = TickTypeBid
    getAPrice = mDataSource.CurrentTick(TickTypeBid).Price
Else
    Assert False, "No price available"
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getNextBracketId() As Long
Const ProcName As String = "getNextBracketId"
On Error GoTo Err

getNextBracketId = Int((mTwsAPI.NextOrderId + 4) / 5) * 5

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function GetTickSizeForOptions( _
                ByVal pResolver As OptionTickSizeResolver, _
                ByVal pExamplePrice As Double, _
                ByVal pPriceTickType As TickTypes, _
                ByVal pDummyOrder As IOrder) As IFuture
Const ProcName As String = "GetTickSizeForOptions"
On Error GoTo Err

pResolver.Initialise Me, _
                    pExamplePrice, _
                    IIf(pPriceTickType = TickTypeAsk, True, False), _
                    pDummyOrder
StartTask pResolver, PriorityNormal
Set GetTickSizeForOptions = pResolver.TickSizeFuture

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function GetTickSizeForStocks( _
                ByVal pExchange As String, _
                ByVal pCurrentPrice As Double, _
                ByVal pDefaultTickSize As Double) As Double
Dim lTickSize As Double

Select Case UCase$(pExchange)
Case "LSE", "BATEUK", "CHIXUK", "TRQXUK"
    If pCurrentPrice <= 0.001 Then
        lTickSize = 0.000001
    ElseIf pCurrentPrice <= 0.002 Then
        lTickSize = 0.000001
    ElseIf pCurrentPrice <= 0.005 Then
        lTickSize = 0.000001
    ElseIf pCurrentPrice <= 0.01 Then
        lTickSize = 0.000002
    ElseIf pCurrentPrice <= 0.02 Then
        lTickSize = 0.000005
    ElseIf pCurrentPrice <= 0.05 Then
        lTickSize = 0.00001
    ElseIf pCurrentPrice <= 0.1 Then
        lTickSize = 0.00002
    ElseIf pCurrentPrice <= 0.2 Then
        lTickSize = 0.00005
    ElseIf pCurrentPrice <= 0.5 Then
        lTickSize = 0.0001
    ElseIf pCurrentPrice <= 1 Then
        lTickSize = 0.0002
    ElseIf pCurrentPrice <= 2 Then
        lTickSize = 0.0005
    ElseIf pCurrentPrice <= 5 Then
        lTickSize = 0.001
    ElseIf pCurrentPrice <= 10 Then
        lTickSize = 0.002
    ElseIf pCurrentPrice <= 20 Then
        lTickSize = 0.005
    ElseIf pCurrentPrice <= 50 Then
        lTickSize = 0.01
    ElseIf pCurrentPrice <= 100 Then
        lTickSize = 0.02
    ElseIf pCurrentPrice <= 200 Then
        lTickSize = 0.05
    ElseIf pCurrentPrice <= 500 Then
        lTickSize = 0.1
    Else
        lTickSize = 0.2
    End If
Case "FWB", "IBIS", "BATEDE", "CHIXDE", "EBS", "FWB", "SWB", "TGATE", "TRQXDE", "VSE"
    If pCurrentPrice <= 0.1 Then
        lTickSize = 0.0005
    ElseIf pCurrentPrice <= 0.2 Then
        lTickSize = 0.001
    ElseIf pCurrentPrice <= 0.5 Then
        lTickSize = 0.002
    ElseIf pCurrentPrice <= 1 Then
        lTickSize = 0.005
    ElseIf pCurrentPrice <= 2 Then
        lTickSize = 0.01
    ElseIf pCurrentPrice <= 5 Then
        lTickSize = 0.02
    ElseIf pCurrentPrice <= 10 Then
        lTickSize = 0.05
    ElseIf pCurrentPrice <= 20 Then
        lTickSize = 0.1
    ElseIf pCurrentPrice <= 50 Then
        lTickSize = 0.2
    ElseIf pCurrentPrice <= 100 Then
        lTickSize = 0.5
    ElseIf pCurrentPrice <= 200 Then
        lTickSize = 1
    ElseIf pCurrentPrice <= 500 Then
        lTickSize = 2
    ElseIf pCurrentPrice <= 1000 Then
        lTickSize = 5
    ElseIf pCurrentPrice <= 2000 Then
        lTickSize = 10
    ElseIf pCurrentPrice <= 5000 Then
        lTickSize = 20
    ElseIf pCurrentPrice <= 10000 Then
        lTickSize = 50
    ElseIf pCurrentPrice <= 20000 Then
        lTickSize = 100
    ElseIf pCurrentPrice <= 50000 Then
        lTickSize = 200
    Else
        lTickSize = 500
    End If
Case "VENTURE"
    If pCurrentPrice <= 0.5 Then
        lTickSize = 0.005
    Else
        lTickSize = 0.01
    End If
Case Else
    lTickSize = pDefaultTickSize
End Select

GetTickSizeForStocks = lTickSize
End Function

Private Sub logMessage( _
                ByVal pMsg As String, _
                ByVal pProcName As String, _
                Optional ByVal pMsgQualifier As String = vbNullString, _
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

Dim lPlexId As Long

Dim lOCAGroup As String
If Not pBracketOrder.StopLossOrder Is Nothing Then lOCAGroup = pBracketOrder.StopLossOrder.ProviderProperties.GetParameterValue(ProviderPropertyOCAGroup)
If lOCAGroup = "" And Not pBracketOrder.TargetOrder Is Nothing Then lOCAGroup = pBracketOrder.TargetOrder.ProviderProperties.GetParameterValue(ProviderPropertyOCAGroup)
If lOCAGroup = "" Then
    lOCAGroup = GenerateGUIDString
    If Not pBracketOrder.StopLossOrder Is Nothing Then pBracketOrder.StopLossOrder.ProviderProperties.SetParameterValue ProviderPropertyOCAGroup, lOCAGroup
    If Not pBracketOrder.TargetOrder Is Nothing Then pBracketOrder.TargetOrder.ProviderProperties.SetParameterValue ProviderPropertyOCAGroup, lOCAGroup
End If

If pEntryOrderChanged Then
    placeAnOrder pBracketOrder.EntryOrder, pContract, True, 0, "", OrderModeEntry
End If
If pStopOrderChanged Then
    If pBracketOrder.StopLossOrder.BrokerId = "" Then
        If lPlexId = 0 Then lPlexId = getNextBracketId
        pBracketOrder.StopLossOrder.BrokerId = lPlexId + StopLossOrderIdIncrement
    End If
    placeAnOrder pBracketOrder.StopLossOrder, pContract, True, lParentId, lOCAGroup, OrderModeStopLoss
End If
If pTargetOrderChanged Then
    If pBracketOrder.TargetOrder.BrokerId = "" Then
        If lPlexId = 0 Then lPlexId = getNextBracketId
        pBracketOrder.TargetOrder.BrokerId = lPlexId + TargetOrderIdIncrement
    End If
    placeAnOrder pBracketOrder.TargetOrder, pContract, True, lParentId, lOCAGroup, OrderModeTarget
End If
If pCloseoutOrderChanged Then
    If pBracketOrder.CloseoutOrder.BrokerId = "" Then
        If lPlexId = 0 Then lPlexId = getNextBracketId
        pBracketOrder.CloseoutOrder.BrokerId = getNextBracketId + CloseoutOrderIdIncrement
    End If
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

Private Function orderCanBePlaced(ByVal pOrder As IOrder) As Boolean
If pOrder Is Nothing Then
    orderCanBePlaced = False
ElseIf pOrder.LimitPriceSpec Is Nothing Then
    orderCanBePlaced = Not pOrder.TriggerPriceSpec.RequiresDelayedPriceResolution
ElseIf pOrder.TriggerPriceSpec Is Nothing Then
    orderCanBePlaced = Not pOrder.LimitPriceSpec.RequiresDelayedPriceResolution
Else
    orderCanBePlaced = Not (pOrder.LimitPriceSpec.RequiresDelayedPriceResolution Or _
                            pOrder.TriggerPriceSpec.RequiresDelayedPriceResolution)
End If
End Function

Private Sub placeAnOrder( _
                ByVal pOrder As IOrder, _
                ByVal pContract As TwsContract, _
                ByVal pTransmit As Boolean, _
                ByVal pParentId As Long, _
                ByVal pOcaGroup As String, _
                ByVal pOrderMode As String)
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

mOrderPlacer.PlaceOrder Me, pOrder, pContract, IIf(mStageOrders, False, pTransmit), pParentId, pOcaGroup, mDataSource

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
                ByVal pOrderMode As String)
Const ProcName As String = "placeOrderEx"
On Error GoTo Err

pOrder.BrokerId = CStr(getNextBracketId + EntryOrderIdIncrement)

If pOrder.ContractSpecifier.SecType = SecTypeCombo Then
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
    StartTask(lTask, PriorityNormal, , lCookie).AddTaskCompletionListener Me
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




