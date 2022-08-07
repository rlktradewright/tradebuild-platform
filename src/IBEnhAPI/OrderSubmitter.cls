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

Implements IGenericTickListener
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
Private mContractRequester                          As ContractsTwsRequester
Private mContractCache                              As ContractCache
Private mContractStore                              As IContractStore

Private mState                                      As OrderSubmitterStates
Private mStateChangeListeners                       As Listeners

Private mOrderSubmissionListeners                   As Listeners
Private mCurrentOrderSubmissionListeners()          As Object

Private mStageOrders                                As Boolean

Private mIsTwsConnectedToIBServers                  As Boolean

Private mDataSource                                 As IMarketDataSource

Private mTicksizeRequestPending                     As Boolean

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mContract                                   As IContract

Private mRulesetFutureBuilder                       As FutureBuilder
Private mTicksizeFutureBuilder                      As FutureBuilder

Private mRuleset                                    As PriceIncrementRuleset

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
Set mStateChangeListeners = New Listeners
Set mOrderSubmissionListeners = New Listeners
Set mTicksizeFutureBuilder = New FutureBuilder
End Sub

'================================================================================
' IOrderSubmitter Interface Members
'================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)
End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Dim lPrice As Double: lPrice = getAPrice
If lPrice = 0 Then Exit Sub

mTicksizeFutureBuilder.Value = CreateBoxedValue(mRuleset.GetIncrement(lPrice))
mTicksizeFutureBuilder.Complete
mDataSource.RemoveGenericTickListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Private Sub IOrderSubmitter_CancelCloseoutOrder(ByVal pBracketOrder As OrderUtils27.IBracketOrder)
Const ProcName As String = "IOrderSubmitter_CancelCloseoutOrder"
On Error GoTo Err

CancelCloseoutOrder pBracketOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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

Private Function IOrderSubmitter_GetTickSize() As IFuture
Const ProcName As String = "IOrderSubmitter_GetTickSize"
On Error GoTo Err

Set IOrderSubmitter_GetTickSize = GetTickSize()

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IOrderSubmitter_IsReadyForOrderType(ByVal pOrderType As OrderTypes) As Boolean
IOrderSubmitter_IsReadyForOrderType = IsReadyForOrderType(pOrderType)
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
        placeAnOrder lDeferredAction.Order, lContract.Specifier, 0, lDeferredAction.Transmit, lDeferredAction.ParentId, lDeferredAction.OcaGroup, lDeferredAction.OrderMode, ""
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

If Not ev.Future.IsAvailable Then
ElseIf TypeOf ev.Future.Value Is IContract Then
    Dim lContract As IContract: Set lContract = ev.Future.Value
    If lContract.Specifier.ProviderProperties Is Nothing Then
        ' contract details were probably from a config entry somewhere,
        ' so get the real thing from TWS
        mFutureWaiter.Add FetchContract(lContract.Specifier, mContractStore)
        gLog "Getting contract details from TWSAPI", ModuleName, ProcName
    ElseIf lContract.ProviderProperties.GetParameterValue("MarketRuleID", "0") = "0" Then
        ' ditto
        mFutureWaiter.Add FetchContract(lContract.Specifier, mContractStore)
        gLog "Getting contract details from TWSAPI", ModuleName, ProcName
    Else
        Set mContract = lContract
        If mTicksizeRequestPending Then doGetTickSize
    End If
ElseIf TypeOf ev.Future.Value Is PriceIncrementRuleset Then
    Set mRuleset = ev.Future.Value
    processRuleset
End If

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

Public Sub CancelCloseoutOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "CancelCloseoutOrder"
On Error GoTo Err

Dim lCloseoutOrder As IOrder
Set lCloseoutOrder = pBracketOrder.CloseoutOrder

Assert Not lCloseoutOrder Is Nothing, "Order plex " & pBracketOrder.Key & " has no closeout Order"

Select Case lCloseoutOrder.Status
Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusFilled, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusCancelled
    Assert False, "Closeout Order state invalid for cancellation"
Case Else
    cancelAnOrder lCloseoutOrder, OrderModeCloseout
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
    execBracketOrder pBracketOrder, gContractSpecToTwsContractSpec(pBracketOrder.Contract.Specifier)
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

Dim lFetcher As New PermittedOrderPropsFetcher

lFetcher.Initialise pContractFuture, mContractRequester, mContractCache, pCookie
Set FetchPermittedOrderProperties = lFetcher.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

mStateChangeListeners.Clear
mOrderSubmissionListeners.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetTickSize() As IFuture
Const ProcName As String = "GetTickSize"
On Error GoTo Err

Set GetTickSize = mTicksizeFutureBuilder.Future

If mContract Is Nothing Then
    mTicksizeRequestPending = True
Else
    doGetTickSize
End If
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pTwsAPI As TwsAPI, _
                ByVal pOrderPlacer As OrderPlacer, _
                ByVal pContractStore As IContractStore, _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
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
Set mContractStore = pContractStore
Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
Set mDataSource = pDataSource
mFutureWaiter.Add mDataSource.ContractFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function IsReadyForOrderType(ByVal pOrderType As OrderTypes) As Boolean
Select Case pOrderType
Case OrderTypeTrail, _
        OrderTypeTrailLimit
    IsReadyForOrderType = mDataSource.HasCurrentTick(TickTypeAsk) Or _
                        mDataSource.HasCurrentTick(TickTypeBid) Or _
                        mDataSource.HasCurrentTick(TickTypeTrade)
Case Else
    IsReadyForOrderType = True
End Select
End Function

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
    placeAnOrder pOrder, gContractSpecToTwsContractSpec(pOrder.ContractSpecifier), 0, True, 0, "", OrderModeEntry, ""
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
                        gContractSpecToTwsContractSpec(pBracketOrder.Contract.Specifier), _
                        pEntryOrderChanged, _
                        pStopOrderChanged, _
                        pTargetOrderChanged, _
                        pCloseoutOrderChanged
End If

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyAboutToPlaceOrder( _
                ByVal pOrder As IOrder)
Const ProcName As String = "NotifyAboutToPlaceOrder"
On Error GoTo Err

If getCurrentOrderSubmissionListeners Then
    Dim lListener As IOrderSubmissionListener
    Dim i As Long
    For i = 0 To UBound(mCurrentOrderSubmissionListeners)
        Set lListener = mCurrentOrderSubmissionListeners(i)
        lListener.NotifyAboutToPlaceOrder pOrder
    Next
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

If getCurrentOrderSubmissionListeners Then
    Dim lListener As IOrderSubmissionListener
    Dim i As Long
    For i = 0 To UBound(mCurrentOrderSubmissionListeners)
        Set lListener = mCurrentOrderSubmissionListeners(i)
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

If getCurrentOrderSubmissionListeners Then
    Dim lListener As IOrderSubmissionListener
    Dim i As Long
    For i = 0 To UBound(mCurrentOrderSubmissionListeners)
        Set lListener = mCurrentOrderSubmissionListeners(i)
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

If getCurrentOrderSubmissionListeners Then
    Dim lListener As IOrderSubmissionListener
    Dim i As Long
    For i = 0 To UBound(mCurrentOrderSubmissionListeners)
        Set lListener = mCurrentOrderSubmissionListeners(i)
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

If getCurrentOrderSubmissionListeners Then
    Dim lListener As IOrderSubmissionListener
    Dim i As Long
    For i = 0 To UBound(mCurrentOrderSubmissionListeners)
        Set lListener = mCurrentOrderSubmissionListeners(i)
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

If getCurrentOrderSubmissionListeners Then
    Dim lListener As IOrderSubmissionListener
    Dim i As Long
    For i = 0 To UBound(mCurrentOrderSubmissionListeners)
        Set lListener = mCurrentOrderSubmissionListeners(i)
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

PlaceOrderEx pOrder, 0, True, 0, "", OrderModeEntry

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

Dim lPlexId As Long
lPlexId = getNextBracketId

pBracketOrder.StopLossOrder.BrokerId = lPlexId + StopLossOrderIdIncrement
PlaceOrderEx pBracketOrder.StopLossOrder, 0, True, 0, lOCAGroup, OrderModeStopLoss

pBracketOrder.TargetOrder.BrokerId = lPlexId + TargetOrderIdIncrement
PlaceOrderEx pBracketOrder.TargetOrder, 0, True, 0, lOCAGroup, OrderModeTarget

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ResubmitStopOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "ResubmitStopOrder"
On Error GoTo Err

Dim lPlexId As Long
lPlexId = getNextBracketId
pBracketOrder.StopLossOrder.BrokerId = lPlexId + StopLossOrderIdIncrement

If pBracketOrder.TargetOrder Is Nothing Then
    PlaceOrderEx pBracketOrder.StopLossOrder, 0, True, 0, "", OrderModeStopLoss
Else
    Dim lOCAGroup As String
    lOCAGroup = GenerateGUIDString
    PlaceOrderEx pBracketOrder.StopLossOrder, 0, True, 0, lOCAGroup, OrderModeStopLoss
    
    pBracketOrder.TargetOrder.BrokerId = lPlexId + TargetOrderIdIncrement
    PlaceOrderEx pBracketOrder.TargetOrder, 0, True, 0, lOCAGroup, OrderModeTarget
End If

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ResubmitTargetOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "ResubmitTargetOrder"
On Error GoTo Err

Dim lPlexId As Long
lPlexId = getNextBracketId
pBracketOrder.TargetOrder.BrokerId = lPlexId + TargetOrderIdIncrement

If pBracketOrder.StopLossOrder Is Nothing Then
    PlaceOrderEx pBracketOrder.TargetOrder, 0, True, 0, "", OrderModeTarget
Else
    Dim lOCAGroup As String
    lOCAGroup = GenerateGUIDString
    
    pBracketOrder.StopLossOrder.BrokerId = lPlexId + StopLossOrderIdIncrement
    PlaceOrderEx pBracketOrder.StopLossOrder, 0, True, 0, lOCAGroup, OrderModeStopLoss
    
    PlaceOrderEx pBracketOrder.TargetOrder, 0, True, 0, lOCAGroup, OrderModeTarget
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

Private Sub doGetTickSize()
Const ProcName As String = "doGetTickSize"
On Error GoTo Err

Dim lRulesetID As Long
lRulesetID = CLng(mContract.ProviderProperties.GetParameterValue("MarketRuleID", "0"))
gLog "lRulesetID=" & lRulesetID, ModuleName, ProcName

Dim lPriceMagnifier As Long
lPriceMagnifier = CLng(mContract.ProviderProperties.GetParameterValue("PriceMagnifier", "1"))
gLog "lPriceMagnifier=" & lPriceMagnifier, ModuleName, ProcName

Dim lRulesetFuture As IFuture
Set lRulesetFuture = mContractRequester.RequestPriceIncrementRuleset(lRulesetID, lPriceMagnifier)

If lRulesetFuture.IsAvailable Then
    Set mRuleset = lRulesetFuture.Value
    processRuleset
Else
    mFutureWaiter.Add lRulesetFuture
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub execBracketOrder( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pContractSpec As TwsContractSpecifier)
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
    placeAnOrder lEntryOrder, pContractSpec, 0, Transmit, 0, "", OrderModeEntry, pBracketOrder.Key
    
    If lStopLossOrderCanBePlaced Then
        Transmit = True
        If lTargetOrderCanBePlaced Then Transmit = False
        lStopLossOrder.BrokerId = lPlexId + StopLossOrderIdIncrement
        placeAnOrder lStopLossOrder, pContractSpec, 0, Transmit, lEntryOrder.BrokerId, "", OrderModeStopLoss, pBracketOrder.Key
    End If
    
    If lTargetOrderCanBePlaced Then
        lTargetOrder.BrokerId = lPlexId + TargetOrderIdIncrement
        placeAnOrder lTargetOrder, pContractSpec, 0, True, lEntryOrder.BrokerId, "", OrderModeTarget, pBracketOrder.Key
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

Private Function getAPrice() As Double
Const ProcName As String = "getCurrentPrice"
On Error GoTo Err

Assert mDataSource.State = MarketDataSourceStateRunning, "Data source is not running"
If mDataSource.HasCurrentTick(TickTypeTrade) Then
    getAPrice = mDataSource.CurrentTick(TickTypeTrade).Price
ElseIf mDataSource.HasCurrentTick(TickTypeAsk) Then
    getAPrice = mDataSource.CurrentTick(TickTypeAsk).Price
ElseIf mDataSource.HasCurrentTick(TickTypeBid) Then
    getAPrice = mDataSource.CurrentTick(TickTypeBid).Price
Else
    getAPrice = 0
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCurrentOrderSubmissionListeners() As Boolean
Const ProcName As String = "getCurrentOrderSubmissionListeners"
On Error GoTo Err

Static sInit As Boolean
Static sSomeListeners As Boolean

If Not sInit Or Not mOrderSubmissionListeners.Valid Then
    sInit = True
    sSomeListeners = mOrderSubmissionListeners.GetCurrentListeners(mCurrentOrderSubmissionListeners)
End If
getCurrentOrderSubmissionListeners = sSomeListeners

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

Private Function getOcaGroup(ByVal pOrder As IOrder) As String
If pOrder Is Nothing Then Exit Function
If pOrder.ProviderProperties Is Nothing Then Exit Function
getOcaGroup = pOrder.ProviderProperties.GetParameterValue(ProviderPropertyOCAGroup)
End Function

Private Function getTickSizeFromRuleset( _
                ByVal pRuleset As PriceIncrementRuleset) As Double
Const ProcName As String = "getTickSizeFromRuleset"
On Error GoTo Err

Dim lTickType As TickTypes
Dim lPrice As Double: lPrice = getAPrice()

getTickSizeFromRuleset = pRuleset.GetIncrement(lPrice)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
                ByVal pContractSpec As TwsContractSpecifier, _
                ByVal pEntryOrderChanged As Boolean, _
                ByVal pStopOrderChanged As Boolean, _
                ByVal pTargetOrderChanged As Boolean, _
                ByVal pCloseoutOrderChanged As Boolean)
Const ProcName As String = "modifyABracketOrder"
On Error GoTo Err

Dim lParentId As Long
If pBracketOrder.Size = 0 Then
    lParentId = CLng("0" & pBracketOrder.EntryOrder.BrokerId)
End If

Dim lOCAGroup As String
lOCAGroup = getOcaGroup(pBracketOrder.StopLossOrder)
If lOCAGroup = "" Then lOCAGroup = getOcaGroup(pBracketOrder.TargetOrder)
If lOCAGroup = "" Then
    lOCAGroup = GenerateTextID
    setOcaGroup pBracketOrder.StopLossOrder, lOCAGroup
    setOcaGroup pBracketOrder.TargetOrder, lOCAGroup
End If

If pEntryOrderChanged Then
    placeAnOrder pBracketOrder.EntryOrder, pContractSpec, 0, True, 0, "", OrderModeEntry, pBracketOrder.Key
End If

Dim lPlexId As Long
If pStopOrderChanged Then
    If pBracketOrder.StopLossOrder.BrokerId = "" Then
        lPlexId = getNextBracketId
        pBracketOrder.StopLossOrder.BrokerId = lPlexId + StopLossOrderIdIncrement
    End If
    placeAnOrder pBracketOrder.StopLossOrder, pContractSpec, 0, True, lParentId, lOCAGroup, OrderModeStopLoss, pBracketOrder.Key
End If
If pTargetOrderChanged Then
    If pBracketOrder.TargetOrder.BrokerId = "" Then
        If lPlexId = 0 Then lPlexId = getNextBracketId
        pBracketOrder.TargetOrder.BrokerId = lPlexId + TargetOrderIdIncrement
    End If
    placeAnOrder pBracketOrder.TargetOrder, pContractSpec, 0, True, lParentId, lOCAGroup, OrderModeTarget, pBracketOrder.Key
End If
If pCloseoutOrderChanged Then
    If pBracketOrder.CloseoutOrder.BrokerId = "" Then
        If lPlexId = 0 Then lPlexId = getNextBracketId
        pBracketOrder.CloseoutOrder.BrokerId = getNextBracketId + CloseoutOrderIdIncrement
    End If
    placeAnOrder pBracketOrder.CloseoutOrder, pContractSpec, 0, True, 0, "", OrderModeCloseout, pBracketOrder.Key
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
                ByVal pContractSpec As TwsContractSpecifier, _
                ByVal pLimitPriceOverride, _
                ByVal pTransmit As Boolean, _
                ByVal pParentId As Long, _
                ByVal pOcaGroup As String, _
                ByVal pOrderMode As String, _
                ByVal pBracketOrderKey As String)
Const ProcName As String = "placeAnOrder"
On Error GoTo Err

logMessage "Place " & pOrderMode & " Order for " & pBracketOrderKey & _
                ": broker id=" & pOrder.BrokerId & _
                "; TradeBuild id=" & pOrder.Id, _
            ProcName

Assert pOrder.Id <> "", "Order has no id"
Assert pOrder.Action <> OrderActionNone, "Order has no action"
Assert pOrder.OrderType <> OrderTypeNone, "Order has no order type"
Assert pOrder.Quantity <> 0, "Order has zero quantity"

NotifyAboutToPlaceOrder pOrder
mOrderPlacer.PlaceOrder Me, pOrder, pContractSpec, pLimitPriceOverride, IIf(mStageOrders, False, pTransmit), pParentId, pOcaGroup, mDataSource

If pOrder.Status = OrderStatusCreated Then
    notifyOpenOrder pOrder
    notifyStatus pOrder, OrderStatusPendingSubmit
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub PlaceOrderEx( _
                ByVal pOrder As IOrder, _
                ByVal pLimitPriceOverride, _
                ByVal pTransmit As Boolean, _
                ByVal pParentId As Long, _
                ByVal pOcaGroup As String, _
                ByVal pOrderMode As String)
Const ProcName As String = "PlaceOrderEx"
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
    placeAnOrder pOrder, gContractSpecToTwsContractSpec(pOrder.ContractSpecifier), pLimitPriceOverride, pTransmit, pParentId, pOcaGroup, pOrderMode, ""
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processRuleset()
Const ProcName As String = "processRuleset"
On Error GoTo Err

Dim lTickSize As Double: lTickSize = mRuleset.GetIncrement(0)
If lTickSize <> 0 Then
    mTicksizeFutureBuilder.Value = CreateBoxedValue(lTickSize)
    mTicksizeFutureBuilder.Complete
Else
    Dim lPrice As Double: lPrice = getAPrice
    If lPrice <> 0 Then
        mTicksizeFutureBuilder.Value = CreateBoxedValue(mRuleset.GetIncrement(lPrice))
        mTicksizeFutureBuilder.Complete
    Else
        mDataSource.AddGenericTickListener Me
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOcaGroup( _
                ByVal pOrder As IOrder, _
                ByVal pOcaGroup As String)
Const ProcName As String = "setOcaGroup"
On Error GoTo Err

If pOrder Is Nothing Then Exit Sub
Dim lProps As Parameters: Set lProps = pOrder.ProviderProperties
If lProps Is Nothing Then Set lProps = New Parameters
lProps.SetParameterValue ProviderPropertyOCAGroup, pOcaGroup
pOrder.ProviderProperties = lProps

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




