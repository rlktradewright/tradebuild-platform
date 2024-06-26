VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimulatedOrderSubmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements IOrderSubmissionListener
Implements IOrderSubmitter

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "SimulatedOrderSubmitter"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mSimulator                                  As MarketSimulator

Private mState                                      As OrderSubmitterStates
Private mStateChangeListeners                       As Listeners

Private mOrderSubmissionListeners                   As Listeners

Private mFactory                                    As SimOrderSubmitterFactory

Private mKey                                        As String

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mContract                                   As IContract

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
Set mStateChangeListeners = New Listeners
Set mOrderSubmissionListeners = New Listeners
setState OrderSubmitterStateNotReady
End Sub

Private Sub Class_Terminate()
Debug.Print "SimulatedOrderSubmitter terminated"
End Sub

'@================================================================================
' IOrderSubmissionListener Interface Members
'@================================================================================

Private Sub IOrderSubmissionListener_NotifyAboutToPlaceOrder(ByVal pOrder As IOrder)
If TypeOf pOrder Is Order Then
    Dim lOrder As Order
    Set lOrder = pOrder
    lOrder.SaveRecoveryInfo
End If
End Sub

Private Sub IOrderSubmissionListener_NotifyError(ByVal pOrderId As String, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IOrderSubmissionListener_NotifyError"
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
        lListener.NotifyError pOrderId, pErrorCode, pErrorMsg
    Next
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmissionListener_NotifyExecutionReport(ByVal pExecutionReport As IExecutionReport)
Const ProcName As String = "IOrderSubmissionListener_NotifyExecutionReport"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmissionListener_NotifyMessage(ByVal pOrderId As String, ByVal pMessage As String)
Const ProcName As String = "IOrderSubmissionListener_NotifyMessage"
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
        lListener.NotifyMessage pOrderId, pMessage
    Next
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmissionListener_NotifyOrderReport(ByVal pOrderReport As IOrderReport)
Const ProcName As String = "IOrderSubmissionListener_NotifyOrderReport"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmissionListener_NotifyOrderStatusReport(ByVal pOrderStatusReport As IOrderStatusReport)
Const ProcName As String = "IOrderSubmissionListener_NotifyOrderStatusReport"
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IOrderSubmitter Interface Members
'@================================================================================

Private Sub IOrderSubmitter_AddOrderSubmissionListener(ByVal pListener As IOrderSubmissionListener)
Const ProcName As String = "IOrderSubmitter_AddOrderSubmissionListener"
On Error GoTo Err

mOrderSubmissionListeners.Add pListener

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_AddStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "IOrderSubmitter_AddStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IOrderSubmitter_CancelBracketOrder(ByVal pBracketOrder As IBracketOrder) As Boolean
Dim EntryOrder As IOrder
Dim Stoporder As IOrder
Dim TargetOrder As IOrder

Const ProcName As String = "IOrderSubmitter_CancelBrackerOrder"
On Error GoTo Err

Set EntryOrder = pBracketOrder.EntryOrder
Select Case EntryOrder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
Case Else
    mSimulator.CancelOrder EntryOrder.Id        ' should automatically cancel the other orders
                                                ' if they have parentid set
    IOrderSubmitter_CancelBracketOrder = True
End Select

Set Stoporder = pBracketOrder.StopLossOrder
If Not Stoporder Is Nothing Then
    Select Case Stoporder.Status
    Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusFilled, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusCancelled
    Case Else
        Select Case EntryOrder.Status
        Case OrderStatuses.OrderStatusFilled, _
            OrderStatuses.OrderStatusCancelling, _
            OrderStatuses.OrderStatusCancelled
            mSimulator.CancelOrder Stoporder.Id
            IOrderSubmitter_CancelBracketOrder = True
        End Select
    End Select
End If

Set TargetOrder = pBracketOrder.TargetOrder
If Not TargetOrder Is Nothing Then
    Select Case TargetOrder.Status
    Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusFilled, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusCancelled
    Case Else
        Select Case EntryOrder.Status
        Case OrderStatuses.OrderStatusFilled, _
            OrderStatuses.OrderStatusCancelling, _
            OrderStatuses.OrderStatusCancelled
            mSimulator.CancelOrder TargetOrder.Id
            IOrderSubmitter_CancelBracketOrder = True
        End Select
    End Select
End If
    
' need some stuff here to cancel if it's an oca group

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName

End Function

Private Sub IOrderSubmitter_CancelCloseoutOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IOrderSubmitter_CancelCloseoutOrder"
On Error GoTo Err

Dim CloseoutOrder As IOrder

Set CloseoutOrder = pBracketOrder.StopLossOrder
Assert Not CloseoutOrder Is Nothing, "bracket order has no closeout order"

Select Case CloseoutOrder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
    Assert True, "stop-loss order state invalid for cancellation"
Case Else
    mSimulator.CancelOrder CloseoutOrder.Id
End Select

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_CancelOrder(ByVal pOrder As IOrder)
Const ProcName As String = "IOrderSubmitter_CancelOrder"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Not implemented"

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_CancelStopOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IOrderSubmitter_CancelStopOrder"
On Error GoTo Err

Dim Stoporder As IOrder

Set Stoporder = pBracketOrder.StopLossOrder
Assert Not Stoporder Is Nothing, "bracket order has no stop-loss order"

Select Case Stoporder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
    Assert True, "stop-loss order state invalid for cancellation"
Case Else
    mSimulator.CancelOrder Stoporder.Id
End Select

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_CancelTargetOrder(ByVal pBracketOrder As IBracketOrder)
Dim TargetOrder As IOrder

Const ProcName As String = "IOrderSubmitter_CancelTargetOrder"
On Error GoTo Err

Set TargetOrder = pBracketOrder.TargetOrder
Assert Not TargetOrder Is Nothing, "bracket order has no target order"


Select Case TargetOrder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
    Assert True, "Target order state invalid for cancellation"
Case Else
    mSimulator.CancelOrder TargetOrder.Id
End Select

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IOrderSubmitter_Capabilities() As OrderSubmitterCapabilities
IOrderSubmitter_Capabilities = OrderSubmitterCapabilityNone
End Property

Private Sub IOrderSubmitter_ExecuteBracketOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IOrderSubmitter_ExecuteBracketOrder"
On Error GoTo Err

Dim lEntryOrder As IOrder: Set lEntryOrder = pBracketOrder.EntryOrder

Dim lStopLossOrder As IOrder: Set lStopLossOrder = pBracketOrder.StopLossOrder
Dim lStopLossOrderCanBePlaced As Boolean: lStopLossOrderCanBePlaced = orderCanBePlaced(lStopLossOrder)

Dim lTargetOrder As IOrder: Set lTargetOrder = pBracketOrder.TargetOrder
Dim lTargetOrderCanBePlaced As Boolean: lTargetOrderCanBePlaced = orderCanBePlaced(lTargetOrder)

If Not lEntryOrder Is Nothing Then
    Dim Transmit As Boolean
    Transmit = True
    If lStopLossOrderCanBePlaced Then Transmit = False
    If lTargetOrderCanBePlaced Then Transmit = False
    PlaceOrder lEntryOrder, "", "", Transmit
    
    Dim OcaGroup As String
    OcaGroup = GenerateTextID

    If lStopLossOrderCanBePlaced Then
        Transmit = True
        If lTargetOrderCanBePlaced Then Transmit = False
        PlaceOrder lStopLossOrder, lEntryOrder.Id, OcaGroup, Transmit
    End If
    
    If lTargetOrderCanBePlaced Then
        PlaceOrder lTargetOrder, lEntryOrder.Id, OcaGroup, True
    End If
    
Else
    ' treat the other orders as an OCA group - still to be implemented
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IOrderSubmitter_FetchPermittedOrderProperties(ByVal pContractFuture As IFuture, Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "IOrderSubmitter_FetchPermittedOrderProperties"
On Error GoTo Err

Set IOrderSubmitter_FetchPermittedOrderProperties = CreateFuture(mSimulator.PermittedOrderProperties, pCookie)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IOrderSubmitter_Finish()
Const ProcName As String = "IOrderSubmitter_Finish"
On Error GoTo Err

mFactory.RemoveSubmitter Me, mSimulator

mSimulator.RemoveOrdersForListener Me
Set mSimulator = Nothing
mOrderSubmissionListeners.Clear
mStateChangeListeners.Clear

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IOrderSubmitter_GetTickSize() As IFuture
Set IOrderSubmitter_GetTickSize = CreateFuture(CreateBoxedValue(mContract.TickSize))
End Function

Private Function IOrderSubmitter_IsReadyForOrderType(ByVal pOrderType As OrderTypes) As Boolean
IOrderSubmitter_IsReadyForOrderType = True
End Function

Private Sub IOrderSubmitter_ModifyBracketOrder( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pEntryOrderChanged As Boolean, _
                ByVal pStopOrderChanged As Boolean, _
                ByVal pTargetOrderChanged As Boolean, _
                ByVal pCloseoutOrderChanged As Boolean)
Const ProcName As String = "IOrderSubmitter_ModifyBracketOrder"
On Error GoTo Err

Dim lOrder As IOrder

Dim lOCAGroup As String
lOCAGroup = GetOcaGroup(pBracketOrder.StopLossOrder)
If lOCAGroup = "" Then lOCAGroup = GetOcaGroup(pBracketOrder.TargetOrder)
If lOCAGroup = "" Then
    lOCAGroup = GenerateTextID
    setOcaGroup pBracketOrder.StopLossOrder, lOCAGroup
    setOcaGroup pBracketOrder.TargetOrder, lOCAGroup
End If

If pEntryOrderChanged Then
    PlaceOrder pBracketOrder.EntryOrder, "", "", True
End If
If pStopOrderChanged Then
    Set lOrder = pBracketOrder.StopLossOrder
    PlaceOrder lOrder, "", lOCAGroup, True
End If
If pTargetOrderChanged Then
    Set lOrder = pBracketOrder.TargetOrder
    PlaceOrder lOrder, "", lOCAGroup, True
End If
If pCloseoutOrderChanged Then
    PlaceOrder pBracketOrder.CloseoutOrder, "", "", True
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_ModifyOrder(ByVal pOrder As IOrder)
Const ProcName As String = "IOrderSubmitter_ModifyOrder"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Not implemented"

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_NotifyRecoveredOrder(ByVal pOrder As IOrder)
Const ProcName As String = "IOrderSubmitter_NotifyRecoveredOrder"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Not implemented"

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let IOrderSubmitter_OrdersAreStaged(ByVal Value As Boolean)
Const ProcName As String = "IOrderSubmitter_OrdersAreStaged"
On Error GoTo Err

Assert Not Value, "Order Staging is not supported", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrderSubmitter_OrdersAreStaged() As Boolean
IOrderSubmitter_OrdersAreStaged = False
End Property

Private Property Get IOrderSubmitter_PermittedOrderTypes() As Long
Dim PermittedOrderTypes As Long
Const ProcName As String = "IOrderSubmitter_PermittedOrderTypes"
On Error GoTo Err

PermittedOrderTypes = OrderTypeLimit Or _
                        OrderTypeMarket Or _
                        OrderTypeStop Or _
                        OrderTypeStopLimit
IOrderSubmitter_PermittedOrderTypes = PermittedOrderTypes

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IOrderSubmitter_PlaceOrder(ByVal pOrder As IOrder)
Const ProcName As String = "IOrderSubmitter_PlaceOrder"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Not implemented"

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_RemoveOrderSubmissionListener(ByVal pListener As IOrderSubmissionListener)
Const ProcName As String = "IOrderSubmitter_RemoveOrderSubmissionListener"
On Error GoTo Err

mOrderSubmissionListeners.Remove pListener

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "IOrderSubmitter_RemoveStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Remove pListener

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_ResubmitStopAndTargetOrders(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IOrderSubmitter_ResubmitStopAndTargetOrders"
On Error GoTo Err


Dim lOCAGroup As String
lOCAGroup = GenerateTextID

Dim lStoporder As IOrder
Set lStoporder = pBracketOrder.StopLossOrder
'lStoporder.BrokerId = ""    ' force a new id to be allocated
PlaceOrder lStoporder, "", lOCAGroup, False

Dim lTargetOrder As IOrder
Set lTargetOrder = pBracketOrder.TargetOrder
'lTargetOrder.BrokerId = "" ' force a new id to be allocated
PlaceOrder lTargetOrder, "", lOCAGroup, True

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_ResubmitStopOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IOrderSubmitter_ResubmitStopOrder"
On Error GoTo Err

Dim lStoporder As IOrder
Set lStoporder = pBracketOrder.StopLossOrder

Dim lTargetOrder As IOrder
Set lTargetOrder = pBracketOrder.TargetOrder
'lStoporder.BrokerId = ""     ' force a new id to be allocated

If Not pBracketOrder.TargetOrder Is Nothing Then
    PlaceOrder lStoporder, "", mSimulator.GetOcaGroup(lTargetOrder.Id), True
Else
    PlaceOrder lStoporder, "", "", True
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_ResubmitTargetOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IOrderSubmitter_ResubmitTargetOrder"
On Error GoTo Err

Dim lStoporder As IOrder
Set lStoporder = pBracketOrder.StopLossOrder

Dim lTargetOrder As IOrder
Set lTargetOrder = pBracketOrder.TargetOrder
'lTargetOrder.BrokerId = ""   ' force a new id to be allocated

If Not pBracketOrder.StopLossOrder Is Nothing Then
    PlaceOrder lTargetOrder, "", mSimulator.GetOcaGroup(lStoporder.Id), True
Else
    PlaceOrder lTargetOrder, "", "", True
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IOrderSubmitter_State() As OrderSubmitterStates
Const ProcName As String = "IOrderSubmitter_State"
On Error GoTo Err

IOrderSubmitter_State = mState

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsAvailable Then Set mContract = ev.Future.Value

Exit Sub

Err:
GOrders.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Key() As String
Key = mKey
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pFactory As SimOrderSubmitterFactory, _
                ByVal pKey As String, _
                ByVal pDataSource As IMarketDataSource)
Set mFactory = pFactory
mKey = pKey
mFutureWaiter.Add pDataSource.ContractFuture
End Sub

Friend Sub SetSimulator( _
                ByVal pSimulator As MarketSimulator)
Const ProcName As String = "SetSimulator"
On Error GoTo Err

Set mSimulator = pSimulator
setState OrderSubmitterStateReady

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function GetOcaGroup(ByVal pOrder As IOrder) As String
If pOrder Is Nothing Then Exit Function
If pOrder.ProviderProperties Is Nothing Then Exit Function
GetOcaGroup = pOrder.ProviderProperties.GetParameterValue(ProviderPropertyOCAGroup)
End Function

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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function orderCanBePlaced(ByVal pOrder As IOrder) As Boolean
If pOrder Is Nothing Then
    orderCanBePlaced = False
Else
    orderCanBePlaced = _
            Not (pOrder.LimitPriceSpec.RequiresDelayedPriceResolution Or _
            pOrder.TriggerPriceSpec.RequiresDelayedPriceResolution)
End If
End Function

Private Sub PlaceOrder( _
                ByVal pOrder As IOrder, _
                ByVal pParentId As String, _
                ByVal pOcaGroup As String, _
                ByVal pTransmit As Boolean)
Const ProcName As String = "placeOrder"
On Error GoTo Err

setOcaGroup pOrder, pOcaGroup
mSimulator.PlaceOrder pOrder, pParentId, pOcaGroup, pTransmit, Me

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setState(ByVal pState As OrderSubmitterStates)
Const ProcName As String = "setState"
On Error GoTo Err

If mState = pState Then Exit Sub
mState = pState
fireStateChange

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub


