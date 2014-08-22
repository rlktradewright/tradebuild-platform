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

Private Const ModuleName                    As String = "SimulatedOrderSubmitter"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mSimulator                          As MarketSimulator

Private mState                              As OrderSubmitterStates
Private mStateChangeListeners               As Listeners

Private mOrderSubmissionListeners           As Listeners

Private mFactory                            As SimOrderSubmitterFactory

Private mKey                                        As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
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

Private Sub IOrderSubmissionListener_NotifyError(ByVal pOrderId As String, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IOrderSubmissionListener_NotifyError"
On Error GoTo Err

Dim lListener As IOrderSubmissionListener
For Each lListener In mOrderSubmissionListeners.CurrentListeners
    lListener.NotifyError pOrderId, pErrorCode, pErrorMsg
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmissionListener_NotifyExecutionReport(ByVal pExecutionReport As IExecutionReport)
Const ProcName As String = "IOrderSubmissionListener_NotifyExecutionReport"
On Error GoTo Err

Dim lListener As IOrderSubmissionListener
For Each lListener In mOrderSubmissionListeners.CurrentListeners
    lListener.NotifyExecutionReport pExecutionReport
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmissionListener_NotifyOrderReport(ByVal pOrderReport As IOrderReport)
Const ProcName As String = "IOrderSubmissionListener_NotifyOrderReport"
On Error GoTo Err

Dim lListener As IOrderSubmissionListener
For Each lListener In mOrderSubmissionListeners.CurrentListeners
    lListener.NotifyOrderReport pOrderReport
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmissionListener_NotifyOrderStatusReport(ByVal pOrderStatusReport As IOrderStatusReport)
Const ProcName As String = "IOrderSubmissionListener_NotifyOrderStatusReport"
On Error GoTo Err

Dim lListener As IOrderSubmissionListener
For Each lListener In mOrderSubmissionListeners.CurrentListeners
    lListener.NotifyOrderStatusReport pOrderStatusReport
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_AddStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "IOrderSubmitter_AddStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Sub IOrderSubmitter_CancelStopOrder(ByVal pBracketOrder As IBracketOrder)
Dim Stoporder As IOrder

Const ProcName As String = "IOrderSubmitter_CancelStopOrder"
On Error GoTo Err

Set Stoporder = pBracketOrder.StopLossOrder
Assert Not Stoporder Is Nothing, "bracket order has no stop order"

Select Case Stoporder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
    Assert True, "Stop order state invalid for cancellation"
Case Else
    mSimulator.CancelOrder Stoporder.Id
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_ExecuteBracketOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IOrderSubmitter_ExecuteBracketOrder"
On Error GoTo Err

Dim EntryOrder As IOrder
Set EntryOrder = pBracketOrder.EntryOrder

Dim Stoporder As IOrder
Set Stoporder = pBracketOrder.StopLossOrder

Dim TargetOrder As IOrder
Set TargetOrder = pBracketOrder.TargetOrder

If Not EntryOrder Is Nothing Then
    Dim Transmit As Boolean
    Transmit = True
    If Not Stoporder Is Nothing Then
        If Stoporder.OrderType <> OrderTypes.OrderTypeAutoStop Then Transmit = False
    End If
    If Not TargetOrder Is Nothing Then
        If TargetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then Transmit = False
    End If
        
    PlaceOrder EntryOrder, "", "", Transmit
    
    Dim OcaGroup As String
    OcaGroup = GenerateTextID

    If Not Stoporder Is Nothing Then
        Transmit = True
        If Not TargetOrder Is Nothing Then
            If TargetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then Transmit = False
        End If
        
        If Stoporder.OrderType <> OrderTypes.OrderTypeAutoStop Then
            ' note that AUTOSTP orders will be sent when the entry order is filled
            PlaceOrder Stoporder, EntryOrder.Id, OcaGroup, Transmit
        End If
    End If
    
    If Not TargetOrder Is Nothing Then
        If TargetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then
            PlaceOrder TargetOrder, EntryOrder.Id, OcaGroup, True
        End If
    End If
    
Else
    ' treat the other orders as an OCA group - still to be implemented
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IOrderSubmitter_FetchPermittedOrderProperties(ByVal pContractFuture As IFuture, Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "IOrderSubmitter_FetchPermittedOrderProperties"
On Error GoTo Err

Set IOrderSubmitter_FetchPermittedOrderProperties = CreateFuture(mSimulator.PermittedOrderProperties, pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_ModifyBracketOrder( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal entryOrderChanged As Boolean, _
                ByVal stopOrderChanged As Boolean, _
                ByVal targetOrderChanged As Boolean, _
                ByVal closeoutOrderChanged As Boolean)
Const ProcName As String = "IOrderSubmitter_ModifyBracketOrder"
On Error GoTo Err

Dim lOrder As IOrder

If entryOrderChanged Then
    PlaceOrder pBracketOrder.EntryOrder, "", "", True
End If
If stopOrderChanged Then
    Set lOrder = pBracketOrder.StopLossOrder
    PlaceOrder lOrder, "", "", True
End If
If targetOrderChanged Then
    Set lOrder = pBracketOrder.TargetOrder
    PlaceOrder lOrder, "", "", True
End If
If closeoutOrderChanged Then
    PlaceOrder pBracketOrder.CloseoutOrder, "", "", True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_ModifyOrder(ByVal pOrder As IOrder)
Const ProcName As String = "IOrderSubmitter_ModifyOrder"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Not implemented"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_NotifyRecoveredOrder(ByVal pOrder As IOrder)
Const ProcName As String = "IOrderSubmitter_NotifyRecoveredOrder"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Not implemented"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IOrderSubmitter_PlaceOrder(ByVal pOrder As IOrder)
Const ProcName As String = "IOrderSubmitter_PlaceOrder"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Not implemented"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_RemoveOrderSubmissionListener(ByVal pListener As IOrderSubmissionListener)
Const ProcName As String = "IOrderSubmitter_RemoveOrderSubmissionListener"
On Error GoTo Err

mOrderSubmissionListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_RemoveStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "IOrderSubmitter_RemoveStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmitter_ResubmitStopAndTargetOrders(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IOrderSubmitter_ResubmitStopAndTargetOrders"
On Error GoTo Err


Dim lOcaGroup As String
lOcaGroup = GenerateTextID

Dim lStoporder As IOrder
Set lStoporder = pBracketOrder.StopLossOrder
'lStoporder.BrokerId = ""    ' force a new id to be allocated
PlaceOrder lStoporder, "", lOcaGroup, False

Dim lTargetOrder As IOrder
Set lTargetOrder = pBracketOrder.TargetOrder
'lTargetOrder.BrokerId = "" ' force a new id to be allocated
PlaceOrder lTargetOrder, "", lOcaGroup, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IOrderSubmitter_State() As OrderSubmitterStates
Const ProcName As String = "IOrderSubmitter_State"
On Error GoTo Err

IOrderSubmitter_State = mState

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Key() As String
Key = mKey
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ByVal pFactory As SimOrderSubmitterFactory, ByVal pKey As String)
Set mFactory = pFactory
mKey = pKey
End Sub

Friend Sub SetSimulator( _
                ByVal pSimulator As MarketSimulator)
Const ProcName As String = "SetSimulator"
On Error GoTo Err

Set mSimulator = pSimulator
setState OrderSubmitterStateReady

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireStateChange()
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = mState

Dim lListener As StateChangeListener
For Each lListener In mStateChangeListeners.CurrentListeners
    lListener.Change ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub PlaceOrder( _
                ByVal pOrder As IOrder, _
                ByVal pParentId As String, _
                ByVal pOcaGroup As String, _
                ByVal pTransmit As Boolean)
Const ProcName As String = "placeOrder"
On Error GoTo Err

mSimulator.PlaceOrder pOrder, pParentId, pOcaGroup, pTransmit, Me

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


