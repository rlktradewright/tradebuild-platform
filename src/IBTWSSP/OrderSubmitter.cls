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
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements ContractDetailsRequester
Implements TradeBuildSP.IOrderSubmitter

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

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mOrderSubmissionServiceProvider As OrderSubmissionSrvcProvider

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mServiceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer

Private mTradeBuildContract As Contract
Private mContractWrapper As ContractWrapper

Private mName As String

Private WithEvents mTWSAPI As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1

Private mHandle As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
finish
Debug.Print "IBTWSSP.OrderSubmitter terminated"
End Sub

'================================================================================
' ContractDetailsRequester Interface Members
'================================================================================

Private Sub ContractDetailsRequester_cancelRequest( _
                ByRef request As TWSContractDetailsRequestQueueEntry, _
                ByVal reason As String)

End Sub

Private Sub ContractDetailsRequester_contractsLoaded( _
                ByRef request As TWSContractDetailsRequestQueueEntry)
If request.contractWrappers.Count = 1 Then
    Set mContractWrapper = request.contractWrappers(1)
    mServiceConsumer.Ready
Else
    mServiceConsumer.notifyEvent StandardSPEventCodes.LONotUniqueContract, _
                            "Contract is not uniquely specified", _
                            mHandle
                            
End If
End Sub

Private Sub ContractDetailsRequester_contractSpecifierInvalid( _
                ByRef request As TWSContractDetailsRequestQueueEntry, _
                ByVal reason As String)
mServiceConsumer.notifyEvent StandardSPEventCodes.LOUnknownContract, _
                        "Contract not known", _
                        mHandle
End Sub

'================================================================================
' IOrderSubmitter Interface Members
'================================================================================

Private Function IOrderSubmitter_cancelOrderPlex(ByVal pOrderPlex As TradeBuildSP.IOrderPlex) As Boolean
Dim entryOrder As TradeBuildSP.IOrder
Dim stopOrder As TradeBuildSP.IOrder
Dim targetOrder As TradeBuildSP.IOrder

On Error GoTo err
Set entryOrder = pOrderPlex.entryOrder
Select Case entryOrder.status
Case TradeBuildSP.OrderStatuses.OrderStatusCreated, _
    TradeBuildSP.OrderStatuses.OrderStatusFilled, _
    TradeBuildSP.OrderStatuses.OrderStatusCancelling, _
    TradeBuildSP.OrderStatuses.OrderStatusCancelled
Case Else
    cancelOrder entryOrder, OrderModeEntry    ' should automatically cancel the other orders
                                                    ' if they have parentid set
    IOrderSubmitter_cancelOrderPlex = True
End Select

Set stopOrder = pOrderPlex.stopOrder
If Not stopOrder Is Nothing Then
    Select Case stopOrder.status
    Case TradeBuildSP.OrderStatuses.OrderStatusCreated, _
        TradeBuildSP.OrderStatuses.OrderStatusFilled, _
        TradeBuildSP.OrderStatuses.OrderStatusCancelling, _
        TradeBuildSP.OrderStatuses.OrderStatusCancelled
    Case Else
        Select Case entryOrder.status
        Case TradeBuildSP.OrderStatuses.OrderStatusFilled, _
            TradeBuildSP.OrderStatuses.OrderStatusCancelling, _
            TradeBuildSP.OrderStatuses.OrderStatusCancelled
            cancelOrder stopOrder, OrderModeStop
            IOrderSubmitter_cancelOrderPlex = True
        End Select
    End Select
End If

Set targetOrder = pOrderPlex.targetOrder
If Not targetOrder Is Nothing Then
    Select Case targetOrder.status
    Case TradeBuildSP.OrderStatuses.OrderStatusCreated, _
        TradeBuildSP.OrderStatuses.OrderStatusFilled, _
        TradeBuildSP.OrderStatuses.OrderStatusCancelling, _
        TradeBuildSP.OrderStatuses.OrderStatusCancelled
    Case Else
        Select Case entryOrder.status
        Case TradeBuildSP.OrderStatuses.OrderStatusFilled, _
            TradeBuildSP.OrderStatuses.OrderStatusCancelling, _
            TradeBuildSP.OrderStatuses.OrderStatusCancelled
            cancelOrder targetOrder, OrderModeTarget
            IOrderSubmitter_cancelOrderPlex = True
        End Select
    End Select
End If
    
' need some stuff here to cancel if it's an oca group
Exit Function
err:
handleFatalError err.number, _
                "IOrderSubmitter_cancelOrderPlex", _
                err.Description

End Function

Private Sub IOrderSubmitter_cancelStopOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim stopOrder As TradeBuildSP.IOrder

On Error GoTo err
Set stopOrder = pOrderPlex.stopOrder
If stopOrder Is Nothing Then err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "IBTWSSP.OrderSubmitter::cancelStopOrder", _
                                    "Order plex " & pOrderPlex.key & " has no stop order"

Select Case stopOrder.status
Case TradeBuildSP.OrderStatuses.OrderStatusCreated, _
    TradeBuildSP.OrderStatuses.OrderStatusFilled, _
    TradeBuildSP.OrderStatuses.OrderStatusCancelling, _
    TradeBuildSP.OrderStatuses.OrderStatusCancelled
    err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "IBTWSSP.OrderSubmitter::cancelStopOrder", _
                                    "Stop order state invalid for cancellation"
Case Else
    cancelOrder stopOrder, OrderModeStop
End Select
Exit Sub
err:
handleFatalError err.number, _
                "IOrderSubmitter_cancelStopOrder", _
                err.Description
End Sub

Private Sub IOrderSubmitter_cancelTargetOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim targetOrder As TradeBuildSP.IOrder

On Error GoTo err
Set targetOrder = pOrderPlex.targetOrder
If targetOrder Is Nothing Then err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "IBTWSSP.OrderSubmitter::cancelTargetOrder", _
                                    "Order plex " & pOrderPlex.key & " has no target order"


Select Case targetOrder.status
Case TradeBuildSP.OrderStatuses.OrderStatusCreated, _
    TradeBuildSP.OrderStatuses.OrderStatusCancelling, _
    TradeBuildSP.OrderStatuses.OrderStatusCancelled
    err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "IBTWSSP.OrderSubmitter::cancelTargetOrder", _
                                    "Target order state invalid for cancellation"
Case Else
    cancelOrder targetOrder, OrderModeTarget
End Select
Exit Sub
err:
handleFatalError err.number, _
                "IOrderSubmitter_cancelTargetOrder", _
                err.Description
End Sub

Private Property Let IOrderSubmitter_clockRate(ByVal RHS As Single)

End Property

Private Sub IOrderSubmitter_executeOrderPlex(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim entryOrder As TradeBuildSP.IOrder
Dim stopOrder As TradeBuildSP.IOrder
Dim targetOrder As TradeBuildSP.IOrder
Dim transmit As Boolean

On Error GoTo err
Set entryOrder = pOrderPlex.entryOrder
Set stopOrder = pOrderPlex.stopOrder
Set targetOrder = pOrderPlex.targetOrder

If Not entryOrder Is Nothing Then
    transmit = True
    If Not stopOrder Is Nothing Then
        If stopOrder.orderType <> TradeBuildSP.OrderTypes.OrderTypeAutoStop Then transmit = False
    End If
    If Not targetOrder Is Nothing Then
        If targetOrder.orderType <> TradeBuildSP.OrderTypes.OrderTypeAutoLimit Then transmit = False
    End If
        
    placeOrder entryOrder, transmit, "0", "", OrderModeEntry
    
    If Not stopOrder Is Nothing Then
        transmit = True
        If Not targetOrder Is Nothing Then
            If targetOrder.orderType <> TradeBuildSP.OrderTypes.OrderTypeAutoLimit Then transmit = False
        End If
        
        If stopOrder.orderType <> TradeBuildSP.OrderTypes.OrderTypeAutoStop Then
            ' note that AUTOSTP orders will be sent when the entry order is filled
            placeOrder stopOrder, transmit, entryOrder.brokerId, "", OrderModeStop
        End If
    End If
    
    If Not targetOrder Is Nothing Then
        If targetOrder.orderType <> TradeBuildSP.OrderTypes.OrderTypeAutoLimit Then
            placeOrder targetOrder, True, entryOrder.brokerId, "", OrderModeTarget
        End If
    End If
    
Else
    ' treat the other orders as an OCA group - still to be implemented
End If
Exit Sub
err:
handleFatalError err.number, _
                "IOrderSubmitter_executeOrderPlex", _
                err.Description
End Sub

Private Sub IOrderSubmitter_finish()

mOrderSubmissionServiceProvider.orderSubmitterFinished Me
finish

End Sub

Private Sub IOrderSubmitter_modifyOrderPlex( _
                ByVal pOrderPlex As TradeBuildSP.IOrderPlex, _
                ByVal entryOrderChanged As Boolean, _
                ByVal stopOrderChanged As Boolean, _
                ByVal targetOrderChanged As Boolean, _
                ByVal closeoutOrderChanged As Boolean)

Dim ocaGroup As String
Dim parentId As String

On Error GoTo err

Select Case pOrderPlex.entryOrder.status
Case OrderStatusFilled, _
        OrderStatusCancelling, _
        OrderStatusCancelled
    parentId = "0"
Case Else
    parentId = pOrderPlex.entryOrder.brokerId
End Select

If pOrderPlex.entryOrder.quantityRemaining = 0 Then parentId = "0"

If Not pOrderPlex.stopOrder Is Nothing Then ocaGroup = pOrderPlex.stopOrder.providerProperties.getParameterValue("OCA group")
If ocaGroup = "" And Not pOrderPlex.targetOrder Is Nothing Then ocaGroup = pOrderPlex.targetOrder.providerProperties.getParameterValue("OCA group")
If ocaGroup = "" Then ocaGroup = generateGUIDString

If entryOrderChanged Then
    placeOrder pOrderPlex.entryOrder, True, "0", "", OrderModeEntry
End If
If stopOrderChanged Then
    placeOrder pOrderPlex.stopOrder, True, parentId, ocaGroup, OrderModeStop
End If
If targetOrderChanged Then
    placeOrder pOrderPlex.targetOrder, True, parentId, ocaGroup, OrderModeTarget
End If
If closeoutOrderChanged Then
    placeOrder pOrderPlex.closeoutOrder, True, "0", "", OrderModeCloseout
End If
Exit Sub
err:
handleFatalError err.number, _
                "IOrderSubmitter_modifyOrderPlex", _
                err.Description
    
End Sub

Private Property Get IOrderSubmitter_permittedOrderAttributes() As Long
IOrderSubmitter_permittedOrderAttributes = mContractWrapper.permittedOrderAttributes
End Property

Private Property Get IOrderSubmitter_permittedOrderTifs() As Long
IOrderSubmitter_permittedOrderTifs = mContractWrapper.permittedOrderTifs
End Property

Private Property Get IOrderSubmitter_permittedOrderTypes() As Long
IOrderSubmitter_permittedOrderTypes = mContractWrapper.permittedOrderTypes
End Property

Private Property Get IOrderSubmitter_permittedStopTriggerMethods() As Long
IOrderSubmitter_permittedStopTriggerMethods = mContractWrapper.permittedStopTriggerMethods
End Property

Private Sub IOrderSubmitter_resubmitStopAndTargetOrders(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim stopOrder As TradeBuildSP.IOrder
Dim targetOrder As TradeBuildSP.IOrder
Dim ocaGroup As String

On Error GoTo err
ocaGroup = generateGUIDString

Set stopOrder = pOrderPlex.stopOrder
stopOrder.brokerId = ""   ' force a new id to be allocated
placeOrder stopOrder, False, "0", ocaGroup, OrderModeStop

Set targetOrder = pOrderPlex.targetOrder
targetOrder.brokerId = "" ' force a new id to be allocated
placeOrder targetOrder, True, "0", ocaGroup, OrderModeTarget
Exit Sub
err:
handleFatalError err.number, _
                "IOrderSubmitter_resubmitStopAndTargetOrders", _
                err.Description
End Sub

Private Sub IOrderSubmitter_resubmitStopOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim stopOrder As TradeBuildSP.IOrder
Dim targetOrder As TradeBuildSP.IOrder
Dim ocaGroup As String

On Error GoTo err
Set stopOrder = pOrderPlex.stopOrder
Set targetOrder = pOrderPlex.targetOrder
stopOrder.brokerId = ""     ' force a new id to be allocated
If targetOrder Is Nothing Then
    placeOrder stopOrder, True, "0", "", OrderModeStop
Else
    ocaGroup = generateGUIDString
    placeOrder stopOrder, False, "0", ocaGroup, OrderModeStop
    placeOrder targetOrder, True, "0", ocaGroup, OrderModeTarget
End If
Exit Sub
err:
handleFatalError err.number, _
                "IOrderSubmitter_resubmitStopOrder", _
                err.Description
End Sub

Private Sub IOrderSubmitter_resubmitTargetOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim stopOrder As TradeBuildSP.IOrder
Dim targetOrder As TradeBuildSP.IOrder
Dim ocaGroup As String

On Error GoTo err
Set stopOrder = pOrderPlex.stopOrder
Set targetOrder = pOrderPlex.targetOrder
targetOrder.brokerId = ""   ' force a new id to be allocated
If stopOrder Is Nothing Then
    placeOrder targetOrder, True, "0", "", OrderModeTarget
Else
    ocaGroup = generateGUIDString
    placeOrder stopOrder, False, "0", ocaGroup, OrderModeStop
    placeOrder targetOrder, True, "0", ocaGroup, OrderModeTarget
End If
Exit Sub
err:
handleFatalError err.number, _
                "IOrderSubmitter_resubmitTargetOrder", _
                err.Description
End Sub

'================================================================================
' mTWSAPI Event Handlers
'================================================================================

Private Sub mTWSAPI_Connected()
On Error GoTo err

If Not mContractWrapper Is Nothing Then mServiceConsumer.Ready

Exit Sub

err:
handleFatalError err.number, _
                "mTWSAPI_Connected", _
                err.Description
End Sub

Private Sub mTWSAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
On Error GoTo err

mServiceConsumer.NotReady
Exit Sub
err:
handleFatalError err.number, _
                "mTWSAPI_ConnectFailed", _
                err.Description
End Sub

Private Sub mTWSAPI_ConnectionToIBClosed()
On Error GoTo err

mServiceConsumer.NotReady
Exit Sub
err:
handleFatalError err.number, _
                "mTWSAPI_ConnectionToIBClosed", _
                err.Description
End Sub

Private Sub mTWSAPI_ConnectionToIBRecovered()
On Error GoTo err
If Not mContractWrapper Is Nothing Then mServiceConsumer.Ready
Exit Sub
err:
handleFatalError err.number, _
                "mTWSAPI_ConnectionToIBRecovered", _
                err.Description
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed( _
                ByVal reconnecting As Boolean)
On Error GoTo err

mServiceConsumer.NotReady
Exit Sub
err:
handleFatalError err.number, _
                "mTWSAPI_ConnectionToTWSClosed", _
                err.Description
End Sub

Private Sub mTWSAPI_Disconnected(ByVal reason As String)
On Error GoTo err

mServiceConsumer.NotReady
Exit Sub
err:
handleFatalError err.number, _
                "mTWSAPI_Disconnected", _
                err.Description
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub executionDetails( _
                ByVal pContractSpecifier As ContractSpecifier, _
                ByVal AccountIdentifier As String, _
                ByVal Action As TradeBuildSP.OrderActions, _
                ByVal execId As String, _
                ByVal FillingExchange As String, _
                ByVal IsLiquidation As Boolean, _
                ByVal OrderBrokerId As String, _
                ByVal OrderTradeBuildId As String, _
                ByVal Price As Double, _
                ByVal Quantity As Long, _
                ByVal fillTime As Date)
mServiceConsumer.executionDetails pContractSpecifier, _
                                AccountIdentifier, _
                                Action, _
                                execId, _
                                FillingExchange, _
                                IsLiquidation, _
                                OrderBrokerId, _
                                OrderTradeBuildId, _
                                Price, _
                                Quantity, _
                                fillTime
End Sub

Friend Sub finish()
Set mOrderSubmissionServiceProvider = Nothing
Set mCommonServiceConsumer = Nothing
Set mServiceConsumer = Nothing
Set mTradeBuildContract = Nothing
Set mTWSAPI = Nothing
End Sub

Friend Sub Initialise( _
                ByVal pOrderSubmissionServiceProvider As IOrderSubmissionSrvcProvider, _
                ByVal commonServiceConsumer As ICommonServiceConsumer, _
                ByVal orderSubmissionServiceConsumer As IOrderSubmissionSrvcConsumer, _
                ByVal pContract As Contract, _
                ByVal pTWSAPI As TWSAPI, _
                ByVal handle As Long)
Set mOrderSubmissionServiceProvider = pOrderSubmissionServiceProvider
Set mCommonServiceConsumer = commonServiceConsumer
Set mServiceConsumer = orderSubmissionServiceConsumer
Set mTradeBuildContract = pContract
Set mTWSAPI = pTWSAPI
mHandle = handle

mTWSAPI.RequestContract mTradeBuildContract.specifier, Me
End Sub

Friend Sub invalidOrder( _
                ByVal orderId As String, _
                ByVal message As String)
                
mServiceConsumer.RejectOrder orderId, _
                        message
End Sub

Friend Sub orderReport( _
                ByVal lOrderReport As orderReport)
mServiceConsumer.orderReport lOrderReport
End Sub

Friend Sub updateStatus( _
                ByVal orderId As String, _
                ByVal status As TradeBuildSP.OrderStatuses)

mServiceConsumer.OrderStatus orderId, status
End Sub
'================================================================================
' Helper Functions
'================================================================================

Private Sub cancelOrder( _
                ByVal pOrder As TradeBuildSP.IOrder, _
                ByRef orderMode As String)
gLogger.Log LogLevelNormal, _
            "Cancel " & orderMode & " order: " & _
            "broker id=" & pOrder.brokerId & _
            "; TradeBuild id=" & pOrder.tradeBuildId
mTWSAPI.cancelOrder pOrder.brokerId
End Sub

Private Sub handleFatalError( _
                ByVal number As Long, _
                ByVal source As String, _
                ByVal Description As String)

logMessage "Error " & err.number & ": " & _
            ProjectName & "." & ModuleName & ":" & source & vbCrLf & _
            err.Description, _
            LogLevelSevere

gReleaseAllTWSAPIInstances

mCommonServiceConsumer.NotifyFatalError number, _
                            source, _
                            Description, _
                            mHandle

End Sub

Private Sub logMessage( _
                ByVal message As String, _
                Optional ByVal logLevel As LogLevels = LogLevelNormal)
If Not gLogger.isLoggable(logLevel) Then Exit Sub
gLogger.Log logLevel, _
            ProjectName & "." & ModuleName & ": " & message
End Sub

Private Sub placeOrder( _
                ByVal pOrder As TradeBuildSP.IOrder, _
                ByVal transmit As Boolean, _
                ByVal parentId As String, _
                ByVal ocaGroup As String, _
                ByRef orderMode As String)
gLogger.Log LogLevelNormal, _
            "Place " & orderMode & " order: " & _
            "broker id=" & pOrder.brokerId & _
            "; TradeBuild id=" & pOrder.tradeBuildId
mTWSAPI.placeOrder Me, pOrder, transmit, parentId, ocaGroup, ""
If pOrder.status = OrderStatusCreated Then mServiceConsumer.OrderStatus pOrder.tradeBuildId, OrderStatusPendingSubmit
End Sub



