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
Private mPort As Long
Private mServer As String
Private mClientID As Long

Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean
Private mProviderKey As String

Private mTWSLogLevel As TWSLogLevels

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
    mServiceConsumer.NotifyEvent StandardSPEventCodes.LONotUniqueContract, _
                            "Contract is not uniquely specified", _
                            mOrderSubmissionServiceProvider.handle
                            
End If
End Sub

Private Sub ContractDetailsRequester_contractSpecifierInvalid( _
                ByRef request As TWSContractDetailsRequestQueueEntry, _
                ByVal reason As String)
mServiceConsumer.NotifyEvent StandardSPEventCodes.LOUnknownContract, _
                        "Contract not known", _
                        mOrderSubmissionServiceProvider.handle
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
    mTWSAPI.cancelOrder entryOrder.brokerId    ' should automatically cancel the other orders
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
            mTWSAPI.cancelOrder stopOrder.brokerId
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
            mTWSAPI.cancelOrder targetOrder.brokerId
            IOrderSubmitter_cancelOrderPlex = True
        End Select
    End Select
End If
    
' need some stuff here to cancel if it's an oca group
Exit Function
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmitter" & "::" & "IOrderSubmitter_cancelOrderPlex", _
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
    mTWSAPI.cancelOrder stopOrder.brokerId
End Select
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmitter" & "::" & "IOrderSubmitter_cancelStopOrder", _
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
    mTWSAPI.cancelOrder targetOrder.brokerId
End Select
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmitter" & "::" & "IOrderSubmitter_cancelTargetOrder", _
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
        
    placeOrder entryOrder, transmit, "0", ""
    
    If Not stopOrder Is Nothing Then
        transmit = True
        If Not targetOrder Is Nothing Then
            If targetOrder.orderType <> TradeBuildSP.OrderTypes.OrderTypeAutoLimit Then transmit = False
        End If
        
        If stopOrder.orderType <> TradeBuildSP.OrderTypes.OrderTypeAutoStop Then
            ' note that AUTOSTP orders will be sent when the entry order is filled
            placeOrder stopOrder, transmit, entryOrder.brokerId, ""
        End If
    End If
    
    If Not targetOrder Is Nothing Then
        If targetOrder.orderType <> TradeBuildSP.OrderTypes.OrderTypeAutoLimit Then
            placeOrder targetOrder, True, entryOrder.brokerId, ""
        End If
    End If
    
Else
    ' treat the other orders as an OCA group - still to be implemented
End If
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmitter" & "::" & "IOrderSubmitter_executeOrderPlex", _
                err.Description
End Sub

Private Sub IOrderSubmitter_finish()
finish

mOrderSubmissionServiceProvider.orderSubmitterFinished Me

End Sub

Private Sub IOrderSubmitter_modifyOrderPlex( _
                ByVal pOrderPlex As TradeBuildSP.IOrderPlex, _
                ByVal entryOrderChanged As Boolean, _
                ByVal stopOrderChanged As Boolean, _
                ByVal targetOrderChanged As Boolean, _
                ByVal closeoutOrderChanged As Boolean)

On Error GoTo err
If entryOrderChanged Then
    placeOrder pOrderPlex.entryOrder, True, "0", ""
End If
If stopOrderChanged Then
    placeOrder pOrderPlex.stopOrder, True, pOrderPlex.entryOrder.brokerId, ""
End If
If targetOrderChanged Then
    placeOrder pOrderPlex.targetOrder, True, pOrderPlex.entryOrder.brokerId, ""
End If
If closeoutOrderChanged Then
    placeOrder pOrderPlex.closeoutOrder, True, "0", ""
End If
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmitter" & "::" & "IOrderSubmitter_modifyOrderPlex", _
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
placeOrder stopOrder, False, "0", ocaGroup

Set targetOrder = pOrderPlex.targetOrder
targetOrder.brokerId = "" ' force a new id to be allocated
placeOrder targetOrder, True, "0", ocaGroup
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmitter" & "::" & "IOrderSubmitter_resubmitStopAndTargetOrders", _
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
    placeOrder stopOrder, True, "0", ""
Else
    ocaGroup = generateGUIDString
    placeOrder stopOrder, False, "0", ocaGroup
    placeOrder targetOrder, True, "0", ocaGroup
End If
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmitter" & "::" & "IOrderSubmitter_resubmitStopOrder", _
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
If pOrderPlex.stopOrder Is Nothing Then
    placeOrder targetOrder, True, "0", ""
Else
    ocaGroup = generateGUIDString
    placeOrder stopOrder, False, "0", ocaGroup
    placeOrder targetOrder, True, "0", ocaGroup
End If
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmitter" & "::" & "IOrderSubmitter_resubmitTargetOrder", _
                err.Description
End Sub

'================================================================================
' mTWSAPI Event Handlers
'================================================================================

Private Sub mTWSAPI_Connected()
mServiceConsumer.NotifyEvent StandardSPEventCodes.LOConnectedToBroker, _
                        "Connected to " & mTWSAPI.connectionString, _
                        mOrderSubmissionServiceProvider.handle
If Not mContractWrapper Is Nothing Then mServiceConsumer.Ready
End Sub

Private Sub mTWSAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
On Error GoTo err
If Not retrying Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.LOCantConnectToBroker, _
                            Description, _
                            mOrderSubmissionServiceProvider.handle
Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.LORetryConnectToBroker, _
                            Description, _
                            mOrderSubmissionServiceProvider.handle
End If
mServiceConsumer.NotReady
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmitter" & "::" & "mTWSAPI_ConnectFailed", _
                err.Description
End Sub

Private Sub mTWSAPI_ConnectionToIBClosed()
On Error GoTo err
mServiceConsumer.NotifyEvent StandardSPEventCodes.LOReConnectingToBroker, _
                        "Lost connection from TWS to IB servers", _
                        mOrderSubmissionServiceProvider.handle
mServiceConsumer.NotReady
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmitter" & "::" & "mTWSAPI_ConnectionToIBClosed", _
                err.Description
End Sub

Private Sub mTWSAPI_ConnectionToIBRecovered()
On Error GoTo err
If Not mContractWrapper Is Nothing Then mServiceConsumer.Ready
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmitter" & "::" & "mTWSAPI_ConnectionToIBRecovered", _
                err.Description
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed( _
                ByVal reconnecting As Boolean)
If Not reconnecting Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.LOLostConnectionToBroker, _
                            "Lost connection to TWS", _
                            mOrderSubmissionServiceProvider.handle
Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.LOReConnectingToBroker, _
                            "Lost connection to TWS (reconnecting)", _
                            mOrderSubmissionServiceProvider.handle
End If
mServiceConsumer.NotReady
End Sub

Private Sub mTWSAPI_Disconnected(ByVal reason As String)
On Error GoTo err
mServiceConsumer.NotifyEvent StandardSPEventCodes.LODisconnectedFromBroker, _
                        reason, _
                        mOrderSubmissionServiceProvider.handle
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmitter" & "::" & "mTWSAPI_Disconnected", _
                err.Description
End Sub
'================================================================================
' Properties
'================================================================================

Friend Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Friend Property Let commonServiceConsumer(ByVal RHS As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = RHS
End Property

Friend Property Let connectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
End Property

Friend Property Get connectionRetryIntervalSecs() As Long
connectionRetryIntervalSecs = mConnectionRetryIntervalSecs
End Property

Friend Property Let Contract(ByVal value As ContractUtils26.Contract)
Set mTradeBuildContract = value
End Property

Public Property Let keepConnection(ByVal value As Boolean)
mkeepConnection = value
End Property

Friend Property Let name(ByVal value As String)
mName = value
End Property

Friend Property Let OrderSubmissionServiceProvider( _
                ByVal value As OrderSubmissionSrvcProvider)
Set mOrderSubmissionServiceProvider = value
End Property

Friend Property Let port(ByVal value As Long)
mPort = value
End Property

Friend Property Let providerKey(ByVal value As String)
mProviderKey = value
End Property

Friend Property Let server(ByVal value As String)
mServer = IIf(value = "", "127.0.0.1", value)
End Property

Friend Property Let serviceConsumer(ByVal value As TradeBuildSP.IOrderSubmissionSrvcConsumer)
Set mServiceConsumer = value
End Property

Friend Property Let TWSLogLevel(ByVal value As TWSLogLevels)
mTWSLogLevel = value
End Property

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
If Not mTWSAPI Is Nothing Then gReleaseTWSAPIInstance mTWSAPI
Set mTWSAPI = Nothing
Set mServiceConsumer = Nothing
End Sub

Friend Sub invalidOrder( _
                ByVal orderId As String, _
                ByVal message As String)
                
mServiceConsumer.NotifyEvent StandardSPEventCodes.LOInvalidOrder, _
                        message, _
                        mOrderSubmissionServiceProvider.handle
End Sub

Friend Sub start()
If mTWSAPI Is Nothing Then
    Set mTWSAPI = gGetTWSAPIInstance(mServer, mPort, mClientID, mProviderKey, mConnectionRetryIntervalSecs, mkeepConnection, mTWSLogLevel)
    If mTWSAPI.connectionState = ConnConnected Then
        mServiceConsumer.NotifyEvent StandardSPEventCodes.LOConnectedToBroker, _
                                "Connected to " & mTWSAPI.connectionString, _
                                mOrderSubmissionServiceProvider.handle
    End If
End If


mTWSAPI.RequestContract mTradeBuildContract.specifier, Nothing, Me
End Sub

Friend Sub updateStatus( _
                ByVal orderId As String, _
                ByVal status As TradeBuildSP.OrderStatuses)

mServiceConsumer.orderStatus orderId, status
End Sub
'================================================================================
' Helper Functions
'================================================================================

Private Sub handleFatalError( _
                ByVal number As Long, _
                ByVal source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError number, _
                                source, _
                                Description, _
                                mOrderSubmissionServiceProvider.handle

finish
mOrderSubmissionServiceProvider.orderSubmitterFinished Me
End Sub

Private Sub placeOrder( _
                ByVal pOrder As TradeBuildSP.IOrder, _
                ByVal transmit As Boolean, _
                ByVal parentId As String, _
                ByVal ocaGroup As String)
mTWSAPI.placeOrder Me, pOrder, transmit, parentId, ocaGroup, ""
End Sub



