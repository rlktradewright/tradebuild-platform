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

Implements IContractDetailsConsumer
Implements IOrderSubmitter

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
Private Const StopOrderIdIncrement      As Long = 0
Private Const TargetOrderIdIncrement    As Long = 0

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
Private mOrderSubmissionListener                    As IOrderSubmissionListener

Private mContractDetails                            As TwsContractDetails

Private WithEvents mTwsAPI                          As TwsAPI
Attribute mTwsAPI.VB_VarHelpID = -1
Private mContractRequester                          As ContractDetailsRequester

Private mState                                      As OrderSubmitterStates
Private mStateChangeListeners                       As EnumerableCollection

Private mPermittedOrderTypes                        As OrderTypes
Private mPermittedOrderTifs                         As OrderTIFs
Private mPermittedOrderAttributes                   As OrderAttributes
Private mPermittedStopTriggerMethods                As StopTriggerMethods

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mStateChangeListeners = New EnumerableCollection
End Sub

Private Sub Class_Terminate()
Finish
Debug.Print "IBTwsSP.OrderSubmitter terminated"
End Sub

'================================================================================
' IContractDetailsConsumer Interface Members
'================================================================================

Private Sub IContractDetailsConsumer_EndContractDetails(ByVal pRequestId As Long)
Const ProcName As String = "IContractDetailsConsumer_EndContractDetails"
On Error GoTo Err

Dim i As Long
Dim lOrderTypesStr() As String
Dim lOrderType As OrderTypes

If mContractDetails Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Failed to retrieve contract details"

lOrderTypesStr = Split(mContractDetails.OrderTypes, ",")

mPermittedOrderAttributes = OrderAttributes.OrderAttBlockOrder
mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttMinimumQuantity
mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttNBBOPriceCap
mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttOriginatorRef
mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttOverrideConstraints
mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttPercentOffset
mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttRule80A
mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttAction
mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttLimitPrice
mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttOrderType
mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttQuantity
mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttTimeInForce
mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttTriggerPrice
mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttStopTriggerMethod

For i = 0 To UBound(lOrderTypesStr)
    lOrderType = gTwsOrderTypeFromString(lOrderTypesStr(i))
    
    If lOrderType <> OrderTypeNone Then
        mPermittedOrderTypes = mPermittedOrderTypes Or lOrderType
    Else
        Select Case lOrderTypesStr(i)
        Case "AON"  ' all or none
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttAllOrNone
        Case "DAY"  ' today only
            mPermittedOrderTifs = mPermittedOrderTifs Or OrderTIFs.OrderTIFDay
        Case "DIS"  ' discretionary amount
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttDiscretionaryAmount
        Case "ELECONLY"  ' electronic trade only
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttETradeOnly
        Case "FIRMONLY"  ' firm quote only
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttFirmQuoteOnly
        Case "GAT"  ' good after Time
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttGoodAfterTime
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttGoodAfterTimeTZ
        Case "GTC"  ' good till cancelled
            mPermittedOrderTifs = mPermittedOrderTifs Or OrderTIFs.OrderTIFGoodTillCancelled
        Case "GTD"  ' good till date
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttGoodTillDate
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttGoodTillDateTZ
        Case "GTT"  ' good till Time
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttGoodTillDate
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttGoodTillDateTZ
        Case "HID"  ' hidden?
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttHidden
        Case "ICE"  ' iceberg
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttDisplaySize
        Case "IOC"  ' immediate or cancel
            mPermittedOrderTifs = mPermittedOrderTifs Or OrderTIFs.OrderTIFImmediateOrCancel
        Case "RTH"
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttIgnoreRTH
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttRTHOnly
        Case "SWEEP"
            mPermittedOrderAttributes = mPermittedOrderAttributes Or OrderAttributes.OrderAttSweepToFill
        Case "ACTIVETIM", _
            "ALLOC", _
            "ALGO", _
            "AVGCOST", _
            "AUC", _
            "BASKET", _
            "COND", _
            "CONDORDER", _
            "CONSCOST", _
            "DARKPOLL", _
            "DEACT", _
            "DEACTDIS", _
            "DEACTEOD", _
            "FOK", _
            "HPENNY", _
            "LARGE", _
            "LTH"
        Case "NONALGO", _
            "OCA", _
            "OPG", _
            "OPGREROUT", _
            "PAON", _
            "PEGMID", _
            "POSTONLY", _
            "PREOPGRTH", _
            "RELSTK", _
            "RTHIGNOPG", _
            "SCALE", _
            "SCALERST", _
            "SMARTSTG"
        Case "SOI", _
            "TIMEPRIO", _
            "TRAILLIT", _
            "TRAILMIT", _
            "VOLAT", _
            "WHATIF"
            ' not sure what these signify
        Case Else
            If lOrderTypesStr(i) <> "" Then
                logMessage "Unknown Order type received from Tws", pMsgQualifier:=lOrderTypesStr(i), pProcName:=ProcName
            End If
        End Select
    End If
Next

mPermittedStopTriggerMethods = StopTriggerMethods.StopTriggerBidAsk Or _
                                StopTriggerMethods.StopTriggerDefault Or _
                                StopTriggerMethods.StopTriggerDoubleBidAsk Or _
                                StopTriggerMethods.StopTriggerDoubleLast Or _
                                StopTriggerMethods.StopTriggerLast Or _
                                StopTriggerMethods.StopTriggerLastOrBidAsk Or _
                                StopTriggerMethods.StopTriggerMidPoint

setState OrderSubmitterStateReady

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyContract(ByVal pRequestId As Long, pContractDetails As TwsContractDetails)
Const ProcName As String = "IContractDetailsConsumer_NotifyContract"
On Error GoTo Err

If Not mContractDetails Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Contract not uniquely specified"
Set mContractDetails = pContractDetails

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IContractDetailsConsumer_NotifyError"
On Error GoTo Err

Err.Raise ErrorCodes.ErrIllegalStateException, , "Failed to retrieve contract details: errorcode=" & pErrorCode & "; description=" & pErrorMsg

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' IOrderSubmitter Interface Members
'================================================================================

Private Sub IOrderSubmitter_AddStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "IOrderSubmitter_AddStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener, Hex$(ObjPtr(pListener))

Exit Sub

Err:
If Err.number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function IOrderSubmitter_CancelOrderPlex(ByVal pOrderPlex As IOrderPlex) As Boolean
Const ProcName As String = "IOrderSubmitter_cancelOrderPlex"
On Error GoTo Err

Dim entryOrder As IOrder
Dim stopOrder As IOrder
Dim targetOrder As IOrder

Set entryOrder = pOrderPlex.entryOrder
Select Case entryOrder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
Case Else
    CancelOrder entryOrder, OrderModeEntry    ' should automatically cancel the other orders
                                                    ' if they have parentid set
    IOrderSubmitter_CancelOrderPlex = True
End Select

Set stopOrder = pOrderPlex.stopOrder
If Not stopOrder Is Nothing Then
    Select Case stopOrder.Status
    Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusFilled, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusCancelled
    Case Else
        Select Case entryOrder.Status
        Case OrderStatuses.OrderStatusFilled, _
            OrderStatuses.OrderStatusCancelling, _
            OrderStatuses.OrderStatusCancelled
            CancelOrder stopOrder, OrderModeStop
            IOrderSubmitter_CancelOrderPlex = True
        End Select
    End Select
End If

Set targetOrder = pOrderPlex.targetOrder
If Not targetOrder Is Nothing Then
    Select Case targetOrder.Status
    Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusFilled, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusCancelled
    Case Else
        Select Case entryOrder.Status
        Case OrderStatuses.OrderStatusFilled, _
            OrderStatuses.OrderStatusCancelling, _
            OrderStatuses.OrderStatusCancelled
            CancelOrder targetOrder, OrderModeTarget
            IOrderSubmitter_CancelOrderPlex = True
        End Select
    End Select
End If
    
' need some stuff here to cancel if it's an oca group

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub IOrderSubmitter_CancelStopOrder(ByVal pOrderPlex As IOrderPlex)
Const ProcName As String = "IOrderSubmitter_cancelStopOrder"
On Error GoTo Err

Dim stopOrder As IOrder

Set stopOrder = pOrderPlex.stopOrder
If stopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "Order plex " & pOrderPlex.Key & " has no stop Order"

Select Case stopOrder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "Stop Order state invalid for cancellation"
Case Else
    CancelOrder stopOrder, OrderModeStop
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderSubmitter_CancelTargetOrder(ByVal pOrderPlex As IOrderPlex)
Const ProcName As String = "IOrderSubmitter_cancelTargetOrder"
Dim targetOrder As IOrder

On Error GoTo Err
Set targetOrder = pOrderPlex.targetOrder
If targetOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "Order plex " & pOrderPlex.Key & " has no target Order"


Select Case targetOrder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "Target Order state invalid for cancellation"
Case Else
    CancelOrder targetOrder, OrderModeTarget
End Select
Exit Sub
Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Let IOrderSubmitter_ClockRate(ByVal RHS As Single)

End Property

Private Sub IOrderSubmitter_ExecuteOrderPlex(ByVal pOrderPlex As IOrderPlex)
Const ProcName As String = "IOrderSubmitter_executeOrderPlex"
Dim entryOrder As IOrder
Dim stopOrder As IOrder
Dim targetOrder As IOrder
Dim Transmit As Boolean
Dim lPlexId As Long

On Error GoTo Err
Set entryOrder = pOrderPlex.entryOrder
Set stopOrder = pOrderPlex.stopOrder
Set targetOrder = pOrderPlex.targetOrder

lPlexId = getPlexId

If Not entryOrder Is Nothing Then
    entryOrder.BrokerId = lPlexId + EntryOrderIdIncrement
    Transmit = True
    If Not stopOrder Is Nothing Then
        stopOrder.BrokerId = lPlexId + StopOrderIdIncrement
        If stopOrder.OrderType <> OrderTypes.OrderTypeAutoStop Then Transmit = False
    End If
    If Not targetOrder Is Nothing Then
        targetOrder.BrokerId = lPlexId + TargetOrderIdIncrement
        If targetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then Transmit = False
    End If
        
    PlaceOrder entryOrder, Transmit, "0", "", OrderModeEntry
    
    If Not stopOrder Is Nothing Then
        Transmit = True
        If Not targetOrder Is Nothing Then
            If targetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then Transmit = False
        End If
        
        If stopOrder.OrderType <> OrderTypes.OrderTypeAutoStop Then
            ' note that AUTOSTP orders will be sent when the entry Order is filled
            PlaceOrder stopOrder, Transmit, entryOrder.BrokerId, "", OrderModeStop
        End If
    End If
    
    If Not targetOrder Is Nothing Then
        If targetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then
            PlaceOrder targetOrder, True, entryOrder.BrokerId, "", OrderModeTarget
        End If
    End If
    
Else
    ' treat the other orders as an OCA group - still to be implemented
End If
Exit Sub
Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderSubmitter_Finish()
Const ProcName As String = "IOrderSubmitter_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderSubmitter_ModifyOrderPlex( _
                ByVal pOrderPlex As IOrderPlex, _
                ByVal entryOrderChanged As Boolean, _
                ByVal stopOrderChanged As Boolean, _
                ByVal targetOrderChanged As Boolean, _
                ByVal closeoutOrderChanged As Boolean)
Const ProcName As String = "IOrderSubmitter_modifyOrderPlex"


Dim OcaGroup As String
Dim ParentId As String

On Error GoTo Err

Select Case pOrderPlex.entryOrder.Status
Case OrderStatusFilled, _
        OrderStatusCancelling, _
        OrderStatusCancelled
    ParentId = "0"
Case Else
    ParentId = pOrderPlex.entryOrder.BrokerId
End Select

If pOrderPlex.entryOrder.QuantityRemaining = 0 Then ParentId = "0"

If Not pOrderPlex.stopOrder Is Nothing Then OcaGroup = pOrderPlex.stopOrder.ProviderProperties.GetParameterValue("OCA group")
If OcaGroup = "" And Not pOrderPlex.targetOrder Is Nothing Then OcaGroup = pOrderPlex.targetOrder.ProviderProperties.GetParameterValue("OCA group")
If OcaGroup = "" Then OcaGroup = GenerateGUIDString

If entryOrderChanged Then
    PlaceOrder pOrderPlex.entryOrder, True, "0", "", OrderModeEntry
End If
If stopOrderChanged Then
    PlaceOrder pOrderPlex.stopOrder, True, ParentId, OcaGroup, OrderModeStop
End If
If targetOrderChanged Then
    PlaceOrder pOrderPlex.targetOrder, True, ParentId, OcaGroup, OrderModeTarget
End If
If closeoutOrderChanged Then
    PlaceOrder pOrderPlex.CloseoutOrder, True, "0", "", OrderModeCloseout
End If
Exit Sub
Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get IOrderSubmitter_PermittedOrderAttributes() As Long
IOrderSubmitter_PermittedOrderAttributes = mPermittedOrderAttributes
End Property

Private Property Get IOrderSubmitter_PermittedOrderTifs() As Long
IOrderSubmitter_PermittedOrderTifs = mPermittedOrderTifs
End Property

Private Property Get IOrderSubmitter_PermittedOrderTypes() As Long
IOrderSubmitter_PermittedOrderTypes = mPermittedOrderTypes
End Property

Private Property Get IOrderSubmitter_PermittedStopTriggerMethods() As Long
IOrderSubmitter_PermittedStopTriggerMethods = mPermittedStopTriggerMethods
End Property

Private Sub IOrderSubmitter_RemoveStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "IOrderSubmitter_RemoveStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener, Hex$(ObjPtr(pListener))

Exit Sub

Err:
If Err.number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderSubmitter_ResubmitStopAndTargetOrders(ByVal pOrderPlex As IOrderPlex)
Const ProcName As String = "IOrderSubmitter_resubmitStopAndTargetOrders"
Dim stopOrder As IOrder
Dim targetOrder As IOrder
Dim OcaGroup As String

On Error GoTo Err
OcaGroup = GenerateGUIDString

Set stopOrder = pOrderPlex.stopOrder
stopOrder.BrokerId = ""   ' force a new id to be allocated
PlaceOrder stopOrder, False, "0", OcaGroup, OrderModeStop

Set targetOrder = pOrderPlex.targetOrder
targetOrder.BrokerId = "" ' force a new id to be allocated
PlaceOrder targetOrder, True, "0", OcaGroup, OrderModeTarget
Exit Sub
Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderSubmitter_ResubmitStopOrder(ByVal pOrderPlex As IOrderPlex)
Const ProcName As String = "IOrderSubmitter_resubmitStopOrder"
Dim stopOrder As IOrder
Dim targetOrder As IOrder
Dim OcaGroup As String

On Error GoTo Err
Set stopOrder = pOrderPlex.stopOrder
Set targetOrder = pOrderPlex.targetOrder
stopOrder.BrokerId = ""     ' force a new id to be allocated
If targetOrder Is Nothing Then
    PlaceOrder stopOrder, True, "0", "", OrderModeStop
Else
    OcaGroup = GenerateGUIDString
    PlaceOrder stopOrder, False, "0", OcaGroup, OrderModeStop
    PlaceOrder targetOrder, True, "0", OcaGroup, OrderModeTarget
End If
Exit Sub
Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderSubmitter_ResubmitTargetOrder(ByVal pOrderPlex As IOrderPlex)
Const ProcName As String = "IOrderSubmitter_resubmitTargetOrder"
Dim stopOrder As IOrder
Dim targetOrder As IOrder
Dim OcaGroup As String

On Error GoTo Err
Set stopOrder = pOrderPlex.stopOrder
Set targetOrder = pOrderPlex.targetOrder
targetOrder.BrokerId = ""   ' force a new id to be allocated
If stopOrder Is Nothing Then
    PlaceOrder targetOrder, True, "0", "", OrderModeTarget
Else
    OcaGroup = GenerateGUIDString
    PlaceOrder stopOrder, False, "0", OcaGroup, OrderModeStop
    PlaceOrder targetOrder, True, "0", OcaGroup, OrderModeTarget
End If
Exit Sub
Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get IOrderSubmitter_State() As OrderSubmitterStates
IOrderSubmitter_State = mState
End Property

'================================================================================
' mTwsAPI Event Handlers
'================================================================================

Private Sub mTwsAPI_Connected()
Const ProcName As String = "mTwsAPI_Connected"
On Error GoTo Err

If Not mContractDetails Is Nothing Then setState OrderSubmitterStateReady

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mTwsAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
Const ProcName As String = "mTwsAPI_ConnectFailed"
On Error GoTo Err

setState OrderSubmitterStateNotReady
Exit Sub
Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mTwsAPI_ConnectionToIBClosed()
Const ProcName As String = "mTwsAPI_ConnectionToIBClosed"
On Error GoTo Err

setState OrderSubmitterStateNotReady
Exit Sub
Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mTwsAPI_ConnectionToIBRecovered()
Const ProcName As String = "mTwsAPI_ConnectionToIBRecovered"
On Error GoTo Err
If Not mContractDetails Is Nothing Then setState OrderSubmitterStateReady
Exit Sub
Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mTwsAPI_ConnectionToTwsClosed( _
                ByVal reconnecting As Boolean)
Const ProcName As String = "mTwsAPI_ConnectionToTwsClosed"
On Error GoTo Err

setState OrderSubmitterStateNotReady
Exit Sub
Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mTwsAPI_Disconnected(ByVal reason As String, ByVal pForced As Boolean)
Const ProcName As String = "mTwsAPI_Disconnected"
On Error GoTo Err

setState OrderSubmitterStateNotReady
Exit Sub
Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"

On Error GoTo Err

mStateChangeListeners.Clear

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTwsAPI As TwsAPI, _
                ByVal pOrderPlacer As OrderPlacer, _
                ByVal pContractSpec As ContractSpecifier, _
                ByVal pOrderSubmissionListener As IOrderSubmissionListener, _
                ByVal pContractRequester As ContractDetailsRequester)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mOrderPlacer = pOrderPlacer
Set mOrderSubmissionListener = pOrderSubmissionListener
Set mTwsAPI = pTwsAPI
Set mContractRequester = pContractRequester

mContractRequester.RequestContractDetails pContractSpec, Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub NotifyError( _
                ByVal pOrderId As String, _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "NotifyError"
On Error GoTo Err

mOrderSubmissionListener.NotifyError pOrderId, pErrorCode, pErrorMsg

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub NotifyExecutionReport(ByVal pExecutionReport As IExecutionReport)
Const ProcName As String = "NotifyExecutionReport"
On Error GoTo Err

mOrderSubmissionListener.NotifyExecutionReport pExecutionReport

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub NotifyOrderReport(ByVal pOrderReport As IOrderReport)
Const ProcName As String = "NotifyOrderReport"
On Error GoTo Err

mOrderSubmissionListener.NotifyOrderReport pOrderReport

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub NotifyOrderStatusReport(ByVal pOrderStatusReport As IOrderStatusReport)
Const ProcName As String = "NotifyOrderStatusReport"
On Error GoTo Err

mOrderSubmissionListener.NotifyOrderStatusReport pOrderStatusReport

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub CancelOrder( _
                ByVal pOrder As IOrder, _
                ByRef orderMode As String)
Const ProcName As String = "CancelOrder"

On Error GoTo Err

logMessage "Cancel " & orderMode & " Order: " & _
            "broker id=" & pOrder.BrokerId & _
            "; TradeBuild id=" & pOrder.id, _
            ProcName

mOrderPlacer.CancelOrder pOrder


Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub fireStateChange()
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim lListener As StateChangeListener
Dim ev As StateChangeEventData

Set ev.source = Me
ev.state = mState
For Each lListener In mStateChangeListeners
    lListener.Change ev
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function getPlexId() As Long
Const ProcName As String = "getPlexId"
On Error GoTo Err

getPlexId = Int((mTwsAPI.NextOrderId + 4) / 5) * 5

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Sub PlaceOrder( _
                ByVal pOrder As IOrder, _
                ByVal pTransmit As Boolean, _
                ByVal pParentId As String, _
                ByVal pOcaGroup As String, _
                ByRef pOrderMode As String)
Const ProcName As String = "PlaceOrder"

On Error GoTo Err

logMessage "Place " & pOrderMode & " Order: " & _
                "broker id=" & pOrder.BrokerId & _
                "; TradeBuild id=" & pOrder.id, _
            ProcName

mOrderPlacer.PlaceOrder Me, pOrder, mContractDetails.Summary, pTransmit, pParentId, pOcaGroup

If pOrder.Status = OrderStatusCreated Then
    Dim statusRpt As New OrderStatusReport
    statusRpt.Initialise pOrder.id, OrderStatusPendingSubmit
    mOrderSubmissionListener.NotifyOrderStatusReport statusRpt
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setState(ByVal pState As OrderSubmitterStates)
Const ProcName As String = "setState"
On Error GoTo Err

If mState = pState Then Exit Sub
mState = pState
fireStateChange

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub




