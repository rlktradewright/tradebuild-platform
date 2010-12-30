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

Private WithEvents mTWSAPI As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1

Private mHandle As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Finish
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
Const ProcName As String = "ContractDetailsRequester_contractsLoaded"
Dim failpoint As String
On Error GoTo Err

If mServiceConsumer Is Nothing Then Exit Sub

If request.contractWrappers.Count = 1 Then
    Set mContractWrapper = request.contractWrappers(1)
    mServiceConsumer.Ready
Else
    mServiceConsumer.notifyEvent StandardSPEventCodes.LONotUniqueContract, _
                            "Contract is not uniquely specified", _
                            mHandle
                            
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub ContractDetailsRequester_contractSpecifierInvalid( _
                ByRef request As TWSContractDetailsRequestQueueEntry, _
                ByVal reason As String)
Const ProcName As String = "ContractDetailsRequester_contractSpecifierInvalid"
Dim failpoint As String
On Error GoTo Err

If mServiceConsumer Is Nothing Then Exit Sub

mServiceConsumer.notifyEvent StandardSPEventCodes.LOUnknownContract, _
                        "Contract not known", _
                        mHandle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' IOrderSubmitter Interface Members
'================================================================================

Private Function IOrderSubmitter_cancelOrderPlex(ByVal pOrderPlex As TradeBuildSP.IOrderPlex) As Boolean
Const ProcName As String = "IOrderSubmitter_cancelOrderPlex"
Dim entryOrder As TradeBuildSP.IOrder
Dim stopOrder As TradeBuildSP.IOrder
Dim targetOrder As TradeBuildSP.IOrder

On Error GoTo Err
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
Err:
handleFatalError ProcName

End Function

Private Sub IOrderSubmitter_cancelStopOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Const ProcName As String = "IOrderSubmitter_cancelStopOrder"
Dim stopOrder As TradeBuildSP.IOrder

On Error GoTo Err
Set stopOrder = pOrderPlex.stopOrder
If stopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "Order plex " & pOrderPlex.key & " has no stop order"

Select Case stopOrder.status
Case TradeBuildSP.OrderStatuses.OrderStatusCreated, _
    TradeBuildSP.OrderStatuses.OrderStatusFilled, _
    TradeBuildSP.OrderStatuses.OrderStatusCancelling, _
    TradeBuildSP.OrderStatuses.OrderStatusCancelled
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "Stop order state invalid for cancellation"
Case Else
    cancelOrder stopOrder, OrderModeStop
End Select
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub IOrderSubmitter_cancelTargetOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Const ProcName As String = "IOrderSubmitter_cancelTargetOrder"
Dim targetOrder As TradeBuildSP.IOrder

On Error GoTo Err
Set targetOrder = pOrderPlex.targetOrder
If targetOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "Order plex " & pOrderPlex.key & " has no target order"


Select Case targetOrder.status
Case TradeBuildSP.OrderStatuses.OrderStatusCreated, _
    TradeBuildSP.OrderStatuses.OrderStatusCancelling, _
    TradeBuildSP.OrderStatuses.OrderStatusCancelled
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "Target order state invalid for cancellation"
Case Else
    cancelOrder targetOrder, OrderModeTarget
End Select
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Property Let IOrderSubmitter_clockRate(ByVal RHS As Single)

End Property

Private Sub IOrderSubmitter_executeOrderPlex(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Const ProcName As String = "IOrderSubmitter_executeOrderPlex"
Dim entryOrder As TradeBuildSP.IOrder
Dim stopOrder As TradeBuildSP.IOrder
Dim targetOrder As TradeBuildSP.IOrder
Dim transmit As Boolean
Dim orderPlexId As Long

On Error GoTo Err
Set entryOrder = pOrderPlex.entryOrder
Set stopOrder = pOrderPlex.stopOrder
Set targetOrder = pOrderPlex.targetOrder

orderPlexId = mTWSAPI.GetNextOrderPlexId

If Not entryOrder Is Nothing Then
    entryOrder.BrokerId = orderPlexId + 1
    transmit = True
    If Not stopOrder Is Nothing Then
        stopOrder.BrokerId = orderPlexId + 2
        If stopOrder.OrderType <> TradeBuildSP.OrderTypes.OrderTypeAutoStop Then transmit = False
    End If
    If Not targetOrder Is Nothing Then
        targetOrder.BrokerId = orderPlexId + 3
        If targetOrder.OrderType <> TradeBuildSP.OrderTypes.OrderTypeAutoLimit Then transmit = False
    End If
        
    placeOrder entryOrder, transmit, "0", "", OrderModeEntry
    
    If Not stopOrder Is Nothing Then
        transmit = True
        If Not targetOrder Is Nothing Then
            If targetOrder.OrderType <> TradeBuildSP.OrderTypes.OrderTypeAutoLimit Then transmit = False
        End If
        
        If stopOrder.OrderType <> TradeBuildSP.OrderTypes.OrderTypeAutoStop Then
            ' note that AUTOSTP orders will be sent when the entry order is filled
            placeOrder stopOrder, transmit, entryOrder.BrokerId, "", OrderModeStop
        End If
    End If
    
    If Not targetOrder Is Nothing Then
        If targetOrder.OrderType <> TradeBuildSP.OrderTypes.OrderTypeAutoLimit Then
            placeOrder targetOrder, True, entryOrder.BrokerId, "", OrderModeTarget
        End If
    End If
    
Else
    ' treat the other orders as an OCA group - still to be implemented
End If
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub IOrderSubmitter_Finish()
Const ProcName As String = "IOrderSubmitter_Finish"
On Error GoTo Err

finishIt
mOrderSubmissionServiceProvider.orderSubmitterFinished Me
Set mOrderSubmissionServiceProvider = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderSubmitter_modifyOrderPlex( _
                ByVal pOrderPlex As TradeBuildSP.IOrderPlex, _
                ByVal entryOrderChanged As Boolean, _
                ByVal stopOrderChanged As Boolean, _
                ByVal targetOrderChanged As Boolean, _
                ByVal closeoutOrderChanged As Boolean)
Const ProcName As String = "IOrderSubmitter_modifyOrderPlex"


Dim ocaGroup As String
Dim parentId As String

On Error GoTo Err

Select Case pOrderPlex.entryOrder.status
Case OrderStatusFilled, _
        OrderStatusCancelling, _
        OrderStatusCancelled
    parentId = "0"
Case Else
    parentId = pOrderPlex.entryOrder.BrokerId
End Select

If pOrderPlex.entryOrder.QuantityRemaining = 0 Then parentId = "0"

If Not pOrderPlex.stopOrder Is Nothing Then ocaGroup = pOrderPlex.stopOrder.ProviderProperties.GetParameterValue("OCA group")
If ocaGroup = "" And Not pOrderPlex.targetOrder Is Nothing Then ocaGroup = pOrderPlex.targetOrder.ProviderProperties.GetParameterValue("OCA group")
If ocaGroup = "" Then ocaGroup = GenerateGUIDString

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
    placeOrder pOrderPlex.CloseoutOrder, True, "0", "", OrderModeCloseout
End If
Exit Sub
Err:
handleFatalError ProcName
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
Const ProcName As String = "IOrderSubmitter_resubmitStopAndTargetOrders"
Dim stopOrder As TradeBuildSP.IOrder
Dim targetOrder As TradeBuildSP.IOrder
Dim ocaGroup As String

On Error GoTo Err
ocaGroup = GenerateGUIDString

Set stopOrder = pOrderPlex.stopOrder
stopOrder.BrokerId = ""   ' force a new id to be allocated
placeOrder stopOrder, False, "0", ocaGroup, OrderModeStop

Set targetOrder = pOrderPlex.targetOrder
targetOrder.BrokerId = "" ' force a new id to be allocated
placeOrder targetOrder, True, "0", ocaGroup, OrderModeTarget
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub IOrderSubmitter_resubmitStopOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Const ProcName As String = "IOrderSubmitter_resubmitStopOrder"
Dim stopOrder As TradeBuildSP.IOrder
Dim targetOrder As TradeBuildSP.IOrder
Dim ocaGroup As String

On Error GoTo Err
Set stopOrder = pOrderPlex.stopOrder
Set targetOrder = pOrderPlex.targetOrder
stopOrder.BrokerId = ""     ' force a new id to be allocated
If targetOrder Is Nothing Then
    placeOrder stopOrder, True, "0", "", OrderModeStop
Else
    ocaGroup = GenerateGUIDString
    placeOrder stopOrder, False, "0", ocaGroup, OrderModeStop
    placeOrder targetOrder, True, "0", ocaGroup, OrderModeTarget
End If
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub IOrderSubmitter_resubmitTargetOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Const ProcName As String = "IOrderSubmitter_resubmitTargetOrder"
Dim stopOrder As TradeBuildSP.IOrder
Dim targetOrder As TradeBuildSP.IOrder
Dim ocaGroup As String

On Error GoTo Err
Set stopOrder = pOrderPlex.stopOrder
Set targetOrder = pOrderPlex.targetOrder
targetOrder.BrokerId = ""   ' force a new id to be allocated
If stopOrder Is Nothing Then
    placeOrder targetOrder, True, "0", "", OrderModeTarget
Else
    ocaGroup = GenerateGUIDString
    placeOrder stopOrder, False, "0", ocaGroup, OrderModeStop
    placeOrder targetOrder, True, "0", ocaGroup, OrderModeTarget
End If
Exit Sub
Err:
handleFatalError ProcName
End Sub

'================================================================================
' mTWSAPI Event Handlers
'================================================================================

Private Sub mTWSAPI_Connected()
Const ProcName As String = "mTWSAPI_Connected"
On Error GoTo Err

If Not mContractWrapper Is Nothing Then mServiceConsumer.Ready

Exit Sub

Err:
handleFatalError ProcName
End Sub

Private Sub mTWSAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
Const ProcName As String = "mTWSAPI_ConnectFailed"
On Error GoTo Err

mServiceConsumer.NotReady
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub mTWSAPI_ConnectionToIBClosed()
Const ProcName As String = "mTWSAPI_ConnectionToIBClosed"
On Error GoTo Err

mServiceConsumer.NotReady
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub mTWSAPI_ConnectionToIBRecovered()
Const ProcName As String = "mTWSAPI_ConnectionToIBRecovered"
On Error GoTo Err
If Not mContractWrapper Is Nothing Then mServiceConsumer.Ready
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed( _
                ByVal reconnecting As Boolean)
Const ProcName As String = "mTWSAPI_ConnectionToTWSClosed"
On Error GoTo Err

mServiceConsumer.NotReady
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub mTWSAPI_Disconnected(ByVal reason As String)
Const ProcName As String = "mTWSAPI_Disconnected"
On Error GoTo Err

mServiceConsumer.NotReady
Exit Sub
Err:
handleFatalError ProcName
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub ExecutionReport( _
                ByVal pExecReport As ExecutionReport)
Const ProcName As String = "ExecutionReport"
Dim failpoint As String
On Error GoTo Err

mServiceConsumer.ExecutionReport pExecReport

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"
Dim failpoint As String
On Error GoTo Err

finishIt
Set mOrderSubmissionServiceProvider = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pOrderSubmissionServiceProvider As IOrderSubmissionSrvcProvider, _
                ByVal CommonServiceConsumer As ICommonServiceConsumer, _
                ByVal orderSubmissionServiceConsumer As IOrderSubmissionSrvcConsumer, _
                ByVal pContract As Contract, _
                ByVal pTWSAPI As TWSAPI, _
                ByVal Handle As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mOrderSubmissionServiceProvider = pOrderSubmissionServiceProvider
Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = orderSubmissionServiceConsumer
Set mTradeBuildContract = pContract
Set mTWSAPI = pTWSAPI
mHandle = Handle

mTWSAPI.RequestContract mTradeBuildContract.specifier, Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub InvalidOrder( _
                ByVal orderID As String, _
                ByVal message As String)
Const ProcName As String = "InvalidOrder"
                
Dim failpoint As String
On Error GoTo Err

mServiceConsumer.RejectOrder orderID, _
                        message

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub OrderDeferred( _
                ByVal orderID As String, _
                ByVal message As String)
Const ProcName As String = "OrderDeferred"

Dim failpoint As String
On Error GoTo Err

mServiceConsumer.notifyEvent StandardSPEventCodes.LOOrderDeferred, message, mHandle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub OrderReport( _
                ByVal lOrderReport As OrderReport)
Const ProcName As String = "OrderReport"
Dim failpoint As String
On Error GoTo Err

mServiceConsumer.OrderReport lOrderReport

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub OrderStatusReport( _
                ByVal pOrderStatusReport As OrderStatusReport)
Const ProcName As String = "OrderStatusReport"


Dim failpoint As String
On Error GoTo Err

mServiceConsumer.OrderStatusReport pOrderStatusReport

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub
'================================================================================
' Helper Functions
'================================================================================

Private Sub cancelOrder( _
                ByVal pOrder As TradeBuildSP.IOrder, _
                ByRef orderMode As String)
Const ProcName As String = "cancelOrder"
Dim failpoint As String
On Error GoTo Err

logMessage "Cancel " & orderMode & " order: " & _
            "broker id=" & pOrder.BrokerId & _
            "; TradeBuild id=" & pOrder.tradebuildId, _
            "cancelOrder"
If pOrder.status = OrderStatusPendingSubmit Then
    ' if an order is cancelled immediately after being placed, and before
    ' an orderStatus or openOrder notification has been received from TWS,
    ' TWS sometimes just cancels it without ever sending either of these,
    ' so we need to generate one as a backstop
    mTWSAPI.GenerateDeferredOrderCancelNotification pOrder.BrokerId, mTWSAPI.clientID
End If

mTWSAPI.cancelOrder pOrder.BrokerId


Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub finishIt()
Const ProcName As String = "finishIt"
Dim failpoint As String
On Error GoTo Err

Set mCommonServiceConsumer = Nothing
Set mServiceConsumer = Nothing
Set mTradeBuildContract = Nothing
Set mTWSAPI = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As String)
Dim errNum As Long: errNum = Err.number
Dim errSource As String: errSource = Err.source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
errSource = errSource & vbCrLf & _
            ProjectName & "." & _
            ModuleName & ":" & _
            pProcName & _
            IIf(pFailpoint <> "", " At " & pFailpoint, "")

Err.Raise errNum, errSource, errDesc

' NB: will never get to here so no need for Exit Sub

Err:
mCommonServiceConsumer.NotifyFatalError Err.number, Err.source, Err.Description, mHandle
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Sub placeOrder( _
                ByVal pOrder As TradeBuildSP.IOrder, _
                ByVal transmit As Boolean, _
                ByVal parentId As String, _
                ByVal ocaGroup As String, _
                ByRef orderMode As String)
Const ProcName As String = "placeOrder"
Dim failpoint As String
On Error GoTo Err

gLogger.Log "Place " & orderMode & " order: " & _
                "broker id=" & pOrder.BrokerId & _
                "; TradeBuild id=" & pOrder.tradebuildId, _
            ProcName, ModuleName
mTWSAPI.placeOrder Me, pOrder, transmit, parentId, ocaGroup, ""
If pOrder.status = OrderStatusCreated Then
    Dim statusRpt As New OrderStatusReport
    statusRpt.Initialise pOrder.tradebuildId, OrderStatusPendingSubmit
    mServiceConsumer.OrderStatusReport statusRpt
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub



