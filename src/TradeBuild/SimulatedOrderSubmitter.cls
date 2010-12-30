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
Attribute VB_Exposed = False
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

Implements TradeBuildSP.IOrderSubmitter
Implements TradeBuildSP.IStreamingDataConsumer

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "SimulatedOrderSubmitter"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mServiceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer

Private mSimulator As AdvancedOrderSimulator

Private mMarketIndex As Long

Private mContract As Contract

Private mFirstAskReceived       As Boolean
Private mFirstBidReceived       As Boolean
Private mReadyNotified          As Boolean


'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "SimulatedOrderSubmitter terminated"
End Sub

'@================================================================================
' IOrderSubmitter Interface Members
'@================================================================================

Private Function IOrderSubmitter_cancelOrderPlex(ByVal pOrderPlex As TradeBuildSP.IOrderPlex) As Boolean
Dim EntryOrder As Order
Dim Stoporder As Order
Dim TargetOrder As Order

Const ProcName As String = "IOrderSubmitter_cancelOrderPlex"
On Error GoTo Err

Set EntryOrder = pOrderPlex.EntryOrder
Select Case EntryOrder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
Case Else
    mSimulator.CancelOrder EntryOrder.BrokerId  ' should automatically cancel the other orders
                                                    ' if they have parentid set
    IOrderSubmitter_cancelOrderPlex = True
End Select

Set Stoporder = pOrderPlex.Stoporder
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
            mSimulator.CancelOrder Stoporder.BrokerId
            IOrderSubmitter_cancelOrderPlex = True
        End Select
    End Select
End If

Set TargetOrder = pOrderPlex.TargetOrder
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
            mSimulator.CancelOrder TargetOrder.BrokerId
            IOrderSubmitter_cancelOrderPlex = True
        End Select
    End Select
End If
    
' need some stuff here to cancel if it's an oca group

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Private Sub IOrderSubmitter_cancelStopOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim Stoporder As Order

Const ProcName As String = "IOrderSubmitter_cancelStopOrder"
On Error GoTo Err

Set Stoporder = pOrderPlex.Stoporder
If Stoporder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "Order plex has no stop order"

Select Case Stoporder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "Stop order state invalid for cancellation"
Case Else
    mSimulator.CancelOrder Stoporder.BrokerId
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderSubmitter_cancelTargetOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim TargetOrder As Order

Const ProcName As String = "IOrderSubmitter_cancelTargetOrder"
On Error GoTo Err

Set TargetOrder = pOrderPlex.TargetOrder
If TargetOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "Order plex has no target order"


Select Case TargetOrder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "Target order state invalid for cancellation"
Case Else
    mSimulator.CancelOrder TargetOrder.BrokerId
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Let IOrderSubmitter_ClockRate(ByVal RHS As Single)
Const ProcName As String = "IOrderSubmitter_ClockRate"
On Error GoTo Err

mSimulator.ClockRate(mMarketIndex) = RHS

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub IOrderSubmitter_executeOrderPlex(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim EntryOrder As Order
Dim Stoporder As Order
Dim TargetOrder As Order
Dim transmit As Boolean
Dim OcaGroup As String

Const ProcName As String = "IOrderSubmitter_executeOrderPlex"
On Error GoTo Err

Set EntryOrder = pOrderPlex.EntryOrder
Set Stoporder = pOrderPlex.Stoporder
Set TargetOrder = pOrderPlex.TargetOrder

If Not EntryOrder Is Nothing Then
    transmit = True
    If Not Stoporder Is Nothing Then
        If Stoporder.OrderType <> OrderTypes.OrderTypeAutoStop Then transmit = False
    End If
    If Not TargetOrder Is Nothing Then
        If TargetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then transmit = False
    End If
        
    placeOrder EntryOrder, "", "", transmit
    
    OcaGroup = GenerateTextID

    If Not Stoporder Is Nothing Then
        transmit = True
        If Not TargetOrder Is Nothing Then
            If TargetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then transmit = False
        End If
        
        If Stoporder.OrderType <> OrderTypes.OrderTypeAutoStop Then
            ' note that AUTOSTP orders will be sent when the entry order is filled
            placeOrder Stoporder, EntryOrder.Id, OcaGroup, transmit
        End If
    End If
    
    If Not TargetOrder Is Nothing Then
        If TargetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then
            placeOrder TargetOrder, EntryOrder.Id, OcaGroup, True
        End If
    End If
    
Else
    ' treat the other orders as an OCA group - still to be implemented
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderSubmitter_Finish()
Const ProcName As String = "IOrderSubmitter_Finish"
On Error GoTo Err

mSimulator.stopMarket mMarketIndex
Set mServiceConsumer = Nothing
Set mSimulator = Nothing
Set mContract = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderSubmitter_modifyOrderPlex( _
                ByVal pOrderPlex As TradeBuildSP.IOrderPlex, _
                ByVal entryOrderChanged As Boolean, _
                ByVal stopOrderChanged As Boolean, _
                ByVal targetOrderChanged As Boolean, _
                ByVal closeoutOrderChanged As Boolean)
Dim lOrder As Order

Const ProcName As String = "IOrderSubmitter_modifyOrderPlex"
On Error GoTo Err

If entryOrderChanged Then
    placeOrder pOrderPlex.EntryOrder, "", "", True
End If
If stopOrderChanged Then
    Set lOrder = pOrderPlex.Stoporder
    placeOrder lOrder, "", "", True
End If
If targetOrderChanged Then
    Set lOrder = pOrderPlex.Stoporder
    placeOrder lOrder, "", "", True
End If
If closeoutOrderChanged Then
    placeOrder pOrderPlex.CloseoutOrder, "", "", True
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
    
End Sub

Private Property Get IOrderSubmitter_permittedOrderAttributes() As Long
Dim attributes As Long
Const ProcName As String = "IOrderSubmitter_permittedOrderAttributes"
On Error GoTo Err

attributes = OrderAttributeIds.OrderAttAction
attributes = attributes Or OrderAttributeIds.OrderAttLimitPrice
attributes = attributes Or OrderAttributeIds.OrderAttOrderType
attributes = attributes Or OrderAttributeIds.OrderAttOriginatorRef
attributes = attributes Or OrderAttributeIds.OrderAttQuantity
attributes = attributes Or OrderAttributeIds.OrderAttTimeInForce
attributes = attributes Or OrderAttributeIds.OrderAttTriggerPrice
IOrderSubmitter_permittedOrderAttributes = attributes

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get IOrderSubmitter_permittedOrderTifs() As Long
Dim permittedTifs As Long
Const ProcName As String = "IOrderSubmitter_permittedOrderTifs"
On Error GoTo Err

permittedTifs = TIFDay
IOrderSubmitter_permittedOrderTifs = permittedTifs

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get IOrderSubmitter_permittedOrderTypes() As Long
Dim permittedOrderTypes As Long
Const ProcName As String = "IOrderSubmitter_permittedOrderTypes"
On Error GoTo Err

permittedOrderTypes = OrderTypeLimit Or _
                        OrderTypeMarket Or _
                        OrderTypeStop Or _
                        OrderTypeStopLimit
IOrderSubmitter_permittedOrderTypes = permittedOrderTypes

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get IOrderSubmitter_permittedStopTriggerMethods() As Long
Dim permittedStopTriggerMethods As Long
Const ProcName As String = "IOrderSubmitter_permittedStopTriggerMethods"
On Error GoTo Err

permittedStopTriggerMethods = StopTriggerLast
IOrderSubmitter_permittedStopTriggerMethods = permittedStopTriggerMethods

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub IOrderSubmitter_resubmitStopAndTargetOrders(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim OcaGroup As String
Dim Stoporder As Order
Dim TargetOrder As Order

Const ProcName As String = "IOrderSubmitter_resubmitStopAndTargetOrders"
On Error GoTo Err

OcaGroup = GenerateTextID

Set Stoporder = pOrderPlex.Stoporder
Stoporder.BrokerId = ""    ' force a new id to be allocated
placeOrder Stoporder, "", OcaGroup, False

Set TargetOrder = pOrderPlex.TargetOrder
TargetOrder.BrokerId = "" ' force a new id to be allocated
placeOrder TargetOrder, "", OcaGroup, True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderSubmitter_resubmitStopOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim Stoporder As Order
Dim TargetOrder As Order

Const ProcName As String = "IOrderSubmitter_resubmitStopOrder"
On Error GoTo Err

Set Stoporder = pOrderPlex.Stoporder
Set TargetOrder = pOrderPlex.TargetOrder
Stoporder.BrokerId = ""     ' force a new id to be allocated
If Not pOrderPlex.TargetOrder Is Nothing Then
    placeOrder Stoporder, "", mSimulator.GetOrder(TargetOrder.Id), True
Else
    placeOrder Stoporder, "", "", True
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderSubmitter_resubmitTargetOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim Stoporder As Order
Dim TargetOrder As Order

Const ProcName As String = "IOrderSubmitter_resubmitTargetOrder"
On Error GoTo Err

Set Stoporder = pOrderPlex.Stoporder
Set TargetOrder = pOrderPlex.TargetOrder
TargetOrder.BrokerId = ""   ' force a new id to be allocated
If Not pOrderPlex.Stoporder Is Nothing Then
    placeOrder TargetOrder, "", mSimulator.GetOrder(Stoporder.Id), True
Else
    placeOrder TargetOrder, "", "", True
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' IStreamingDataConsumer Interface Members
'@================================================================================

Private Sub IStreamingDataConsumer_Tick(Tick As TickUtils26.GenericTick)
Const ProcName As String = "IStreamingDataConsumer_Tick"
On Error GoTo Err

Select Case Tick.tickType
Case TickTypes.TickTypeAsk
    mFirstAskReceived = True
    If mFirstBidReceived And Not mReadyNotified Then
        mReadyNotified = True
        mServiceConsumer.ready
    End If
        
    mSimulator.Ask Tick.Timestamp, mMarketIndex, Tick.price
Case TickTypes.TickTypeBid
    mFirstBidReceived = True
    If mFirstAskReceived And Not mReadyNotified Then
        mReadyNotified = True
        mServiceConsumer.ready
    End If
        
    mSimulator.Bid Tick.Timestamp, mMarketIndex, Tick.price
Case TickTypes.TickTypeClosePrice
Case TickTypes.TickTypeHighPrice
Case TickTypes.TickTypeLowPrice
Case TickTypes.TickTypeMarketDepth
Case TickTypes.TickTypeMarketDepthReset
Case TickTypes.TickTypeOpenInterest
Case TickTypes.TickTypeOpenPrice
Case TickTypes.TickTypeTrade
    mSimulator.Trade Tick.Timestamp, mMarketIndex, Tick.price
Case TickTypes.TickTypeVolume
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Contract(ByVal value As Contract)
Const ProcName As String = "Contract"
On Error GoTo Err

Set mContract = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let serviceConsumer(ByVal value As TradeBuildSP.IOrderSubmissionSrvcConsumer)
Const ProcName As String = "serviceConsumer"
On Error GoTo Err

Set mServiceConsumer = value
mServiceConsumer.RequestTickData Me, False

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let simulator(ByVal value As AdvancedOrderSimulator)
Const ProcName As String = "simulator"
On Error GoTo Err

Set mSimulator = value
mMarketIndex = mSimulator.startMarket(mServiceConsumer, _
                                    mContract)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub placeOrder( _
                ByVal pOrder As Order, _
                ByVal ParentId As String, _
                ByVal OcaGroup As String, _
                ByVal transmit As Boolean)
Const ProcName As String = "placeOrder"
On Error GoTo Err

mSimulator.placeOrder pOrder, ParentId, OcaGroup, mMarketIndex, transmit
If pOrder.Status = OrderStatusCreated Then mServiceConsumer.OrderStatusReport gCreateOrderStatusReport(pOrder.Id, OrderStatusPendingSubmit)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

