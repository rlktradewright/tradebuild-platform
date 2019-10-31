VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RolloverController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IChangeListener
Implements IJSONable
Implements ITask

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Public Enum RolloverControllerStates
    RolloverControllerStateInitial = 0
    RolloverControllerStateClosingCurrentPosition
    RolloverControllerStateFindingNextContract
    RolloverControllerStateEstablishingNewPosition
    RolloverControllerStateCompleted
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "RolloverController"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private WithEvents mBracketOrder                    As BracketOrder
Attribute mBracketOrder.VB_VarHelpID = -1

Private mState                                      As RolloverControllerStates

Private mStateChangeListeners                       As New Listeners

Private mOrderManager                               As OrderManager

Private mScopeName                                  As String
Private mNextPositionManager                        As PositionManager
Private WithEvents mNextOrderContext                As OrderContext
Attribute mNextOrderContext.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mState = RolloverControllerStateInitial
End Sub

'@================================================================================
' IChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
On Error GoTo Err

Dim lChangeType As BracketOrderChangeTypes
lChangeType = ev.changeType
Select Case lChangeType
Case BracketOrderChangeTypes.BracketOrderCloseoutOrderFilled
    If mBracketOrder.Size = 0 Then mTaskContext.Continue
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IJSONable Interface Members
'@================================================================================

Private Function IJSONAble_ToJSON() As String
IJSONAble_ToJSON = ToJSON
End Function

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

Select Case mState

Case RolloverControllerStateInitial
    mBracketOrder.CloseOut
    mState = RolloverControllerStateClosingCurrentPosition
    fireStateChange RolloverControllerStateClosingCurrentPosition
    mTaskContext.Suspend -1
Case RolloverControllerStateClosingCurrentPosition
    If mBracketOrder.Size <> 0 Then mTaskContext.Suspend -1
    mState = RolloverControllerStateFindingNextContract
    fireStateChange RolloverControllerStateFindingNextContract
Case RolloverControllerStateFindingNextContract
    findNextContract
    mState = RolloverControllerStateEstablishingNewPosition
    fireStateChange mState
Case RolloverControllerStateEstablishingNewPosition
    Set mNextOrderContext = mNextPositionManager.OrderContexts.DefaultOrderContext
    mTaskContext.Suspend -1

End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' mNextOrderContext Event Handlers
'@================================================================================

Private Sub mNextOrderContext_Change(ev As ChangeEventData)
Const ProcName As String = "mNextOrderContext_Change"
On Error GoTo Err

If ev.changeType <> OrderContextChangeTypes.OrderContextReadyStateChanged Then Exit Sub

Dim lEntryOrderType As OrderTypes: lEntryOrderType = mBracketOrder.EntryOrder.OrderType

Dim lStopLossOrderType As OrderTypes: lStopLossOrderType = OrderTypeNone
If Not mBracketOrder.StopLossOrder Is Nothing Then lStopLossOrderType = mBracketOrder.StopLossOrder.OrderType

Dim lTargetOrderType As OrderTypes: lTargetOrderType = OrderTypeNone
If Not mBracketOrder.TargetOrder Is Nothing Then lTargetOrderType = mBracketOrder.TargetOrder.OrderType

If Not mNextOrderContext.IsReadyForBracketOrder(mBracketOrder) Then
    Exit Sub
End If

placeNextBracketOrder

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let ScopeName(ByVal Value As String)
mScopeName = Value
End Property

Friend Property Get ScopeName() As String
ScopeName = mScopeName
End Property

Public Property Let NextPositionManager(ByVal Value As PositionManager)
Set mNextPositionManager = Value
End Property

Friend Property Get NextPositionManager() As PositionManager
Set NextPositionManager = mNextPositionManager
End Property

Public Property Let State(ByVal Value As RolloverControllerStates)
mState = Value
End Property

Friend Property Get State() As RolloverControllerStates
State = mState
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "AddStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pOrderManager As OrderManager, _
                ByVal pScopeName As String)
Set mBracketOrder = pBracketOrder
mBracketOrder.AddChangeListener Me
Set mOrderManager = pOrderManager
mScopeName = pScopeName
End Sub

Friend Function ToJSON() As String
Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeNameAndValue "State", mState, RolloverControllerStates.RolloverControllerStateInitial

jb.EncodeNameAndValue "ScopeName", mScopeName
jb.EncodeNameAndValue "NextPositionManager", mNextPositionManager

jb.EndObject

ToJSON = jb.JSONString
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub findNextContract()
Const ProcName As String = "findNextContract"
On Error GoTo Err

Dim lCurrentContractSpec As IContractSpecifier
Set lCurrentContractSpec = mBracketOrder.Contract.Specifier

Select Case lCurrentContractSpec.SecType
Case SecTypeFuture
    Dim lNextContractSpec As IContractSpecifier
    Set lNextContractSpec = CreateContractSpecifier("", _
                                                lCurrentContractSpec.Symbol, _
                                                lCurrentContractSpec.Exchange, _
                                                lCurrentContractSpec.SecType, _
                                                lCurrentContractSpec.CurrencyCode, _
                                                "1", _
                                                lCurrentContractSpec.Multiplier)
    
    Dim lContractFuture As IFuture
    Set lContractFuture = FetchContract(lNextContractSpec, _
                                        mOrderManager.ContractStorePrimary, _
                                        mOrderManager.ContractStoreSecondary)
    
    Set mNextPositionManager = mOrderManager.CreateRecoverablePositionManager( _
                                            gNextId, _
                                            mOrderManager.MarketDataManager.CreateMarketDataSource( _
                                                            lContractFuture, _
                                                            True), _
                                            mOrderManager.OrderSubmitterFactory, _
                                            mScopeName, _
                                            mBracketOrder.GroupName, _
                                            mBracketOrder.IsSimulated)

Case SecTypeOption

Case SecTypeFuturesOption

End Select



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireStateChange( _
                ByVal pState As RolloverControllerStates)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = pState

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

Private Sub placeNextBracketOrder()
Const ProcName As String = "placeNextBracketOrder"
On Error GoTo Err

Dim lEntryOrder As IOrder
Set lEntryOrder = mNextOrderContext.CreateEntryOrder(OrderTypeMarket)

Dim lStopLossOrder As IOrder
Dim lTargetOrder As IOrder
With mBracketOrder
    If Not .StopLossOrder Is Nothing Then
        Set lStopLossOrder = mNextOrderContext.CreateStopLossOrder( _
                                        .StopLossOrder.OrderType, _
                                        gNewPriceSpecifier(.StopLossOrder.LimitPriceSpec.Price, _
                                                            .StopLossOrder.LimitPriceSpec.PriceType, _
                                                            .StopLossOrder.LimitPriceSpec.Offset, _
                                                            .StopLossOrder.LimitPriceSpec.OffsetType), _
                                        gNewPriceSpecifier(.StopLossOrder.TriggerPriceSpec.Price, _
                                                            .StopLossOrder.TriggerPriceSpec.PriceType, _
                                                            .StopLossOrder.TriggerPriceSpec.Offset, _
                                                            .StopLossOrder.TriggerPriceSpec.OffsetType), _
                                        .StopLossOrder.IgnoreRegularTradingHours)
    End If
    If Not .TargetOrder Is Nothing Then
        Set lTargetOrder = mNextOrderContext.CreateTargetOrder( _
                                        .TargetOrder.OrderType, _
                                        gNewPriceSpecifier(.TargetOrder.LimitPriceSpec.Price, _
                                                            .TargetOrder.LimitPriceSpec.PriceType, _
                                                            .TargetOrder.LimitPriceSpec.Offset, _
                                                            .TargetOrder.LimitPriceSpec.OffsetType), _
                                        gNewPriceSpecifier(.TargetOrder.TriggerPriceSpec.Price, _
                                                            .TargetOrder.TriggerPriceSpec.PriceType, _
                                                            .TargetOrder.TriggerPriceSpec.Offset, _
                                                            .TargetOrder.TriggerPriceSpec.OffsetType), _
                                        .TargetOrder.IgnoreRegularTradingHours)
    End If
End With

Dim lBo As IBracketOrder
Set lBo = mNextOrderContext.CreateBracketOrder( _
                mBracketOrder.EntryOrder.Action, _
                mBracketOrder.Size, _
                lEntryOrder, _
                lStopLossOrder, _
                lTargetOrder)

lBo.Description = mBracketOrder.Description
lBo.EntryOrder.TimeInForce = mBracketOrder.EntryOrder.TimeInForce
lBo.EntryReason = mBracketOrder.EntryReason
If Not lBo.StopLossOrder Is Nothing Then
    lBo.StopLossOrder.TimeInForce = mBracketOrder.StopLossOrder.TimeInForce
    lBo.StopReason = mBracketOrder.StopReason
End If
If Not lBo.TargetOrder Is Nothing Then
    lBo.TargetOrder.TimeInForce = mBracketOrder.TargetOrder.TimeInForce
    lBo.TargetReason = mBracketOrder.TargetReason
End If
lBo.RolloverSpecification = gCreateRolloverSpecification(lBo, _
                                    mBracketOrder.RolloverSpecification.Days, _
                                    mBracketOrder.RolloverSpecification.Time, _
                                    mBracketOrder.RolloverSpecification.CloseAtMarket, _
                                    mBracketOrder.RolloverSpecification.CloseSpreadfactor, _
                                    mBracketOrder.RolloverSpecification.EntryAtMarket, _
                                    mBracketOrder.RolloverSpecification.EntrySpreadfactor)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


