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
Implements ITask

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

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

Private mBracketOrder                               As BracketOrder
Attribute mBracketOrder.VB_VarHelpID = -1

Private mState                                      As RolloverControllerStates

Private mStateChangeListeners                       As New Listeners

Private mOrderManager                               As OrderManager

Private mScopeName                                  As String

Private mNextPositionManager                        As PositionManager

Private WithEvents mNextOrderContext                As OrderContext
Attribute mNextOrderContext.VB_VarHelpID = -1
Private mNextContract                               As Contract
Private mNextBracketOrder                           As BracketOrder

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
setState RolloverControllerStateInitial
Set mFutureWaiter = New FutureWaiter
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
    setState RolloverControllerStateFindingNextContract
    mTaskContext.Continue
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
    gLog mBracketOrder.Key & ": Initiating rollover: " & mBracketOrder.Contract.Specifier.ToString, ProcName, ModuleName
    gLogBracketOrderMessage "Initiating rollover", _
                            Nothing, _
                            mBracketOrder.Contract, _
                            mBracketOrder.Key, _
                            mBracketOrder.IsSimulated, _
                            Me
    setState RolloverControllerStateClosingCurrentPosition
Case RolloverControllerStateClosingCurrentPosition
    If mBracketOrder.Size <> 0 Then
        gLog mBracketOrder.Key & ": Closing out current position: " & mBracketOrder.Contract.Specifier.ToString, ProcName, ModuleName
        gLogBracketOrderMessage "Closing out current position", _
                                Nothing, _
                                mBracketOrder.Contract, _
                                mBracketOrder.Key, _
                                mBracketOrder.IsSimulated, _
                                Me
        mBracketOrder.CloseOut
        setState RolloverControllerStateFindingNextContract
        mTaskContext.Suspend -1
    Else
        gLog mBracketOrder.Key & ": Current position already closed out: " & mBracketOrder.Contract.Specifier.ToString, ProcName, ModuleName
        setState RolloverControllerStateFindingNextContract
    End If
Case RolloverControllerStateFindingNextContract
    gLog mBracketOrder.Key & ": Finding next contract: " & mBracketOrder.Contract.Specifier.ToString, ProcName, ModuleName
    findNextContract 1
    mTaskContext.Suspend -1
Case RolloverControllerStateCreatingNextPositionManager
    gLog mBracketOrder.Key & ": Creating next position manager: " & mNextContract.Specifier.ToString, ProcName, ModuleName
    
    Dim lContractFuture As IFuture
    Dim lDataSource As IMarketDataSource
    Set lDataSource = mOrderManager.MarketDataManager.CreateMarketDataSource( _
                                        CreateFuture(mNextContract), _
                                        True)
    lDataSource.StartMarketData

    Set mNextPositionManager = mOrderManager.CreateRecoverablePositionManager( _
                                                    GIdProvider.gNextId, _
                                                    lDataSource, _
                                                    mOrderManager.OrderSubmitterFactory, _
                                                    mScopeName, _
                                                    mBracketOrder.GroupName, _
                                                    mBracketOrder.IsSimulated)
    Set mNextOrderContext = mNextPositionManager.OrderContexts.DefaultOrderContext
    setState RolloverControllerStateWaitingForOrderContext
Case RolloverControllerStateWaitingForOrderContext
    If Not mNextOrderContext.IsProviderReady Then
        mTaskContext.Suspend -1
    Else
        setState RolloverControllerStateCreatingNextBracketOrder
    End If
Case RolloverControllerStateCreatingNextBracketOrder
    Set mNextBracketOrder = createNextBracketOrder
    If Not mNextOrderContext.IsReadyForBracketOrder(mNextBracketOrder) Then
        gLog mBracketOrder.Key & ": Waiting for next order context: " & mNextContract.Specifier.ToString, ProcName, ModuleName
        setState RolloverControllerStateWaitingForOrderContext
        mTaskContext.Suspend -1
    Else
        setState RolloverControllerStateExecutingNextBracketOrder
    End If
Case RolloverControllerStateExecutingNextBracketOrder
    mNextOrderContext.ExecuteBracketOrder mNextBracketOrder
    Set mNextOrderContext = Nothing
    setState RolloverControllerStateCompleted
    mTaskContext.Finish Empty, False
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
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsAvailable Then
    Dim lContract As IContract: Set lContract = ev.Future.Value
    If lContract.ExpiryDate > mBracketOrder.Contract.ExpiryDate Then
        ' we've found the next contract
        Set mNextContract = lContract
        gLog mBracketOrder.Key & ": Got rollover contract: " & mNextContract.Specifier.ToString, ProcName, ModuleName
        gLogBracketOrderMessage "Found next contract", _
                                Nothing, _
                                mNextContract, _
                                "", _
                                mBracketOrder.IsSimulated, _
                                Me
        setState RolloverControllerStateCreatingNextPositionManager
        mTaskContext.Continue
    Else
        Dim lExpiryOffset As Long: lExpiryOffset = ev.ContinuationData
        findNextContract lExpiryOffset + 1
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mNextOrderContext Event Handlers
'@================================================================================

Private Sub mNextOrderContext_Change(ev As ChangeEventData)
Const ProcName As String = "mNextOrderContext_Change"
On Error GoTo Err

If ev.changeType <> OrderContextChangeTypes.OrderContextReadyStateChanged Then Exit Sub

If Not mNextOrderContext.IsProviderReady Then
    Exit Sub
End If

Set mNextContract = mNextOrderContext.ContractFuture.Value

If mNextBracketOrder Is Nothing Then
    setState RolloverControllerStateCreatingNextBracketOrder
    mTaskContext.Continue
    Exit Sub
ElseIf mNextOrderContext.IsReadyForBracketOrder(mNextBracketOrder) Then
    setState RolloverControllerStateExecutingNextBracketOrder
    mTaskContext.Continue
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get NextBracketOrder() As IBracketOrder
Set NextBracketOrder = mNextBracketOrder
End Property

Public Property Get NextContract() As IContract
Set NextContract = mNextContract
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pOrderManager As OrderManager, _
                ByVal pScopeName As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mBracketOrder = pBracketOrder
mBracketOrder.AddChangeListener Me
Set mOrderManager = pOrderManager
mScopeName = pScopeName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createNextBracketOrder() As BracketOrder
Const ProcName As String = "createNextBracketOrder"
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
                mBracketOrder.EntryOrder.QuantityFilled, _
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
lBo.RolloverSpecification = gCreateRolloverSpecification( _
                                    mBracketOrder.RolloverSpecification.Days, _
                                    mBracketOrder.RolloverSpecification.Time, _
                                    mBracketOrder.RolloverSpecification.CloseOrderType, _
                                    mBracketOrder.RolloverSpecification.CloseLimitPriceSpec, _
                                    mBracketOrder.RolloverSpecification.CloseTriggerPriceSpec, _
                                    mBracketOrder.RolloverSpecification.EntryOrderType, _
                                    mBracketOrder.RolloverSpecification.EntryLimitPriceSpec, _
                                    mBracketOrder.RolloverSpecification.EntryTriggerPriceSpec)

gLog lBo.Key & ": Created next bracket order: " & lBo.Contract.Specifier.ToString, ProcName, ModuleName
Set createNextBracketOrder = lBo

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createNextPositionManager( _
                ByVal pContract As IContract) As PositionManager
Const ProcName As String = "createNextPositionManager"
On Error GoTo Err

Dim lDataSource As IMarketDataSource
Set lDataSource = mOrderManager.MarketDataManager.CreateMarketDataSource( _
                                        CreateFuture(pContract), _
                                        True)
lDataSource.StartMarketData

Dim lPm As PositionManager
Set lPm = mOrderManager.CreateRecoverablePositionManager( _
                                gNextId, _
                                lDataSource, _
                                mOrderManager.OrderSubmitterFactory, _
                                mScopeName, _
                                mBracketOrder.GroupName, _
                                mBracketOrder.IsSimulated)

gLog "Created next position manager: name=" & lPm.Name, ProcName, ModuleName

Set createNextPositionManager = lPm

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub findNextContract(ByVal pExpiryOffset As Long)
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
                                                pExpiryOffset, _
                                                lCurrentContractSpec.Multiplier)
    
    mFutureWaiter.Add FetchContract(lNextContractSpec, _
                                        mOrderManager.ContractStorePrimary, _
                                        mOrderManager.ContractStoreSecondary), _
                                    pExpiryOffset
    
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

Private Sub setState(ByVal pState As RolloverControllerStates)
Const ProcName As String = "setState"
On Error GoTo Err

If mState = pState Then Exit Sub

mState = pState
fireStateChange mState

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
