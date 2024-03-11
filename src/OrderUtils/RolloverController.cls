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
Private mOrderContext                               As OrderContext
Attribute mOrderContext.VB_VarHelpID = -1

Private mState                                      As RolloverControllerStates

Private mStateChangeListeners                       As New Listeners

Private mOrderManager                               As OrderManager

Private mScopeName                                  As String

Private mNextPositionManager                        As PositionManager

Private WithEvents mNextOrderContext                As OrderContext
Attribute mNextOrderContext.VB_VarHelpID = -1
Private mNextContract                               As Contract
Private WithEvents mNextBracketOrder                As BracketOrder
Attribute mNextBracketOrder.VB_VarHelpID = -1

Private mStrikeSelectionMode                        As OptionStrikeSelectionModes
Private mStrikeParameter                            As Double
Private mStrikeOperator                             As OptionStrikeSelectionOperators


Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mTimerList                                  As TimerList
Private mClock                                      As Clock

Private WithEvents mNextBracketFillTLI              As TimerListItem
Attribute mNextBracketFillTLI.VB_VarHelpID = -1

Private mUseMarketEntry                             As Boolean

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

If ev.Source Is mBracketOrder Then
    Dim lChangeType As BracketOrderChangeTypes
    lChangeType = ev.changeType
    Select Case lChangeType
    Case BracketOrderChangeTypes.BracketOrderCloseoutOrderFilled
        If mBracketOrder.Size = 0 Then
            mBracketOrder.RemoveChangeListener Me
            setState RolloverControllerStateFindingNextContract
            mTaskContext.Continue
        End If
    Case BracketOrderChangeTypes.BracketOrderCompleted
        ' the initial bracket order has completed without the closeout
        ' order being filled. For example, the contract has reached expiry
        ' and is worthless, so IB sends a cancel notification for the
        ' closeout order, but this doesn't happen till the end of the
        ' trading day. We don't proceeed with the rollover.
        gLog mBracketOrder.Key & ": Ignoring rollover - initial contract has expired: " & mBracketOrder.Contract.Specifier.ToString, ProcName, ModuleName
        setState RolloverControllerStateFailed
        mTaskContext.Finish Empty, False
    End Select
ElseIf ev.Source Is mOrderContext Then
    If isCurrentPriceAvailable Then
        mOrderContext.RemoveChangeListener Me
        mTaskContext.Continue
    End If
End If

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

Dim failpoint As String

Select Case mState

Case RolloverControllerStateInitial
    failpoint = "RolloverControllerStateInitial"
    
    If IsContractExpired(mBracketOrder.Contract, mClock) Then
        gLog mBracketOrder.Key & ": Ignoring rollover - initial contract has expired: " & mBracketOrder.Contract.Specifier.ToString, ProcName, ModuleName
        setState RolloverControllerStateFailed
        mTaskContext.Finish Empty, False
    Else
        gLog mBracketOrder.Key & ": Initiating rollover: " & mBracketOrder.Contract.Specifier.ToString, ProcName, ModuleName
        gLogBracketOrderMessage "Initiating rollover", _
                                mBracketOrder.DataSource, _
                                mBracketOrder.Contract, _
                                mBracketOrder.Key, _
                                mBracketOrder.IsSimulated, _
                                Me
        setState RolloverControllerStateClosingCurrentPosition
    End If
Case RolloverControllerStateClosingCurrentPosition
    failpoint = "RolloverControllerStateClosingCurrentPosition"
    
    If mBracketOrder.Size <> 0 Then
        gLog mBracketOrder.Key & ": Closing out current position: " & mBracketOrder.Contract.Specifier.ToString, ProcName, ModuleName
        gLogBracketOrderMessage "Closing out current position", _
                                mBracketOrder.DataSource, _
                                mBracketOrder.Contract, _
                                mBracketOrder.Key, _
                                mBracketOrder.IsSimulated, _
                                Me
        mBracketOrder.AddChangeListener Me
        initiateCloseout
        mTaskContext.Suspend -1
    Else
        gLog mBracketOrder.Key & ": Current position already closed out: " & mBracketOrder.Contract.Specifier.ToString, ProcName, ModuleName
        setState RolloverControllerStateFindingNextContract
    End If
Case RolloverControllerStateFindingNextContract
    failpoint = "RolloverControllerStateFindingNextContract"
    
    gLog mBracketOrder.Key & ": Finding next contract: " & mBracketOrder.Contract.Specifier.ToString, ProcName, ModuleName
    findNextContract mBracketOrder.RolloverSpecification.RolloverExpiryOffset, _
                    mBracketOrder.RolloverSpecification.RolloverDaysBeforeExpiryToSwitch
    mTaskContext.Suspend -1
Case RolloverControllerStateCreatingNextPositionManager
    failpoint = "RolloverControllerStateCreatingNextPositionManager"
    
    Set mNextPositionManager = createNextPositionManager(mNextContract)
    Set mNextOrderContext = mNextPositionManager.OrderContexts.DefaultOrderContext
    setState RolloverControllerStateWaitingForOrderContext
Case RolloverControllerStateWaitingForOrderContext
    failpoint = "RolloverControllerStateWaitingForOrderContext"
    
    If Not mNextOrderContext.IsProviderReady Then
        mTaskContext.Suspend -1
    Else
        setState RolloverControllerStateCreatingNextBracketOrder
    End If
Case RolloverControllerStateCreatingNextBracketOrder
    failpoint = "RolloverControllerStateCreatingNextBracketOrder"
    
    If mNextBracketOrder Is Nothing Then
        Dim lId As String: lId = GIdProvider.gNextId
        gLog mBracketOrder.Key & ": Creating next bracket order: " & lId & "; " & mNextContract.Specifier.ToString, ProcName, ModuleName
        Set mNextBracketOrder = createNextBracketOrder(lId, mUseMarketEntry, mNextContract.Specifier)
    End If
    If Not mNextOrderContext.IsReadyForBracketOrder(mNextBracketOrder) Then
        gLog mBracketOrder.Key & _
                ": Waiting for next order context: " & _
                mNextOrderContext.NotReadyMessage & ": " & _
                mNextContract.Specifier.ToString, _
                ProcName, ModuleName
        setState RolloverControllerStateWaitingForOrderContext
        mTaskContext.Suspend -1
    Else
        setState RolloverControllerStateExecutingNextBracketOrder
    End If
Case RolloverControllerStateExecutingNextBracketOrder
    failpoint = "RolloverControllerStateExecutingNextBracketOrder"
    
    mNextOrderContext.ExecuteBracketOrder mNextBracketOrder
    Set mOrderContext = Nothing
    
    If mBracketOrder.RolloverSpecification.EntryTimeoutSecs = 0 Or _
        mUseMarketEntry _
    Then
        setState RolloverControllerStateCompleted
    Else
        Set mNextBracketFillTLI = mTimerList.Add( _
                                    Nothing, _
                                    mBracketOrder.RolloverSpecification.EntryTimeoutSecs, _
                                    ExpiryTimeUnitSeconds)
        mTaskContext.Suspend -1
    End If
Case RolloverControllerStateCancellingNextBracketOrder
    failpoint = "RolloverControllerStateCancellingNextBracketOrder"
    
    mNextBracketOrder.Cancel
    mTaskContext.Suspend -1
Case RolloverControllerStateCompleted
    failpoint = "RolloverControllerStateCompleted"
    
    gLog mBracketOrder.Key & " -> " & mNextBracketOrder.Key & ": Rollover completed: " & mNextBracketOrder.Contract.Specifier.ToString, ProcName, ModuleName
    logRollover mBracketOrder, mNextBracketOrder
    mBracketOrder.RemoveChangeListener Me
    Set mOrderContext = Nothing
    Set mBracketOrder = Nothing
    Set mNextOrderContext = Nothing
    Set mNextBracketOrder = Nothing
    mTaskContext.Finish Empty, False
Case RolloverControllerStateFailed

End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, failpoint
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

If Not ev.Future.IsAvailable Then Exit Sub

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
    If lExpiryOffset = 10 Then
        ' we can't find a suitable contract
        setState RolloverControllerStateFailed
        mTaskContext.Continue
    Else
        findNextContract lExpiryOffset + 1, _
                        mBracketOrder.RolloverSpecification.RolloverDaysBeforeExpiryToSwitch
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mNextBracketFillTLI Event Handlers
'@================================================================================

Private Sub mNextBracketFillTLI_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mNextBracketFillTLI_StateChange"
On Error GoTo Err

Dim lState As TimerListItemStates: lState = ev.State
If lState <> TimerListItemStateExpired Then Exit Sub

setState RolloverControllerStateCancellingNextBracketOrder
mTaskContext.Continue

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mNextBracketOrder Event Handlers
'@================================================================================

Private Sub mNextBracketOrder_Completed()
Const ProcName As String = "mNextBracketOrder_Completed"
On Error GoTo Err

If mNextBracketOrder.Size = 0 Then
    ' the cancellation must have completed
    mUseMarketEntry = True
    setState RolloverControllerStateCreatingNextBracketOrder
    mTaskContext.Continue
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mNextBracketOrder_EntryOrderFilled()
Const ProcName As String = "mNextBracketOrder_EntryOrderFilled"
On Error GoTo Err

If Not mNextBracketFillTLI Is Nothing Then mNextBracketFillTLI.Cancel

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

If State <> RolloverControllerStateWaitingForOrderContext Then Exit Sub
If ev.changeType <> OrderContextChangeTypes.OrderContextReadyStateChanged Then Exit Sub
If Not mNextOrderContext.IsProviderReady Then Exit Sub

Set mNextContract = mNextOrderContext.ContractFuture.Value

If mNextBracketOrder Is Nothing Then
    setState RolloverControllerStateCreatingNextBracketOrder
    mTaskContext.Continue
    Exit Sub
ElseIf mNextOrderContext.IsReadyForBracketOrder(mNextBracketOrder) Then
    setState RolloverControllerStateExecutingNextBracketOrder
    mTaskContext.Continue
Else
    gLog mBracketOrder.Key & _
            ": Waiting for next order context: " & _
            mNextOrderContext.NotReadyMessage & ": " & _
            mNextContract.Specifier.ToString, _
            ProcName, ModuleName
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

Friend Sub Cancel()
Const ProcName As String = "Cancel"
On Error GoTo Err

mBracketOrder.RemoveChangeListener Me

mFutureWaiter.Cancel
If Not mNextBracketOrder Is Nothing Then mNextBracketOrder.CloseOut
If Not mNextBracketFillTLI Is Nothing Then mNextBracketFillTLI.Cancel

Set mOrderContext = Nothing
Set mBracketOrder = Nothing
Set mNextOrderContext = Nothing
Set mNextBracketOrder = Nothing
Set mNextBracketFillTLI = Nothing

mTaskContext.Finish Empty, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pOrderManager As OrderManager, _
                ByVal pScopeName As String, _
                ByVal pTimerList As TimerList)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mBracketOrder = pBracketOrder
Set mOrderContext = mBracketOrder.OrderContext
Set mOrderManager = pOrderManager
mScopeName = pScopeName
Set mTimerList = pTimerList
Set mClock = mTimerList.Clock

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

Private Function createNextBracketOrder( _
                ByVal pId As String, _
                ByVal pUseMarketOrder As Boolean, _
                ByVal pCurrentContractSpec As IContractSpecifier) As BracketOrder
Const ProcName As String = "createNextBracketOrder"
On Error GoTo Err

Dim lEntryOrder As IOrder
Set lEntryOrder = createNextEntryOrder(pUseMarketOrder)

Dim lStopLossOrder As IOrder
Set lStopLossOrder = createNextStoplossOrder

Dim lTargetOrder As IOrder
Set lTargetOrder = createNextTargetOrder

Dim lAutosizeType As BracketAutosizeTypes
Dim lAutosizeValue As Double
Dim lQuantity As BoxedDecimal
Dim lLotSize As Long

getQuantityParameters pCurrentContractSpec, lAutosizeType, lAutosizeValue, lQuantity, lLotSize

Dim lBo As BracketOrder
Set lBo = mNextOrderContext.CreateBracketOrderWithId( _
                mBracketOrder.EntryOrder.Action, _
                lQuantity, _
                lEntryOrder, _
                pId, _
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

If lAutosizeType <> BracketAutosizeNone Then
    lBo.AutosizeType = lAutosizeType
    lBo.AutosizeValue = lAutosizeValue
    lBo.LotSize = lLotSize
End If

lBo.RolloverSpecification = createNextRolloverSpec

gLog lBo.Key & ": Created next bracket order: " & lBo.Contract.Specifier.ToString, ProcName, ModuleName
Set createNextBracketOrder = lBo

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createNextEntryOrder( _
                ByVal pUseMarketOrder As Boolean) As IOrder
Const ProcName As String = "createNextEntryOrder"
On Error GoTo Err

Dim lPrevEntryOrder As IOrder
Set lPrevEntryOrder = mBracketOrder.EntryOrder

With mBracketOrder.RolloverSpecification
    If .EntryOrderType = OrderTypeNone Then
        Set createNextEntryOrder = mNextOrderContext.CreateEntryOrder( _
                                            lPrevEntryOrder.OrderType, _
                                            .EntryLimitPriceSpec, _
                                            .EntryTriggerPriceSpec, _
                                            lPrevEntryOrder.IgnoreRegularTradingHours)
    ElseIf Not pUseMarketOrder Then
        Set createNextEntryOrder = mNextOrderContext.CreateEntryOrder( _
                                            .EntryOrderType, _
                                            .EntryLimitPriceSpec, _
                                            .EntryTriggerPriceSpec, _
                                            lPrevEntryOrder.IgnoreRegularTradingHours)
    Else
        Set createNextEntryOrder = mNextOrderContext.CreateEntryOrder( _
                                            OrderTypeMarket, _
                                            gNewPriceSpecifier(), _
                                            gNewPriceSpecifier(), _
                                            lPrevEntryOrder.IgnoreRegularTradingHours)
    End If
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createNextPositionManager( _
                ByVal pContract As IContract) As PositionManager
Const ProcName As String = "createNextPositionManager"
On Error GoTo Err

Dim lNextPm As PositionManager

Dim lPositionManagersForGroup As PositionManagers
Set lPositionManagersForGroup = mOrderManager.GetPositionManagersForGroup(mBracketOrder.GroupName)

Dim lPm As PositionManager
For Each lPm In lPositionManagersForGroup
    Dim lContract As IContract: Set lContract = lPm.ContractFuture.Value
    If ContractSpecsEqual(lContract.Specifier, pContract.Specifier) Then
        Set lNextPm = lPm
        Exit For
    End If
Next

If Not lNextPm Is Nothing Then
    gLog "Using existing position manager: name=" & lNextPm.Name, ProcName, ModuleName
Else
    Dim lDataSource As IMarketDataSource
    Set lDataSource = mOrderManager.MarketDataManager.CreateMarketDataSource( _
                                            CreateFuture(pContract), _
                                            True)
    lDataSource.StartMarketData
    
    Set lNextPm = mOrderManager.CreateRecoverablePositionManager( _
                                    gNextId, _
                                    lDataSource, _
                                    mOrderManager.OrderSubmitterFactory, _
                                    mScopeName, _
                                    mBracketOrder.GroupName, _
                                    mBracketOrder.IsSimulated, _
                                    , _
                                    mOrderContext.AccountDataProvider, _
                                    mOrderContext.CurrencyConverter)
    lNextPm.FixedAccountBalance = mOrderContext.FixedAccountBalance
    
    gLog "Created next position manager: name=" & lNextPm.Name, ProcName, ModuleName
End If

Set createNextPositionManager = lNextPm

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createNextRolloverSpec() As RolloverSpecification
Const ProcName As String = "createNextRolloverSpec"
On Error GoTo Err

If mNextContract.Specifier.SecType = SecTypeFuture Then
    Set createNextRolloverSpec = mBracketOrder.RolloverSpecification
Else
    With mBracketOrder.RolloverSpecification
        Dim lInitialStrikeParameter As Double
        
        Select Case .RolloverStrikeSelectionMode
        Case RolloverStrikeModeNone, _
                RolloverStrikeModeAsPrevious
            lInitialStrikeParameter = .InitialStrikeParameter
        Case RolloverStrikeModeCurrentValue, _
                RolloverStrikeModeCurrentProfit, _
                RolloverStrikeModePreviousMonetaryAmount
            lInitialStrikeParameter = mStrikeParameter
        Case RolloverStrikeModeMonetaryAmount, _
                RolloverStrikeModeIncrement
            lInitialStrikeParameter = .RolloverStrikeParameter
        Case RolloverStrikeModeDelta
            lInitialStrikeParameter = .RolloverStrikeParameter
        Case RolloverStrikeModePreviousDelta
            lInitialStrikeParameter = mStrikeParameter
        Case RolloverStrikeModeDeltaIncrement
            lInitialStrikeParameter = mStrikeParameter
        End Select
                
        Set createNextRolloverSpec = gCreateOptionRolloverSpecification( _
                                            .Days, _
                                            .Time, _
                                            .InitialStrikeSelectionMode, _
                                            lInitialStrikeParameter, _
                                            .InitialStrikeOperator, _
                                            .RolloverStrikeSelectionMode, _
                                            .RolloverStrikeParameter, _
                                            .RolloverStrikeSelectionMode, _
                                            .RolloverQuantityMode, _
                                            .RolloverQuantityParameter, _
                                            .RolloverQuantityLotSize, _
                                            .UnderlyingExchangeName, _
                                            .CloseOrderType, _
                                            .CloseTimeoutSecs, _
                                            .CloseLimitPriceSpec, _
                                            .CloseTriggerPriceSpec, _
                                            .EntryOrderType, _
                                            .EntryTimeoutSecs, _
                                            .EntryLimitPriceSpec, _
                                            .EntryTriggerPriceSpec)
        createNextRolloverSpec.setStopLossLimitPriceSpec .StopLossLimitPriceSpec
        createNextRolloverSpec.setStopLossTriggerPriceSpec .StopLossTriggerPriceSpec
        createNextRolloverSpec.setTargetLimitPriceSpec .TargetLimitPriceSpec
        createNextRolloverSpec.setTargetTriggerPriceSpec .TargetTriggerPriceSpec
        createNextRolloverSpec.setRolloverDaysBeforeExpiryToSwitch .RolloverDaysBeforeExpiryToSwitch
        createNextRolloverSpec.setRolloverExpiryOffset .RolloverExpiryOffset
    End With
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createNextStoplossOrder() As IOrder
Const ProcName As String = "createNextStoplossOrder"
On Error GoTo Err

If mBracketOrder.StopLossOrder Is Nothing Then Exit Function

With mBracketOrder.RolloverSpecification
    Set createNextStoplossOrder = mNextOrderContext.CreateStopLossOrder( _
                                    .StopLossOrderType, _
                                    .StopLossLimitPriceSpec, _
                                    .StopLossTriggerPriceSpec, _
                                    mBracketOrder.StopLossOrder.IgnoreRegularTradingHours)
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createNextTargetOrder() As IOrder
Const ProcName As String = "createNextTargetOrder"
On Error GoTo Err

If mBracketOrder.TargetOrder Is Nothing Then Exit Function

With mBracketOrder.RolloverSpecification
    Set createNextTargetOrder = mNextOrderContext.CreateTargetOrder( _
                                    .TargetOrderType, _
                                    .TargetLimitPriceSpec, _
                                    .TargetTriggerPriceSpec, _
                                    mBracketOrder.TargetOrder.IgnoreRegularTradingHours)
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub findNextContract( _
                ByVal pExpiryOffset As Long, _
                ByVal pDaysBeforeExpiryToSwitch As Long)
Const ProcName As String = "findNextContract"
On Error GoTo Err

Dim lCurrentContractSpec As IContractSpecifier
Set lCurrentContractSpec = mBracketOrder.Contract.Specifier

Dim lNextContractSpec As IContractSpecifier

Select Case lCurrentContractSpec.SecType
Case SecTypeFuture
    mFutureWaiter.Add findNextContractForFuture( _
                            lCurrentContractSpec, _
                            pExpiryOffset, _
                            pDaysBeforeExpiryToSwitch), _
                        pExpiryOffset
    
Case SecTypeOption, _
        SecTypeFuturesOption
    mFutureWaiter.Add findNextContractForOption( _
                            lCurrentContractSpec, _
                            pExpiryOffset, _
                            pDaysBeforeExpiryToSwitch), _
                        pExpiryOffset

End Select



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function findNextContractForFuture( _
                ByVal pCurrentContractSpec As IContractSpecifier, _
                ByVal pExpiryOffset As Long, _
                ByVal pDaysBeforeExpiryToSwitch As Long) As IFuture
Const ProcName As String = "findNextContractForFuture"
On Error GoTo Err

Dim lNextContractSpec As IContractSpecifier
Set lNextContractSpec = CreateContractSpecifier("", _
                            pCurrentContractSpec.Symbol, _
                            pCurrentContractSpec.TradingClass, _
                            pCurrentContractSpec.Exchange, _
                            SecTypeFuture, _
                            pCurrentContractSpec.CurrencyCode, _
                            CStr(pExpiryOffset) & "[" & CStr(pDaysBeforeExpiryToSwitch) & "d]", _
                            pCurrentContractSpec.Multiplier)

Set findNextContractForFuture = FetchContract( _
                                    lNextContractSpec, _
                                    mOrderManager.ContractStorePrimary, _
                                    mOrderManager.ContractStoreSecondary)


Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function findNextContractForOption( _
                ByVal pCurrentContractSpec As IContractSpecifier, _
                ByVal pExpiryOffset As Long, _
                ByVal pDaysBeforeExpiryToSwitch As Long) As IFuture
Const ProcName As String = "findNextContractForOption"
On Error GoTo Err

Dim lRolloverSpec As RolloverSpecification
Set lRolloverSpec = mBracketOrder.RolloverSpecification

Dim lStrikeSelectionMode As RolloverStrikeModes
lStrikeSelectionMode = lRolloverSpec.RolloverStrikeSelectionMode

Select Case lStrikeSelectionMode
Case RolloverStrikeModeNone, _
        RolloverStrikeModeAsPrevious
    mStrikeSelectionMode = lRolloverSpec.InitialStrikeSelectionMode
    mStrikeParameter = lRolloverSpec.InitialStrikeParameter
    mStrikeOperator = lRolloverSpec.InitialStrikeOperator
    gLog "Rollover strike mode as previous: " & _
        gOptionStrikeSelectionModeToString(mStrikeSelectionMode) & _
        "; parameter: " & mStrikeParameter & _
        "; operator: " & gOptionStrikeSelectionOperatorToString(mStrikeOperator), _
        ProcName, ModuleName
Case RolloverStrikeModeCurrentValue
    If Not isCurrentPriceAvailable Then
        mOrderContext.AddChangeListener Me
        mTaskContext.Suspend -1
        Exit Function
    End If

    mStrikeSelectionMode = OptionStrikeSelectionModeExpenditure
    mStrikeParameter = getCurrentValue(pCurrentContractSpec.Multiplier)
    mStrikeOperator = lRolloverSpec.InitialStrikeOperator
    gLog "Rollover strike mode current value" & _
        "; max expenditure: " & mStrikeParameter & _
        "; operator: " & gOptionStrikeSelectionOperatorToString(mStrikeOperator), _
        ProcName, ModuleName
Case RolloverStrikeModeCurrentProfit
    If Not isCurrentPriceAvailable Then
        mOrderContext.AddChangeListener Me
        mTaskContext.Suspend -1
        Exit Function
    End If

    mStrikeSelectionMode = OptionStrikeSelectionModeExpenditure
    mStrikeOperator = lRolloverSpec.InitialStrikeOperator
    
    Dim lValue As Long: lValue = getCurrentValue(pCurrentContractSpec.Multiplier)
    Dim lProfit As Long: lProfit = getCurrentProfit(pCurrentContractSpec.Multiplier)
    If lProfit <= 0 Then
        mStrikeParameter = lValue
    Else
        mStrikeParameter = CLng(lValue - (lValue * lProfit) / 100)
    End If
    gLog "Rollover strike mode current profit" & _
        "; max expenditure: " & mStrikeParameter & _
        "; value: " & lValue & _
        "; profit: " & lProfit & _
        "; operator: " & gOptionStrikeSelectionOperatorToString(mStrikeOperator), _
        ProcName, ModuleName
Case RolloverStrikeModeMonetaryAmount
    mStrikeSelectionMode = OptionStrikeSelectionModeExpenditure
    mStrikeParameter = lRolloverSpec.RolloverStrikeParameter
    mStrikeOperator = lRolloverSpec.RolloverStrikeOperator
    gLog "Rollover strike mode amount" & _
        "; max expenditure: " & mStrikeParameter & _
        "; operator: " & gOptionStrikeSelectionOperatorToString(mStrikeOperator), _
        ProcName, ModuleName
Case RolloverStrikeModePreviousMonetaryAmount
    mStrikeSelectionMode = OptionStrikeSelectionModeExpenditure
    mStrikeParameter = lRolloverSpec.InitialStrikeParameter
    mStrikeOperator = lRolloverSpec.InitialStrikeOperator
    gLog "Rollover strike mode previous amount" & _
        "; max expenditure: " & mStrikeParameter & _
        "; operator: " & gOptionStrikeSelectionOperatorToString(mStrikeOperator), _
        ProcName, ModuleName
Case RolloverStrikeModeIncrement
    Assert False, "Rollover via strike increment not supported", ErrorCodes.ErrUnsupportedOperationException
Case RolloverStrikeModeDelta
    mStrikeSelectionMode = OptionStrikeSelectionModeDelta
    mStrikeParameter = lRolloverSpec.RolloverStrikeParameter
    mStrikeOperator = lRolloverSpec.RolloverStrikeOperator
    gLog "Rollover strike mode delta" & _
        "; delta: " & mStrikeParameter & _
        "; operator: " & gOptionStrikeSelectionOperatorToString(mStrikeOperator), _
        ProcName, ModuleName
Case RolloverStrikeModePreviousDelta
    mStrikeSelectionMode = OptionStrikeSelectionModeDelta
    mStrikeParameter = lRolloverSpec.InitialStrikeParameter
    mStrikeOperator = lRolloverSpec.RolloverStrikeOperator
    gLog "Rollover strike mode previous delta" & _
        "; delta: " & mStrikeParameter & _
        "; operator: " & gOptionStrikeSelectionOperatorToString(mStrikeOperator), _
        ProcName, ModuleName
Case RolloverStrikeModeDeltaIncrement
    mStrikeSelectionMode = OptionStrikeSelectionModeDelta
    mStrikeParameter = lRolloverSpec.InitialStrikeParameter + _
                        lRolloverSpec.RolloverStrikeParameter
    mStrikeOperator = lRolloverSpec.RolloverStrikeOperator
    gLog "Rollover strike mode previous delta" & _
        "; delta: " & mStrikeParameter & _
        "; operator: " & gOptionStrikeSelectionOperatorToString(mStrikeOperator), _
        ProcName, ModuleName
End Select

Dim lNextContractSpec As IContractSpecifier
Set lNextContractSpec = CreateContractSpecifier("", _
                            pCurrentContractSpec.Symbol, _
                            pCurrentContractSpec.TradingClass, _
                            pCurrentContractSpec.Exchange, _
                            pCurrentContractSpec.SecType, _
                            pCurrentContractSpec.CurrencyCode, _
                            CStr(pExpiryOffset) & "[" & CStr(pDaysBeforeExpiryToSwitch) & "d]", _
                            pCurrentContractSpec.Multiplier, _
                            IIf(mBracketOrder.RolloverSpecification.RolloverStrikeSelectionMode = RolloverStrikeModeIncrement, _
                                pCurrentContractSpec.Strike, _
                                0), _
                            pCurrentContractSpec.Right)

Set findNextContractForOption = gGetOptionContract( _
                                    lNextContractSpec, _
                                    mBracketOrder.EntryOrder.Action, _
                                    mOrderManager.ContractStorePrimary, _
                                    mStrikeSelectionMode, _
                                    mStrikeParameter, _
                                    mStrikeOperator, _
                                    mBracketOrder.RolloverSpecification.UnderlyingExchangeName, _
                                    mOrderManager.MarketDataManager, _
                                    Nothing, _
                                    Globals.MinDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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

Private Sub getQuantityParameters( _
                ByVal pCurrentContractSpec As IContractSpecifier, _
                ByRef pAutosizeType As BracketAutosizeTypes, _
                ByRef pAutosizeValue As Double, _
                ByRef pQuantity As BoxedDecimal, _
                ByRef pLotSize As Long)
Const ProcName As String = "getQuantityParameters"
On Error GoTo Err

Dim lRolloverSpec As RolloverSpecification
Set lRolloverSpec = mBracketOrder.RolloverSpecification

Select Case lRolloverSpec.RolloverQuantityMode
Case RolloverQuantityModeNone, _
        RolloverQuantityModeAsPrevious
    If mBracketOrder.AutosizeType = BracketAutosizeNone Then
        pAutosizeType = BracketAutosizeNone
        Set pQuantity = mBracketOrder.EntryOrder.QuantityFilled
    Else
        pAutosizeType = mBracketOrder.AutosizeType
        pAutosizeValue = mBracketOrder.AutosizeValue
        pLotSize = mBracketOrder.LotSize
        Set pQuantity = DecimalOne
    End If
Case RolloverQuantityModeNumber
    Set pQuantity = lRolloverSpec.RolloverQuantityParameter
Case RolloverQuantityModePreviousNumber
    Set pQuantity = mBracketOrder.EntryOrder.QuantityFilled
Case RolloverQuantityModeMonetaryAmount
    pAutosizeType = BracketAutosizeMonetaryAmount
    pAutosizeValue = lRolloverSpec.RolloverQuantityParameter
    pLotSize = lRolloverSpec.RolloverQuantityLotSize
Case RolloverQuantityModePreviousMonetaryAmount
    pAutosizeType = BracketAutosizeMonetaryAmount
    pAutosizeValue = mBracketOrder.AutosizeValue
    pLotSize = mBracketOrder.LotSize
Case RolloverQuantityModePercentageOfAccount
    pAutosizeType = BracketAutosizePercentOfAccount
    pAutosizeValue = lRolloverSpec.RolloverQuantityParameter
    pLotSize = lRolloverSpec.RolloverQuantityLotSize
Case RolloverQuantityModePreviousPercentageOfAccount
    pAutosizeType = BracketAutosizePercentOfAccount
    pAutosizeValue = mBracketOrder.AutosizeValue
    pLotSize = mBracketOrder.LotSize
Case RolloverQuantityModeCurrentValue
    pAutosizeType = BracketAutosizeMonetaryAmount
    pAutosizeValue = getCurrentValue(pCurrentContractSpec.Multiplier)
    pLotSize = lRolloverSpec.RolloverQuantityLotSize
Case RolloverQuantityModeCurrentProfit
    pAutosizeType = BracketAutosizeMonetaryAmount
    pAutosizeValue = getCurrentValue(pCurrentContractSpec.Multiplier) - _
                    CLng(lRolloverSpec.RolloverQuantityParameter * _
                        getCurrentProfit(pCurrentContractSpec.Multiplier) / 100#)
    pLotSize = lRolloverSpec.RolloverQuantityLotSize
End Select


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getCurrentValue( _
                ByVal pMultiplier As Double) As Long
Const ProcName As String = "getCurrentValue"
On Error GoTo Err

Assert isCurrentPriceAvailable, "Price is not currently available"

getCurrentValue = CLng(getCurrentPrice * pMultiplier)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCurrentProfit(ByVal pMultiplier As Double) As Long
Dim lCumBuyPrice As Double: lCumBuyPrice = mBracketOrder.CumBuyPrice
Dim lCumSellPrice As Double: lCumSellPrice = mBracketOrder.CumSellPrice
Dim lContractsTraded As Long: lContractsTraded = mBracketOrder.EntryOrder.QuantityFilled
getCurrentProfit = CLng((lCumSellPrice + mBracketOrder.Size * getCurrentPrice - lCumBuyPrice) * pMultiplier / lContractsTraded)
End Function

Private Function getCurrentPrice() As Double
Const ProcName As String = "getCurrentPrice"
On Error GoTo Err

If mBracketOrder.LongPosition Then
    getCurrentPrice = mOrderContext.DataSource.CurrentTick(TickTypeBid).Price
Else
    getCurrentPrice = mOrderContext.DataSource.CurrentTick(TickTypeAsk).Price
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub initiateCloseout()
Const ProcName As String = "initiateCloseout"
On Error GoTo Err

Dim lCloseoutSpec As New CloseoutSpecifier
With mBracketOrder.RolloverSpecification
    lCloseoutSpec.ForceCloseAfterSecs = .CloseTimeoutSecs
    lCloseoutSpec.LimitPriceSpec = .CloseLimitPriceSpec
    lCloseoutSpec.OrderType = .CloseOrderType
    If lCloseoutSpec.OrderType = OrderTypeNone Then lCloseoutSpec.OrderType = OrderTypeMarket
    lCloseoutSpec.OutsideRTH = False
    lCloseoutSpec.TriggerPriceSpec = .CloseTriggerPriceSpec
End With
mBracketOrder.DoCloseOut lCloseoutSpec

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function isCurrentPriceAvailable() As Boolean
Const ProcName As String = "isCurrentPriceAvailable"
On Error GoTo Err

Assert mBracketOrder.RolloverSpecification.RolloverStrikeSelectionMode = RolloverStrikeModeCurrentValue Or _
            mBracketOrder.RolloverSpecification.RolloverStrikeSelectionMode = RolloverStrikeModeCurrentProfit, _
        "Current price not needed for this strike selection mode"

If Not mOrderContext.DataSource.HasCurrentTick(TickTypeBid) Then
    isCurrentPriceAvailable = False
ElseIf Not mOrderContext.DataSource.HasCurrentTick(TickTypeAsk) Then
    isCurrentPriceAvailable = False
Else
    isCurrentPriceAvailable = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub logRollover( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pNextBracketOrder As IBracketOrder)
Const ProcName As String = "logRollover"
On Error GoTo Err

gLogBracketOrderRollover _
                pBracketOrder.GroupName & "," & _
                FormatTimestamp(mNextOrderContext.Clock.Timestamp, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & "," & _
                mNextOrderContext.Clock.TimeZone.StandardName & "," & _
                pBracketOrder.Key & "," & _
                gGetContractName(pBracketOrder.Contract.Specifier) & "," & _
                pNextBracketOrder.Key & "," & _
                gGetContractName(pNextBracketOrder.Contract.Specifier), _
                pNextBracketOrder.IsSimulated, _
                Me


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

