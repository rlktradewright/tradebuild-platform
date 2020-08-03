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
Private WithEvents mOrderContext                    As OrderContext
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

Private mMaxExpenditure                             As Long

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mTimerList                                  As TimerList

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

Dim lChangeType As BracketOrderChangeTypes
lChangeType = ev.changeType
Select Case lChangeType
Case BracketOrderChangeTypes.BracketOrderCloseoutOrderFilled
    If mBracketOrder.Size = 0 Then
        mBracketOrder.RemoveChangeListener Me
        setState RolloverControllerStateFindingNextContract
        mTaskContext.Continue
    End If
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
                            mBracketOrder.DataSource, _
                            mBracketOrder.Contract, _
                            mBracketOrder.Key, _
                            mBracketOrder.IsSimulated, _
                            Me
    setState RolloverControllerStateClosingCurrentPosition
Case RolloverControllerStateClosingCurrentPosition
    If mBracketOrder.Size <> 0 Then
        gLog mBracketOrder.Key & ": Closing out current position: " & mBracketOrder.Contract.Specifier.ToString, ProcName, ModuleName
        gLogBracketOrderMessage "Closing out current position", _
                                mBracketOrder.DataSource, _
                                mBracketOrder.Contract, _
                                mBracketOrder.Key, _
                                mBracketOrder.IsSimulated, _
                                Me
        initiateCloseout
        setState RolloverControllerStateFindingNextContract
        mTaskContext.Suspend -1
    Else
        gLog mBracketOrder.Key & ": Current position already closed out: " & mBracketOrder.Contract.Specifier.ToString, ProcName, ModuleName
        setState RolloverControllerStateFindingNextContract
    End If
Case RolloverControllerStateFindingNextContract
    gLog mBracketOrder.Key & ": Finding next contract: " & mBracketOrder.Contract.Specifier.ToString, ProcName, ModuleName
    If Not orderContextHasBidAndAsk Then
        mTaskContext.Suspend -1
    Else
        findNextContract 1
        mTaskContext.Suspend -1
    End If
Case RolloverControllerStateCreatingNextPositionManager
    Set mNextPositionManager = createNextPositionManager(mNextContract)
    Set mNextOrderContext = mNextPositionManager.OrderContexts.DefaultOrderContext
    setState RolloverControllerStateWaitingForOrderContext
Case RolloverControllerStateWaitingForOrderContext
    If Not mNextOrderContext.IsProviderReady Then
        mTaskContext.Suspend -1
    Else
        setState RolloverControllerStateCreatingNextBracketOrder
    End If
Case RolloverControllerStateCreatingNextBracketOrder
    Set mNextBracketOrder = createNextBracketOrder(mUseMarketEntry)
    If Not mNextOrderContext.IsReadyForBracketOrder(mNextBracketOrder) Then
        gLog mBracketOrder.Key & ": Waiting for next order context: " & mNextContract.Specifier.ToString, ProcName, ModuleName
        setState RolloverControllerStateWaitingForOrderContext
        mTaskContext.Suspend -1
    Else
        setState RolloverControllerStateExecutingNextBracketOrder
    End If
Case RolloverControllerStateExecutingNextBracketOrder
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
    mNextBracketOrder.Cancel
    mTaskContext.Suspend -1
Case RolloverControllerStateCompleted
    mBracketOrder.RemoveChangeListener Me
    Set mOrderContext = Nothing
    Set mBracketOrder = Nothing
    Set mNextOrderContext = Nothing
    Set mNextBracketOrder = Nothing
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
' mOrderContext Event Handlers
'@================================================================================

Private Sub mOrderContext_Change(ev As ChangeEventData)
Const ProcName As String = "mOrderContext_Change"
On Error GoTo Err

If ev.changeType <> OrderContextChangeTypes.OrderContextReadyStateChanged Then Exit Sub

If mState = RolloverControllerStateFindingNextContract And _
    orderContextHasBidAndAsk Then mTaskContext.Continue

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
mBracketOrder.AddChangeListener Me
Set mOrderContext = mBracketOrder.OrderContext
Set mOrderManager = pOrderManager
mScopeName = pScopeName
Set mTimerList = pTimerList

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
                ByVal pUseMarketOrder As Boolean) As BracketOrder
Const ProcName As String = "createNextBracketOrder"
On Error GoTo Err

With mBracketOrder.RolloverSpecification
    Dim lEntryOrder As IOrder
    If Not pUseMarketOrder Then
        Set lEntryOrder = mNextOrderContext.CreateEntryOrder( _
                                            .EntryOrderType, _
                                            .EntryLimitPriceSpec, _
                                            .EntryTriggerPriceSpec, _
                                            mBracketOrder.EntryOrder.IgnoreRegularTradingHours)
    Else
        Set lEntryOrder = mNextOrderContext.CreateEntryOrder( _
                                            OrderTypeMarket, _
                                            gNewPriceSpecifier(), _
                                            gNewPriceSpecifier(), _
                                            mBracketOrder.EntryOrder.IgnoreRegularTradingHours)
    End If
End With

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

If mNextContract.Specifier.SecType = SecTypeFuture Then
    lBo.RolloverSpecification = mBracketOrder.RolloverSpecification
Else
    With mBracketOrder.RolloverSpecification
        Dim lStrikeValue As Long
        If .StrikeMode = RolloverStrikeModeCurrentProfit Or _
            .StrikeMode = RolloverStrikeModeCurrentValue Or _
            .StrikeMode = RolloverStrikeModePreviousAmount _
        Then
            lStrikeValue = mMaxExpenditure
        ElseIf .StrikeMode = RolloverStrikeModeAmount Or _
            .StrikeMode = RolloverStrikeModeIncrement _
        Then
            lStrikeValue = .StrikeValue
        End If
        lBo.RolloverSpecification = gCreateOptionRolloverSpecification( _
                                            .Days, _
                                            .Time, _
                                            .StrikeMode, _
                                            lStrikeValue, _
                                            .UnderlyingExchangeName, _
                                            .CloseOrderType, _
                                            .CloseTimeoutSecs, _
                                            .CloseLimitPriceSpec, _
                                            .CloseTriggerPriceSpec, _
                                            .EntryOrderType, _
                                            .EntryTimeoutSecs, _
                                            .EntryLimitPriceSpec, _
                                            .EntryTriggerPriceSpec)
    End With
End If
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
                                    mBracketOrder.IsSimulated)
    
    gLog "Created next position manager: name=" & lNextPm.Name, ProcName, ModuleName
End If

Set createNextPositionManager = lNextPm

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub findNextContract(ByVal pExpiryOffset As Long)
Const ProcName As String = "findNextContract"
On Error GoTo Err

Dim lCurrentContractSpec As IContractSpecifier
Set lCurrentContractSpec = mBracketOrder.Contract.Specifier

Dim lNextContractSpec As IContractSpecifier

Select Case lCurrentContractSpec.SecType
Case SecTypeFuture
    mFutureWaiter.Add findNextContractForFuture( _
                            lCurrentContractSpec, _
                            pExpiryOffset), _
                        pExpiryOffset
    
Case SecTypeOption, _
        SecTypeFuturesOption
    mFutureWaiter.Add findNextContractForOption( _
                            lCurrentContractSpec, _
                            pExpiryOffset), _
                        pExpiryOffset

End Select



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function findNextContractForFuture( _
                ByVal pCurrentContractSpec As IContractSpecifier, _
                ByVal pExpiryOffset As Long) As IFuture
Const ProcName As String = "findNextContractForFuture"
On Error GoTo Err

Dim lNextContractSpec As IContractSpecifier
Set lNextContractSpec = CreateContractSpecifier("", _
                            pCurrentContractSpec.Symbol, _
                            pCurrentContractSpec.Exchange, _
                            SecTypeFuture, _
                            pCurrentContractSpec.CurrencyCode, _
                            pExpiryOffset, _
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
                ByVal pExpiryOffset As Long) As IFuture
Const ProcName As String = "findNextContractForOption"
On Error GoTo Err

Dim lStrikeIncrement As Double
Dim lMultiplier As Double: lMultiplier = pCurrentContractSpec.Multiplier

Dim lBidPrice As Double
Dim lAskPrice As Double
Dim lCurrentValue As Long
If mBracketOrder.RolloverSpecification.StrikeMode = RolloverStrikeModeCurrentValue Or _
    mBracketOrder.RolloverSpecification.StrikeMode = RolloverStrikeModeCurrentProfit _
Then
    lBidPrice = mOrderContext.DataSource.CurrentTick(TickTypeBid).Price
    lAskPrice = mOrderContext.DataSource.CurrentTick(TickTypeAsk).Price

    If mBracketOrder.LongPosition Then
        lCurrentValue = CLng(lBidPrice * lMultiplier)
    Else
        lCurrentValue = CLng(lAskPrice * lMultiplier)
    End If
End If

Dim lCumBuyPrice As Double: lCumBuyPrice = mBracketOrder.CumBuyPrice
Dim lCumSellPrice As Double: lCumSellPrice = mBracketOrder.CumSellPrice
Dim lContractsTraded As Long: lContractsTraded = mBracketOrder.EntryOrder.QuantityFilled

Dim lStrikeValue As Long: lStrikeValue = mBracketOrder.RolloverSpecification.StrikeValue
Select Case mBracketOrder.RolloverSpecification.StrikeMode
Case RolloverStrikeModeNone
    mMaxExpenditure = 0
Case RolloverStrikeModeCurrentValue
    mMaxExpenditure = lCurrentValue
    gLog "Rollover strike mode current value" & _
        "; max expenditure: " & mMaxExpenditure, _
        ProcName, ModuleName
Case RolloverStrikeModeCurrentProfit
    Dim lPrice As Double
    lPrice = IIf(mBracketOrder.LongPosition, lBidPrice, lAskPrice)
    
    Dim lProfit As Long
    lProfit = CLng((lCumSellPrice + mBracketOrder.Size * lPrice - lCumBuyPrice) * lMultiplier / lContractsTraded)
    
    If lProfit <= 0 Then
        mMaxExpenditure = lCurrentValue
    Else
        mMaxExpenditure = CLng(lCurrentValue - (lStrikeValue * lProfit) / 100)
    End If
    gLog "Rollover strike mode current profit" & _
        "; max expenditure: " & mMaxExpenditure & _
        "; value: " & lCurrentValue & _
        "; profit: " & lProfit, _
        ProcName, ModuleName
Case RolloverStrikeModeAmount
    mMaxExpenditure = lStrikeValue
    gLog "Rollover strike mode amount" & _
        "; max expenditure: " & mMaxExpenditure, _
        ProcName, ModuleName
Case RolloverStrikeModePreviousAmount
    mMaxExpenditure = lStrikeValue
    gLog "Rollover strike mode previous amount" & _
        "; max expenditure: " & mMaxExpenditure, _
        ProcName, ModuleName
Case RolloverStrikeModeIncrement
    lStrikeIncrement = lStrikeValue
    gLog "Rollover strike mode increment" & _
        "; max increment by : " & lStrikeIncrement, _
        ProcName, ModuleName
End Select

Dim lNextContractSpec As IContractSpecifier
Set lNextContractSpec = CreateContractSpecifier("", _
                            pCurrentContractSpec.Symbol, _
                            pCurrentContractSpec.Exchange, _
                            pCurrentContractSpec.SecType, _
                            pCurrentContractSpec.CurrencyCode, _
                            pExpiryOffset, _
                            pCurrentContractSpec.Multiplier, _
                            IIf(mBracketOrder.RolloverSpecification.StrikeMode = RolloverStrikeModeIncrement, _
                                pCurrentContractSpec.Strike, _
                                0), _
                            pCurrentContractSpec.Right)

Set findNextContractForOption = gGetOptionContract( _
                                    lNextContractSpec, _
                                    mBracketOrder.EntryOrder.Action, _
                                    mOrderManager.ContractStorePrimary, _
                                    OptionStrikeSelectionModeExpenditure, _
                                    mMaxExpenditure, _
                                    OptionStrikeSelectionOperatorLT, _
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

Private Sub initiateCloseout()
Const ProcName As String = "initiateCloseout"
On Error GoTo Err

Dim lCloseoutSpec As New CloseoutSpecifier
With mBracketOrder.RolloverSpecification
    lCloseoutSpec.ForceCloseAfterSecs = .CloseTimeoutSecs
    lCloseoutSpec.LimitPriceSpec = .CloseLimitPriceSpec
    lCloseoutSpec.OrderType = .CloseOrderType
    lCloseoutSpec.OutsideRTH = False
    lCloseoutSpec.TriggerPriceSpec = .CloseTriggerPriceSpec
End With
mBracketOrder.DoCloseOut lCloseoutSpec

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function orderContextHasBidAndAsk() As Boolean
Const ProcName As String = "orderContextHasBidAndAsk"
On Error GoTo Err

If mBracketOrder.RolloverSpecification.StrikeMode <> RolloverStrikeModeCurrentValue And _
    mBracketOrder.RolloverSpecification.StrikeMode <> RolloverStrikeModeCurrentProfit _
Then
    orderContextHasBidAndAsk = True
ElseIf Not mOrderContext.DataSource.HasCurrentTick(TickTypeBid) Then
    orderContextHasBidAndAsk = False
ElseIf Not mOrderContext.DataSource.HasCurrentTick(TickTypeAsk) Then
    orderContextHasBidAndAsk = False
Else
    orderContextHasBidAndAsk = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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

