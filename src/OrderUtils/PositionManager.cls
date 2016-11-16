VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PositionManager"
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
' Interfaces
'@================================================================================

Implements IChangeListener
Implements IGenericTickListener
Implements IOrderAuthoriser
Implements IOrderSubmissionListener
Implements IPositionMember
Implements IRecoverable
Implements IStateChangeListener

'@================================================================================
' Events
'@================================================================================

Event OrderError(ByVal pOrderId As String, ByVal pErrorCode As Long, ByVal pErrorMsg As String)

Event Change(ev As ChangeEventData)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                        As String = "PositionManager"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mSummary                                As PositionManagerSummary

Private mOrderManager                           As OrderManager

Private mContractFuture                         As IFuture
Private mOrderSubmitter                         As IOrderSubmitter
Private mMoneyManager                           As IMoneyManager
Private mDataSource                             As IMarketDataSource

Private mSessionFuture                          As IFuture
Private WithEvents mSession                     As Session
Attribute mSession.VB_VarHelpID = -1

Private mIsInitialised                          As Boolean
Private mIsFinished                             As Boolean

Private mApplicationIndex                       As Long

Private mProviderIsReady                        As Boolean

Private WithEvents mBracketOrders               As BracketOrders
Attribute mBracketOrders.VB_VarHelpID = -1
Private WithEvents mOrders                      As Orders
Attribute mOrders.VB_VarHelpID = -1
Private mExecutions                             As Executions

Private WithEvents mOrderContexts               As OrderContexts
Attribute mOrderContexts.VB_VarHelpID = -1

' Indicates that money management has told us to end the Session. Note
' that this can happen whilst there are still some OrderContexts in progress,
' for example if the daily Profit has been met. In this case, the Session
' is ended when all OrderContexts have completed.
Private mEndingSession As Boolean

Private mQuantityOutstanding                    As Long

Private mSessionProfit                          As Currency
Private mMaxSessionProfit                       As Currency
Private mMaxSessionLoss                         As Currency
Private mSessionDrawdown                        As Currency

Private mProfitThisTrade                        As Currency
Private mMaxProfitThisTrade                     As Currency
Private mMaxLossThisTrade                       As Currency
Private mDrawdownThisTrade                      As Currency
Private mTradeStartTime                         As Date
Private mTradeStartTimezoneName                 As String

Private mTickValue                              As Currency
Private mTickSize                               As Double
Private mCurrencyCode                           As String

Private mLastBid                                As Double
Private mLastAsk                                As Double

Private mProfitListeners                        As Listeners
Private mChangeListeners                        As Listeners

Private mRecovering                             As Boolean

Private WithEvents mClosePositionsTask          As ClosePositionsTask
Attribute mClosePositionsTask.VB_VarHelpID = -1

' Indicates that all OrderContexts are being closed, either as a result of
' a call to the closePosition functions, or because money management
' has told us to do so
Private mClosingPosition                        As Boolean

Private mRecoverableObjectListeners             As New Listeners

Private mFinishRecoveryPending                  As Boolean

Private mBracketOrderRecoveryController         As BracketOrderRecoveryCtlr

Private mInitialRecoveryDataSaved               As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mSummary = New PositionManagerSummary

mApplicationIndex = GPositionManager.gNextApplicationIndex

Set mBracketOrders = New BracketOrders

Set mOrders = New Orders
Set mExecutions = New Executions

Set mProfitListeners = New Listeners
Set mChangeListeners = New Listeners

End Sub

Private Sub Class_Terminate()
Debug.Print "PositionManager terminated"
End Sub

'@================================================================================
' IChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
On Error GoTo Err

If TypeOf ev.Source Is IBracketOrder Then
    If ev.changeType = BracketOrderChangeTypes.BracketOrderCompleted Then
        Dim lBracketOrder As IBracketOrder
        Set lBracketOrder = ev.Source
        lBracketOrder.RemoveChangeListener Me
        fireChange PositionSizeChanged
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Select Case ev.Tick.TickType
Case TickTypeBid
    processBidTick ev.Tick
Case TickTypeAsk
    processAskTick ev.Tick
Case TickTypeTrade
    If ev.Tick.Price > mLastAsk Then processAskTick ev.Tick  ' because the Ask Price must have gone at least that high
    If ev.Tick.Price < mLastBid Then processBidTick ev.Tick  ' because the Bid Price must have gone at least that low
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IOrderAuthoriser Interface Members
'@================================================================================

Private Function IOrderAuthoriser_AllowOrder( _
                ByVal pIsLongOrder As Boolean, _
                ByRef pOrderSize As Long, _
                ByRef pRisk As Currency, _
                ByRef pMessage As String) As Boolean
Const ProcName As String = "IOrderAuthoriser_AllowOrder"
On Error GoTo Err

checkInitialised

If mClosingPosition Then
    IOrderAuthoriser_AllowOrder = False
    pMessage = "Closing positions"
    Exit Function
End If

If Not mMoneyManager Is Nothing Then
    IOrderAuthoriser_AllowOrder = mMoneyManager.AllowOrder(PositionSize, _
                                        PendingPositionSize, _
                                        pIsLongOrder, _
                                        pOrderSize, _
                                        pRisk + Risk, _
                                        mSessionProfit, _
                                        mMaxSessionProfit - mSessionProfit, _
                                        mProfitThisTrade, _
                                        mMaxProfitThisTrade - mProfitThisTrade, _
                                        mSession.SessionCurrentTime, _
                                        pMessage, _
                                        mEndingSession)
Else
    IOrderAuthoriser_AllowOrder = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' IOrderSubmissionListener Interface Members
'@================================================================================

Private Sub IOrderSubmissionListener_NotifyError(ByVal pOrderId As String, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IOrderSubmissionListener_NotifyError"
On Error GoTo Err

mOrders.Item(pOrderId).NotifyError pErrorCode, pErrorMsg
RaiseEvent OrderError(pOrderId, pErrorCode, pErrorMsg)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmissionListener_NotifyExecutionReport(ByVal pExecutionReport As IExecutionReport)
Const ProcName As String = "IOrderSubmissionListener_NotifyExecutionReport"
On Error GoTo Err

Dim lExec As Execution

If mExecutions.Contains(pExecutionReport.Id) Then
    ' we've already had this notification
    Set lExec = mExecutions.Item(pExecutionReport.Id)
    lExec.SyncToExecutionReport pExecutionReport
    mExecutions.NotifyChange lExec
Else
    If TypeOf pExecutionReport Is Execution Then
        Set lExec = pExecutionReport
    Else
        Set lExec = New Execution
        lExec.Id = pExecutionReport.Id
        If mSummary.IsSimulated Then lExec.IsSimulated = True
        lExec.SyncToExecutionReport pExecutionReport
    End If
    
    lExec.GroupName = mSummary.GroupName
    lExec.PositionManagerName = mSummary.Name
    If Not mBracketOrderRecoveryController Is Nothing Then
        lExec.AddRecoverableObjectListener mBracketOrderRecoveryController
        lExec.SaveRecoveryInfo
    End If
    
    mExecutions.Add lExec
    processExecution lExec
    
    mOrders.Item(lExec.OrderId).NotifyFill lExec

    fireChange PositionManagerChangeTypes.PositionSizeChanged
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmissionListener_NotifyOrderReport(ByVal pOrderReport As IOrderReport)
Const ProcName As String = "IOrderSubmissionListener_NotifyOrderReport"
On Error GoTo Err

mOrders(pOrderReport.Id).SyncToOrderReport pOrderReport

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderSubmissionListener_NotifyOrderStatusReport(ByVal pOrderStatusReport As IOrderStatusReport)
Const ProcName As String = "IOrderSubmissionListener_NotifyOrderStatusReport"
On Error GoTo Err

Dim lOrder As IOrder
Set lOrder = mOrders.Item(pOrderStatusReport.OrderId)
lOrder.Status = pOrderStatusReport.Status

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IPositionMember Interface Members
'@================================================================================

Private Property Get IPositionMember_PositionManagerName() As String
IPositionMember_PositionManagerName = mSummary.Name
End Property

'@================================================================================
' IRecoverable Interface Members
'@================================================================================

Private Sub IRecoverable_AddRecoverableObjectListener(ByVal pListener As IRecoverableObjectListener)
Const ProcName As String = "IRecoverable_AddRecoverableObjectListener"
On Error GoTo Err

AddRecoverableObjectListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IRecoverable_FinishRecovery()
Const ProcName As String = "IRecoverable_FinishRecovery"
On Error GoTo Err

FinishRecovery

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IRecoverable_LoadRecoveryData(ByVal pData As Variant)
Const ProcName As String = "IRecoverable_LoadRecoveryData"
On Error GoTo Err

Set mSummary = pData

init CreateFuture(mSummary.Contract)

mIsInitialised = True


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IRecoverable_ProgId() As String
IRecoverable_ProgId = ProjectName & "." & ModuleName
End Property

Private Property Get IRecoverable_RecoveryData() As Variant
Set IRecoverable_RecoveryData = mSummary
End Property

Private Sub IRecoverable_RemoveRecoverableObjectListener(ByVal pListener As IRecoverableObjectListener)
Const ProcName As String = "IRecoverable_RemoveRecoverableObjectListener"
On Error GoTo Err

RemoveRecoverableObjectListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

If ev.Source Is mOrderSubmitter Then
    Select Case ev.State
    Case OrderSubmitterStateNotReady
        mProviderIsReady = False
        mClosingPosition = False
        fireChange ProviderReadinessChanged
    Case OrderSubmitterStateReady
        mProviderIsReady = True
        fireChange ProviderReadinessChanged
    End Select
ElseIf ev.Source Is mSessionFuture Then
    If ev.State = FutureStates.FutureStateCompleted Then
        Set mSession = mSessionFuture.Value
        If mFinishRecoveryPending Then
            FinishRecovery
        ElseIf mRecovering Then
        Else
            checkSetupNewSession
        End If
    End If
ElseIf ev.Source Is mContractFuture Then
    If ev.State = FutureStates.FutureStateCompleted Then setContract
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mBracketOrders Event Handlers
'@================================================================================

Private Sub mBracketOrders_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mBracketOrders_CollectionChanged"
On Error GoTo Err

Dim lObj As IRecoverable
Select Case ev.changeType
Case CollItemAdded
    If Not mInitialRecoveryDataSaved Then SaveRecoveryInfo

    Dim lBracketOrder As IBracketOrder
    Set lBracketOrder = ev.AffectedItem
    lBracketOrder.AddChangeListener Me
    
    Set lObj = lBracketOrder
    If Not mBracketOrderRecoveryController Is Nothing Then lObj.AddRecoverableObjectListener mBracketOrderRecoveryController
    fireChange PositionSizeChanged
Case CollItemRemoved
    Set lObj = ev.AffectedItem
    If Not mBracketOrderRecoveryController Is Nothing Then lObj.RemoveRecoverableObjectListener mBracketOrderRecoveryController
    fireChange PositionSizeChanged
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mClosePositionsTask Event Handlers
'@================================================================================

Private Sub mClosePositionsITask_PositionClosed()
Const ProcName As String = "mClosePositionsITask_PositionClosed"
On Error GoTo Err

fireChange PositionClosed
Set mClosePositionsTask = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mOrderContexts Event Handlers
'@================================================================================

Private Sub mOrderContexts_PositionsClosed()
mClosingPosition = False
End Sub

'@================================================================================
' mOrders Event Handlers
'@================================================================================

Private Sub mOrders_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mOrders_CollectionChanged"
On Error GoTo Err

Dim lObj As IRecoverable
Select Case ev.changeType
Case CollItemAdded
    Set lObj = ev.AffectedItem
    If Not mBracketOrderRecoveryController Is Nothing Then lObj.AddRecoverableObjectListener mBracketOrderRecoveryController
Case CollItemRemoved
    Set lObj = ev.AffectedItem
    If Not mBracketOrderRecoveryController Is Nothing Then lObj.RemoveRecoverableObjectListener mBracketOrderRecoveryController
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mSession Event Handlers
'@================================================================================

Private Sub mSession_SessionEnded(ev As SessionEventData)
Const ProcName As String = "mSession_SessionEnded"
On Error GoTo Err

'If mProfitIncrement = 0 Then Exit Sub
'GLogProfitProfile FormatTimestamp(mProfitProfileBarStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & "," & _
'                    mProfitProfileBarNumber & "," & _
'                    mProfitProfileBarOpen & "," & _
'                    mProfitProfileBarHigh & "," & _
'                    mProfitProfileBarLow & "," & _
'                    mProfitProfileBarClose, _
'                    mIsSimulated, _
'                    Me
'GLogProfitProfile FormatTimestamp(mSession.SessionCurrentTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & "," & _
'                mProfitProfileBarNumber + 1 & "," & _
'                mSessionProfit & "," & _
'                mSessionProfit & "," & _
'                mSessionProfit & "," & _
'                mSessionProfit, _
'                mIsSimulated, _
'                Me

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mSession_SessionStarted(ev As SessionEventData)
Const ProcName As String = "mSession_SessionStarted"
On Error GoTo Err

If mFinishRecoveryPending Then
    FinishRecovery
ElseIf mRecovering Then
Else
    checkSetupNewSession
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get ApplicationIndex() As Long
ApplicationIndex = mApplicationIndex
End Property

Public Property Get BracketOrders() As BracketOrders
Const ProcName As String = "BracketOrders"
On Error GoTo Err

checkInitialised
Set BracketOrders = mBracketOrders

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ContractFuture() As IFuture
Const ProcName As String = "ContractFuture"
On Error GoTo Err

checkInitialised
Set ContractFuture = mContractFuture

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DataSource() As IMarketDataSource
Const ProcName As String = "DataSource"
On Error GoTo Err

checkInitialised
Set DataSource = mDataSource

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Drawdown() As Currency
Const ProcName As String = "Drawdown"
On Error GoTo Err

checkInitialised

Drawdown = mSessionDrawdown

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DrawdownLastTrade() As Currency
Const ProcName As String = "DrawdownLastTrade"
On Error GoTo Err

checkInitialised
DrawdownLastTrade = mSummary.DrawdownLastTrade

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DrawdownThisTrade() As Currency
Const ProcName As String = "DrawdownThisTrade"
On Error GoTo Err

DrawdownThisTrade = mMaxProfitThisTrade - mProfitThisTrade

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get EndingSession() As Boolean
EndingSession = mEndingSession
End Property

Public Property Get Executions() As Executions
Const ProcName As String = "Executions"
On Error GoTo Err

checkInitialised
Set Executions = mExecutions

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsFinished() As Boolean
IsFinished = mIsFinished
End Property

Public Property Get GroupName() As String
GroupName = mSummary.GroupName
End Property

Public Property Get IsRiskUnlimited() As Boolean
Const ProcName As String = "IsRiskUnlimited"
On Error GoTo Err

checkInitialised
IsRiskUnlimited = (Risk = MaxCurrency)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsSimulated() As Boolean
Const ProcName As String = "IsSimulated"
On Error GoTo Err

checkInitialised
IsSimulated = mSummary.IsSimulated

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MaxLoss() As Currency
Const ProcName As String = "MaxLoss"
On Error GoTo Err

checkInitialised
MaxLoss = mMaxSessionLoss

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MaxLossThisTrade() As Currency
Const ProcName As String = "MaxLossThisTrade"
On Error GoTo Err

MaxLossThisTrade = mMaxLossThisTrade

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MaxProfit() As Currency
Const ProcName As String = "MaxProfit"
On Error GoTo Err

checkInitialised
MaxProfit = mMaxSessionProfit

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MaxProfitThisTrade() As Currency
Const ProcName As String = "MaxProfitThisTrade"
On Error GoTo Err

MaxProfitThisTrade = mMaxProfitThisTrade

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Name() As String
Name = mSummary.Name
End Property

Public Property Get OrderContexts() As OrderContexts
Const ProcName As String = "OrderContexts"
On Error GoTo Err

checkInitialised
Set OrderContexts = mOrderContexts

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Orders() As Orders
Const ProcName As String = "Orders"
On Error GoTo Err

checkInitialised
Set Orders = mOrders

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OrderSubmitter() As IOrderSubmitter
Set OrderSubmitter = mOrderSubmitter
End Property

Public Property Get PendingPositionSize() As Integer
Const ProcName As String = "PendingPositionSize"
On Error GoTo Err

checkInitialised

Dim Size As Long

Dim OrderContext As OrderContext
For Each OrderContext In mOrderContexts
    Size = Size + OrderContext.PendingSize
Next

PendingPositionSize = Size

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PositionSize() As Integer
Const ProcName As String = "PositionSize"
On Error GoTo Err

checkInitialised
PositionSize = mQuantityOutstanding

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Profit() As Currency
Const ProcName As String = "Profit"
On Error GoTo Err

checkInitialised
Profit = mSessionProfit

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ProfitLastTrade() As Currency
Const ProcName As String = "ProfitLastTrade"
On Error GoTo Err

checkInitialised
ProfitLastTrade = mSummary.ProfitLastTrade

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ProfitThisTrade() As Currency
Const ProcName As String = "ProfitThisTrade"
On Error GoTo Err

checkInitialised
ProfitThisTrade = mProfitThisTrade

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ProviderIsReady() As Boolean
Const ProcName As String = "ProviderIsReady"
On Error GoTo Err

checkInitialised
ProviderIsReady = mProviderIsReady

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Risk() As Currency
Const ProcName As String = "Risk"
On Error GoTo Err

checkInitialised

Dim totRisk As Currency

Dim OrderContext As OrderContext
For Each OrderContext In mOrderContexts
    Dim ocRisk As Currency
    ocRisk = OrderContext.Risk
    If ocRisk = MaxCurrency Then
        totRisk = MaxCurrency
        Exit For
    End If
    totRisk = totRisk + ocRisk
Next

Risk = totRisk

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TradeStartTime() As Date
Const ProcName As String = "TradeStartTime"
On Error GoTo Err

TradeStartTime = mTradeStartTime

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TradeStartTimezoneName() As String
Const ProcName As String = "TradeStartTimezoneName"
On Error GoTo Err

TradeStartTimezoneName = mTradeStartTimezoneName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddChangeListener(ByVal pListener As IChangeListener)
Const ProcName As String = "AddChangeListener"
On Error GoTo Err

mChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddOrderSubmissionListener(ByVal pListener As IOrderSubmissionListener)
Const ProcName As String = "AddOrderSubmissionListener"
On Error GoTo Err

mOrderSubmitter.AddOrderSubmissionListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddProfitListener(ByVal pListener As IProfitListener)
Const ProcName As String = "AddProfitListener"
On Error GoTo Err

mProfitListeners.Add pListener

Dim ev As ProfitEventData
Set ev.Source = Me
ev.CurrencyCode = mCurrencyCode

Dim lProfitTypes As ProfitTypes

If mSessionProfit <> 0 Then lProfitTypes = ProfitTypes.ProfitTypeSessionProfit
If mMaxSessionProfit <> 0 Then lProfitTypes = lProfitTypes Or ProfitTypes.ProfitTypeSessionMaxProfit
If mMaxSessionLoss <> 0 Then lProfitTypes = lProfitTypes Or ProfitTypes.ProfitTypeSessionMaxLoss
If mSessionDrawdown <> 0 Then lProfitTypes = lProfitTypes Or ProfitTypes.ProfitTypeSessionDrawdown

If mProfitThisTrade <> 0 Then lProfitTypes = lProfitTypes Or ProfitTypes.ProfitTypeTradeProfit
If mMaxProfitThisTrade <> 0 Then lProfitTypes = lProfitTypes Or ProfitTypes.ProfitTypeTradeMaxProfit
If mMaxLossThisTrade <> 0 Then lProfitTypes = lProfitTypes Or ProfitTypes.ProfitTypeTradeMaxLoss
If mDrawdownThisTrade <> 0 Then lProfitTypes = lProfitTypes Or ProfitTypes.ProfitTypeTradeDrawdown
    
ev.ProfitTypes = lProfitTypes
pListener.NotifyProfit ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddRecoverableObjectListener(ByVal pListener As IRecoverableObjectListener)
Const ProcName As String = "AddRecoverableObjectListener"
On Error GoTo Err

mRecoverableObjectListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClosePositions( _
                Optional ByVal pFlags As ClosePositionsFlags = ClosePositionsCancelOrders Or ClosePositionsWaitForCancel)
Const ProcName As String = "ClosePositions"
On Error GoTo Err

checkInitialised

If mClosingPosition Then Exit Sub

If PositionSize = 0 And PendingPositionSize = 0 Then
    fireChange PositionClosed
    Exit Sub
End If

mClosingPosition = True

Set mClosePositionsTask = New ClosePositionsTask
mClosePositionsTask.ClosePositions Me, pFlags

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub CompleteRecoveryInitialisation( _
                ByVal pOrderSubmitter As IOrderSubmitter, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pBracketOrderRecoveryCtlr As BracketOrderRecoveryCtlr)
Const ProcName As String = "CompleteRecoveryInitialisation"
On Error GoTo Err

AssertArgument Not pOrderSubmitter Is Nothing, "pOrderSubmitter is nothing"

checkInitialised

Set mOrderSubmitter = pOrderSubmitter
mOrderSubmitter.AddOrderSubmissionListener Me
mOrderSubmitter.AddStateChangeListener Me

If Not pDataSource Is Nothing Then
    Set mDataSource = pDataSource
    Set mSessionFuture = mDataSource.SessionFuture
    If mSessionFuture.IsAvailable Then
        Set mSession = mSessionFuture.Value
    Else
        mSessionFuture.AddStateChangeListener Me
    End If
    mDataSource.AddGenericTickListener Me
Else
    ' this only happens if the contract has expired
    Dim lContract As IContract
    Set lContract = mContractFuture.Value
    
    Dim lSessionBuilder As SessionBuilder
    Set lSessionBuilder = CreateSessionBuilder(lContract.SessionStartTime, lContract.SessionEndTime, GetTimeZone(lContract.TimezoneName), Now)
    Set mSession = lSessionBuilder.Session
End If

Set mBracketOrderRecoveryController = pBracketOrderRecoveryCtlr
If Not mBracketOrderRecoveryController Is Nothing Then mRecoverableObjectListeners.Add mBracketOrderRecoveryController

Set mOrderContexts = mOrderManager.CreateOrderContexts(mSummary.Name, mContractFuture, mOrderSubmitter, mDataSource, mSummary.GroupName, mSummary.IsSimulated, Me, mBracketOrders, mOrders, IIf(mDataSource.IsTickReplay, mDataSource.ClockFuture, Nothing))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If IsFinished Then Exit Sub

Assert Not mRecovering, "Invalid call to Finish during recovery"

If Not needExistingBracketOrders Then DeleteRecoveryInfo

doFinish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pOrderManager As OrderManager, _
                ByVal pName As String, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pRecoveryController As BracketOrderRecoveryCtlr, _
                ByVal pGroupName As String, _
                ByVal pIsSimulated As Boolean, _
                ByVal pMoneyManager As IMoneyManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

Assert Not mIsInitialised, "PositionManager is already initialised"
Assert pName <> "", "A name must be supplied"
AssertArgument Not pOrderSubmitterFactory Is Nothing, "pOrderSubmitterFactory Is Nothing"
AssertArgument Not pDataSource Is Nothing, "pDataSource is nothing"

mSummary.Name = pName
mSummary.GroupName = pGroupName
mSummary.IsSimulated = pIsSimulated

PreInitialise pOrderManager

Set mDataSource = pDataSource
mDataSource.AddGenericTickListener Me

Set mOrderSubmitter = pOrderSubmitterFactory.CreateOrderSubmitter(mDataSource)
mOrderSubmitter.AddOrderSubmissionListener Me
mOrderSubmitter.AddStateChangeListener Me

Set mBracketOrderRecoveryController = pRecoveryController
If Not mBracketOrderRecoveryController Is Nothing Then mRecoverableObjectListeners.Add mBracketOrderRecoveryController

Set mSessionFuture = mDataSource.SessionFuture
If mSessionFuture.IsAvailable Then
    Set mSession = mSessionFuture.Value
Else
    mSessionFuture.AddStateChangeListener Me
End If
    
Set mMoneyManager = pMoneyManager

init mDataSource.ContractFuture

If Not mSession Is Nothing Then checkSetupNewSession

mIsInitialised = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub PreInitialise(ByVal pOrderManager As OrderManager)
Set mOrderManager = pOrderManager
End Sub

Public Sub RemoveChangeListener(ByVal pListener As IChangeListener)
Const ProcName As String = "RemoveChangeListener"
On Error GoTo Err

mChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveOrderSubmissionListener(ByVal pListener As IOrderSubmissionListener)
Const ProcName As String = "RemoveOrderSubmissionListener"
On Error GoTo Err

mOrderSubmitter.RemoveOrderSubmissionListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveProfitListener(ByVal pListener As IProfitListener)
Const ProcName As String = "RemoveProfitListener"
On Error GoTo Err

mProfitListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveRecoverableObjectListener(ByVal pListener As IRecoverableObjectListener)
Const ProcName As String = "RemoveRecoverableObjectListener"
On Error GoTo Err

mRecoverableObjectListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub StartRecovery()
Const ProcName As String = "StartRecovery"
On Error GoTo Err

checkInitialised

mRecovering = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcProfits(Optional ByVal Price As Double)
Const ProcName As String = "calcProfits"
On Error GoTo Err

If mTickSize = 0 Then Exit Sub
If mSummary.CumSellPrice = 0 And mSummary.CumBuyPrice = 0 And mQuantityOutstanding = 0 Then Exit Sub

Dim lProfit As Currency
lProfit = mTickValue * (mSummary.CumSellPrice + mQuantityOutstanding * Price - mSummary.CumBuyPrice) / mTickSize

If lProfit = mSessionProfit Then Exit Sub
    
Dim lProfitTypes As ProfitTypes
mSessionProfit = lProfit
lProfitTypes = ProfitTypes.ProfitTypeSessionProfit

If mSessionProfit > mMaxSessionProfit Then
    mMaxSessionProfit = mSessionProfit
    lProfitTypes = lProfitTypes Or ProfitTypeSessionMaxProfit
    If mSessionDrawdown <> 0 Then
        mSessionDrawdown = 0
        lProfitTypes = lProfitTypes Or ProfitTypeSessionDrawdown
    End If
End If
If mSessionProfit < mMaxSessionLoss Then
    mMaxSessionLoss = mSessionProfit
    lProfitTypes = lProfitTypes Or ProfitTypeSessionMaxLoss
End If
If mSessionProfit < mMaxSessionProfit Then
    mSessionDrawdown = mMaxSessionProfit - mSessionProfit
    lProfitTypes = lProfitTypes Or ProfitTypeSessionDrawdown
End If

mProfitThisTrade = mSessionProfit - mSummary.ProfitLastFlat
lProfitTypes = lProfitTypes Or ProfitTypeTradeProfit

If mProfitThisTrade > mMaxProfitThisTrade Then
    mMaxProfitThisTrade = mProfitThisTrade
    lProfitTypes = lProfitTypes Or ProfitTypeTradeMaxProfit
    If mDrawdownThisTrade <> 0 Then
        mDrawdownThisTrade = 0
        lProfitTypes = lProfitTypes Or ProfitTypeTradeDrawdown
    End If
End If
If mProfitThisTrade < mMaxLossThisTrade Then
    mMaxLossThisTrade = mProfitThisTrade
    lProfitTypes = lProfitTypes Or ProfitTypeTradeMaxLoss
End If
If mProfitThisTrade < mMaxProfitThisTrade Then
    mDrawdownThisTrade = mMaxProfitThisTrade - mProfitThisTrade
    lProfitTypes = lProfitTypes Or ProfitTypeTradeDrawdown
End If

fireProfit lProfitTypes

If lProfitTypes And ProfitTypes.ProfitTypeSessionProfit Then
    gLogProfit mSessionProfit, mSummary.IsSimulated, Me
    gLogDrawDown mSessionDrawdown, mSummary.IsSimulated, Me
End If
If lProfitTypes And ProfitTypeSessionMaxProfit Then
    gLogMaxProfit mMaxSessionProfit, mSummary.IsSimulated, Me
End If
If lProfitTypes And ProfitTypeSessionMaxLoss Then
    gLogMaxLoss mMaxSessionLoss, mSummary.IsSimulated, Me
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub checkCloseall()
Const ProcName As String = "checkCloseall"
On Error GoTo Err

'If mIsSimulated Then Exit Sub

If mMoneyManager Is Nothing Then Exit Sub

Dim msg As String
Dim lEndSession As Boolean
If mMoneyManager.CloseAll(mSessionProfit, _
                            mMaxSessionProfit - mSessionProfit, _
                            mProfitThisTrade, _
                            mMaxProfitThisTrade - mProfitThisTrade, _
                            mSession.SessionCurrentTime, _
                            msg, _
                            lEndSession) _
Then
    If Not mClosingPosition Then
        gLogMoneyManagement "Close all positions: " & msg, mSummary.IsSimulated, Me
        ClosePositions
    End If
Else
    If lEndSession And Not mClosingPosition Then
        If Not mEndingSession Then
            mEndingSession = True
            gLogMoneyManagement "Ending Session: " & msg, mSummary.IsSimulated, Me
        End If
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub checkInitialised()
Assert Not mIsFinished, "PositionManager has been finished"
Assert mIsInitialised, "PositionManager has not been initialised"
End Sub

Private Sub checkSetupNewSession()
Const ProcName As String = "checkSetupNewSession"
On Error GoTo Err

If Int(mSession.SessionCurrentTime) = 0# Then Exit Sub

If mSummary.TradingSessionStartTime = 0# Then mSummary.TradingSessionStartTime = mSession.CurrentSessionStartTime
    
If Not needExistingBracketOrders Then setupNewSession

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub DeleteRecoveryInfo()
Const ProcName As String = "deleteRecoveryInfo"
On Error GoTo Err

If mBracketOrderRecoveryController Is Nothing Then Exit Sub

mInitialRecoveryDataSaved = False
gLog "Deleting order recovery info for Position Manager", ProcName, ModuleName, mSummary.Name
mBracketOrderRecoveryController.DeleteRecoveryInfo mSummary.Name

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doFinish()
Const ProcName As String = "doFinish"
On Error GoTo Err

mDataSource.RemoveGenericTickListener Me
Set mDataSource = Nothing
Set mSession = Nothing

mChangeListeners.Clear
mProfitListeners.Clear
mRecoverableObjectListeners.Clear

mOrderContexts.Finish
Set mOrderContexts = Nothing
Set mMoneyManager = Nothing

mOrderSubmitter.RemoveOrderSubmissionListener Me
mOrderSubmitter.RemoveStateChangeListener Me
Set mOrderSubmitter = Nothing

Set mBracketOrders = Nothing
Set mOrders = Nothing
Set mExecutions = Nothing

mIsFinished = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub FinishRecovery()
Const ProcName As String = "FinishRecovery"
On Error GoTo Err

checkInitialised

If mSession Is Nothing Then
    mFinishRecoveryPending = True
    Exit Sub
End If

If Int(mSession.SessionCurrentTime) = 0# Then
    mFinishRecoveryPending = True
    Exit Sub
End If

mFinishRecoveryPending = False
mRecovering = False

If needExistingBracketOrders Then
    setupSessionBracketOrders
    SaveRecoveryInfo
Else
    DeleteRecoveryInfo
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireChange( _
                ByVal changeType As PositionManagerChangeTypes)
Const ProcName As String = "fireChange"
On Error GoTo Err

Dim ev As ChangeEventData
Set ev.Source = Me
ev.changeType = changeType

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mChangeListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Change ev
    Next
End If

RaiseEvent Change(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireProfit( _
                ByVal pProfitTypes As ProfitTypes)
Const ProcName As String = "fireProfit"
On Error GoTo Err

Dim ev As ProfitEventData
Set ev.Source = Me
ev.ProfitTypes = pProfitTypes
ev.CurrencyCode = mCurrencyCode

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mProfitListeners.Valid Then
    sInit = True
    sSomeListeners = mProfitListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IProfitListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyProfit ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireRecoveryDataChanged()
Const ProcName As String = "fireRecoveryDataChanged"
On Error GoTo Err

Dim ev As RecoveryDataChangedEventData
Set ev.Source = Me

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mRecoverableObjectListeners.Valid Then
    sInit = True
    sSomeListeners = mRecoverableObjectListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IRecoverableObjectListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyRecoveryDataChanged ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function haveActiveBracketOrders() As Boolean
Const ProcName As String = "haveActiveBracketOrders"
On Error GoTo Err

Dim oc As OrderContext
For Each oc In mOrderContexts
    If oc.ActiveBracketOrders.Count <> 0 Then haveActiveBracketOrders = True: Exit Function
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function haveClosedBracketOrders() As Boolean
Const ProcName As String = "haveClosedBracketOrders"
On Error GoTo Err

Dim oc As OrderContext
For Each oc In mOrderContexts
    If oc.ClosedBracketOrders.Count <> 0 Then haveClosedBracketOrders = True: Exit Function
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub init( _
                ByVal pContractFuture As IFuture)
Const ProcName As String = "init"
On Error GoTo Err

Set mContractFuture = pContractFuture

If mOrderContexts Is Nothing And Not mOrderSubmitter Is Nothing Then
    Set mOrderContexts = mOrderManager.GetOrderContexts(mSummary.Name, mSummary.IsSimulated)
    If mOrderContexts Is Nothing Then _
        Set mOrderContexts = mOrderManager.CreateOrderContexts(mSummary.Name, mContractFuture, mOrderSubmitter, mDataSource, mSummary.GroupName, mSummary.IsSimulated, Me, mBracketOrders, mOrders, IIf(mDataSource.IsTickReplay, mDataSource.ClockFuture, Nothing))
End If

If mContractFuture.IsAvailable Then
    setContract
Else
    mContractFuture.AddStateChangeListener Me
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function isFlat() As Boolean
Const ProcName As String = "isFlat"
On Error GoTo Err

isFlat = (PositionSize = 0 And PendingPositionSize = 0)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function needExistingBracketOrders() As Boolean
Const ProcName As String = "needExistingBracketOrders"
On Error GoTo Err

If mQuantityOutstanding <> 0 Or haveActiveBracketOrders Then
    needExistingBracketOrders = True
ElseIf stillInInitialSession And haveClosedBracketOrders Then
    needExistingBracketOrders = True
Else
    needExistingBracketOrders = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processAskTick(ByRef pTick As GenericTick)
Const ProcName As String = "processAskTick"
On Error GoTo Err

If pTick.Price = 0 Then Exit Sub ' can occur at start of Session
If pTick.Price = mLastAsk Then Exit Sub

mLastAsk = pTick.Price

If mQuantityOutstanding < 0 Then
    calcProfits pTick.Price
    checkCloseall
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processBidTick(ByRef pTick As GenericTick)
Const ProcName As String = "processBidTick"
On Error GoTo Err

If pTick.Price = 0 Then Exit Sub  ' can occur at start of Session
If pTick.Price = mLastBid Then Exit Sub

mLastBid = pTick.Price

If mQuantityOutstanding > 0 Then
    calcProfits pTick.Price
    checkCloseall
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processExecution(ByVal pExec As IExecutionReport)
Const ProcName As String = "processExecution"
On Error GoTo Err

If mQuantityOutstanding = 0 Then
    mTradeStartTime = pExec.FillTime
    mTradeStartTimezoneName = pExec.TimezoneName
End If

mQuantityOutstanding = mQuantityOutstanding + gGetSignedQuantity(pExec)

If gGetSignedQuantity(pExec) > 0 Then
    mSummary.CumBuyPrice = mSummary.CumBuyPrice + (pExec.Price * gGetSignedQuantity(pExec))
Else
    mSummary.CumSellPrice = mSummary.CumSellPrice + (pExec.Price * pExec.Quantity)
End If

If mQuantityOutstanding <> 0 Then
    'we need to notify the position change before the Profit
    gLogPosition mQuantityOutstanding, mSummary.IsSimulated, Me

    If mQuantityOutstanding < 0 Then
        If mLastAsk <> 0 Then
            calcProfits mLastAsk
        End If
    Else        ' mQuantityOutstanding > 0
        If mLastBid <> 0 Then
            calcProfits mLastBid
        End If
    End If
Else
    'we need to notify the position change after the final Profit
    calcProfits
    
    gLogPosition mQuantityOutstanding, mSummary.IsSimulated, Me
    gLogTradeProfile FormatTimestamp(mTradeStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & "," & _
            FormatTimestamp(pExec.FillTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & "," & _
            mTradeStartTimezoneName & "," & _
            mProfitThisTrade & "," & _
            mMaxProfitThisTrade & "," & _
            mMaxLossThisTrade, _
            mSummary.IsSimulated, _
            Me
    mSummary.ProfitLastTrade = mProfitThisTrade
    mSummary.DrawdownLastTrade = mMaxProfitThisTrade - mProfitThisTrade
    mProfitThisTrade = 0
    mMaxProfitThisTrade = 0
    mMaxLossThisTrade = 0
    mSummary.ProfitLastFlat = mSessionProfit
    mDrawdownThisTrade = 0
    
    If Not mMoneyManager Is Nothing Then mMoneyManager.NotifyPositionFlat
    
End If

If Not mRecovering Then SaveRecoveryInfo

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub SaveRecoveryInfo()
Const ProcName As String = "SaveRecoveryInfo"
On Error GoTo Err

If Not mContractFuture.IsAvailable Then Exit Sub
If Not mSessionFuture.IsAvailable Then Exit Sub
If Not mSummary.IsDirty Then Exit Sub
If mRecoverableObjectListeners Is Nothing Or mRecoverableObjectListeners.Count = 0 Then Exit Sub

mInitialRecoveryDataSaved = True
fireRecoveryDataChanged
mSummary.setClean

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setContract()
Const ProcName As String = "setContract"
On Error GoTo Err

Dim lContract As IContract
Set lContract = mContractFuture.Value

mSummary.Contract = lContract
mTickValue = lContract.TickValue
mTickSize = lContract.TickSize
mCurrencyCode = lContract.Specifier.CurrencyCode

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupNewSession()
Const ProcName As String = "setupNewSession"
On Error GoTo Err

'If mIsSimulated Then Exit Sub

mBracketOrders.Clear
mOrders.Clear
mExecutions.Clear

mSummary.TradingSessionStartTime = mSession.CurrentSessionStartTime

mSummary.CumBuyPrice = 0
mSummary.CumSellPrice = 0
mQuantityOutstanding = 0

mSummary.ProfitLastTrade = mProfitThisTrade
mSummary.DrawdownLastTrade = mMaxProfitThisTrade - mProfitThisTrade
mProfitThisTrade = 0
mMaxProfitThisTrade = 0
mMaxLossThisTrade = 0
mSummary.ProfitLastFlat = mSessionProfit

mSessionProfit = 0
mMaxSessionProfit = 0
mMaxSessionLoss = 0
mSessionDrawdown = 0

mDrawdownThisTrade = 0
mTradeStartTime = mSummary.TradingSessionStartTime

If mInitialRecoveryDataSaved Then SaveRecoveryInfo

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupSessionBracketOrders()
Const ProcName As String = "setupSessionBracketOrders"
On Error GoTo Err

Dim lOrderContext As OrderContext
For Each lOrderContext In mOrderContexts
    Dim lBracketOrder As IBracketOrder
    For Each lBracketOrder In lOrderContext.ActiveBracketOrders
        mBracketOrders.Add lBracketOrder
    Next

    For Each lBracketOrder In lOrderContext.ClosedBracketOrders
        mBracketOrders.Add lBracketOrder
    Next
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function stillInInitialSession() As Boolean
Const ProcName As String = "stillInInitialSession"
On Error GoTo Err

If mSession.CurrentSessionStartTime = mSummary.TradingSessionStartTime Then
    stillInInitialSession = True
Else
    gLog "Session start times: current=" & FormatTimestamp(mSession.CurrentSessionStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & _
            "; initial=" & FormatTimestamp(mSummary.TradingSessionStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs), _
        ProcName, ModuleName, mSummary.Name, LogLevelDetail
    stillInInitialSession = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

