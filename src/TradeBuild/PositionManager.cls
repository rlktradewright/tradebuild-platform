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

Implements IOrderSubmissionSrvcConsumer
Implements QuoteListener

'@================================================================================
' Events
'@================================================================================

Event PositionChanged()

Event ProviderReady()

Event ProviderNotReady()

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "PositionManager"

'@================================================================================
' Enums
'@================================================================================

Public Enum ReasonTypes
    ReasonNone = 0
    ReasonEntry
    ReasonTarget
    ReasonStop
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mKey As String
Private mIndexApplication As Long

Private mProviderIsReady As Boolean

Private WithEvents mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1
Private mWorkspaceRef As WeakReference

Private mOrders As Orders
Private mExecutions As Executions

Private mIsSimulated As Boolean

Private mOrderContexts As OrderContexts
Private mMoneyManager As MoneyManager
Private mStudies As Studies

Private mBalancingOrderContext As OrderContext
Private WithEvents mBalancingOrderPlex As orderPlex
Attribute mBalancingOrderPlex.VB_VarHelpID = -1
Private mSettlingHandle As Long

Private mClosingPosition As Boolean
    ' indicates that all OrderContexts are being closed, either as a result of
    ' a call to the closePosition functions, or because money management
    ' has told us to do so

Private mEndSession As Boolean
    ' indicates that money management has told us to end the Session. Note
    ' that this can happen whilst there are still some OrderContexts in progress,
    ' for example if the daily Profit has been met. In this case, the Session
    ' is ended when all OrderContexts have completed.

Private mProfitIncrementTicks As Long
Private mProfitIncrement As Currency
    ' specifies the range of the bars in the Profit profile
Private mProfitProfileBarStartTime As Date
Private mProfitProfileBarOpen As Currency
    ' the Profit at the start of the current Profit profile bar
Private mProfitProfileBarHigh As Currency
Private mProfitProfileBarLow As Currency
Private mProfitProfileBarClose As Currency
Private mProfitProfileBarNumber As Long

Private mCumBuyPrice As Double
Private mCumSellPrice As Double
Private mQuantityOutstanding As Long

Private mProfit As Currency
Private mMaxProfit As Currency
Private mMaxLoss As Currency
Private mDrawdown As Currency

Private mProfitLastFlat As Currency
Private mProfitThisTrade As Currency
Private mMaxProfitThisTrade As Currency
Private mMaxLossThisTrade As Currency
Private mDrawdownThisTrade As Currency
Private mTradeStartTime As Date
Private mTradeStartTimezone As String

Private mProfitLastTrade As Currency
Private mDrawdownLastTrade As Currency

Private mTickValue As Currency
Private mTickSize As Double
Private mcurrencyCode As String

Private mLastBid As Double
Private mLastAsk As Double

Private mProfitListeners As Collection
Private mChangeListeners As Collection

Private mOrderSubmitter As TradeBuildSP.IOrderSubmitter


'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mIndexApplication = GPositionManager.gNextApplicationIndex

Set mOrders = New Orders
Set mExecutions = New Executions

Set mOrderContexts = New OrderContexts
mOrderContexts.Initialise Me
Set mBalancingOrderContext = mOrderContexts.Add("$balancing")

Set mStudies = New Studies

Set mProfitListeners = New Collection
Set mChangeListeners = New Collection

End Sub

Private Sub Class_Terminate()
Debug.Print "PositionManager terminated"
End Sub

'@================================================================================
' TradeBuildSP.IOrderSubmissionSrvcConsumer Interface Members
'@================================================================================

Private Sub IOrderSubmissionSrvcConsumer_cancelTickData( _
                ByVal consumer As TradeBuildSP.IStreamingDataConsumer)
Const ProcName As String = "IOrderSubmissionSrvcConsumer_cancelTickData"
Dim failpoint As Long
On Error GoTo Err

mTicker.RemoveDataConsumer consumer

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderSubmissionSrvcConsumer_ExecutionReport(ByVal pExecutionReport As TradeBuildSP.IExecutionReport)
Dim lExec As Execution

Const ProcName As String = "IOrderSubmissionSrvcConsumer_ExecutionReport"
Dim failpoint As Long
On Error GoTo Err

On Error Resume Next
Set lExec = mExecutions.Item(pExecutionReport.ExecId)
On Error GoTo Err

If lExec Is Nothing Then
    If TypeOf pExecutionReport Is Execution Then
        Set lExec = pExecutionReport
        mExecutions.Add lExec
    Else
        Set lExec = New Execution
        lExec.ExecId = pExecutionReport.ExecId
        If mIsSimulated Then lExec.IsSimulated = True
        lExec.SyncToExecutionReport pExecutionReport, mKey
        mExecutions.Add lExec
        mOrders.Item(lExec.OrderTradeBuildId).NotifyFill lExec, mKey
    End If
    
    processExecution pExecutionReport
Else
    ' we've already had this notification
    lExec.SyncToExecutionReport pExecutionReport, mKey
    mExecutions.Change lExec
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderSubmissionSrvcConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)
Const ProcName As String = "IOrderSubmissionSrvcConsumer_NotifyEvent"
Dim failpoint As Long
On Error GoTo Err

gTB.ServiceProviders.CommonServiceConsumerNotifyEvent eventCode, eventMessage, spHandle

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderSubmissionSrvcConsumer_notReady()
Const ProcName As String = "IOrderSubmissionSrvcConsumer_notReady"
Dim failpoint As Long
On Error GoTo Err

mProviderIsReady = False
If Not mOrderContexts Is Nothing Then mOrderContexts.ProviderReady = False
mClosingPosition = False
fireChange ProviderReadinessChanged
RaiseEvent ProviderNotReady

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderSubmissionSrvcConsumer_OrderReport(ByVal OrderReport As TradeBuildSP.IOrderReport)
Const ProcName As String = "IOrderSubmissionSrvcConsumer_OrderReport"
Dim failpoint As Long
On Error GoTo Err

mOrders(OrderReport.tradeBuildId).SyncToOrderReport OrderReport, mKey

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderSubmissionSrvcConsumer_OrderStatusReport(ByVal pOrderStatusReport As TradeBuildSP.IOrderStatusReport)
Dim lOrder As Order
Const ProcName As String = "IOrderSubmissionSrvcConsumer_OrderStatusReport"
Dim failpoint As String
On Error GoTo Err

Set lOrder = mOrders.Item(pOrderStatusReport.orderId)
lOrder.UpdateStatus pOrderStatusReport.Status, mKey

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderSubmissionSrvcConsumer_preFill( _
                ByVal orderId As String, _
                fillPrice As Double, _
                fillSize As Long)

End Sub

Private Sub IOrderSubmissionSrvcConsumer_Ready()
Const ProcName As String = "IOrderSubmissionSrvcConsumer_Ready"
Dim failpoint As Long
On Error GoTo Err

mProviderIsReady = True
mOrderContexts.ProviderReady = True
fireChange ProviderReadinessChanged
RaiseEvent ProviderReady

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderSubmissionSrvcConsumer_rejectOrder( _
                ByVal orderId As String, _
                ByVal reason As String)
Dim ev As NotificationEvent
Dim lOrder As Order
Const ProcName As String = "IOrderSubmissionSrvcConsumer_rejectOrder"
Dim failpoint As Long
On Error GoTo Err

Set lOrder = mOrders.Item(orderId)
lOrder.UpdateStatus OrderStatuses.OrderStatusRejected, mKey

ev.eventCode = ApiNotifyCodes.ApiNotifyOrderRejected
ev.eventMessage = reason
Set ev.source = Me
gTB.notify ev

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderSubmissionSrvcConsumer_requestTickData( _
                ByVal consumer As TradeBuildSP.IStreamingDataConsumer, _
                ByVal includeMarketDepth As Boolean)
Const ProcName As String = "IOrderSubmissionSrvcConsumer_requestTickData"
Dim failpoint As Long
On Error GoTo Err

mTicker.AddDataConsumer consumer, includeMarketDepth

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' QuoteListener Interface Members
'@================================================================================

Private Sub QuoteListener_Ask(ev As QuoteEvent)

Const ProcName As String = "QuoteListener_Ask"
Dim failpoint As Long
On Error GoTo Err

If ev.price = 0 Then Exit Sub ' can occur at start of Session

If ev.price = mLastAsk Then Exit Sub

mLastAsk = ev.price

If mQuantityOutstanding < 0 Then
    calcTradeProfit ev.price
    checkCloseall
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub QuoteListener_Bid(ev As QuoteEvent)

Const ProcName As String = "QuoteListener_Bid"
Dim failpoint As Long
On Error GoTo Err

If ev.price = 0 Then Exit Sub  ' can occur at start of Session

If ev.price = mLastBid Then Exit Sub

mLastBid = ev.price

If mQuantityOutstanding > 0 Then
    calcTradeProfit ev.price
    checkCloseall
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub QuoteListener_High(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_Low(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_OpenInterest(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_PreviousClose(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_sessionOpen(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_Trade(ev As QuoteEvent)

Const ProcName As String = "QuoteListener_Trade"
Dim failpoint As Long
On Error GoTo Err

If ev.price = 0 Then Exit Sub  ' can occur at start of Session

If mQuantityOutstanding = 0 Then Exit Sub
If ev.price >= mLastAsk Then
    If mQuantityOutstanding < 0 Then
        mLastAsk = ev.price ' because the Ask price must have gone at least that High
    Else
        Exit Sub
    End If
End If
If ev.price <= mLastBid Then
    If mQuantityOutstanding > 0 Then
        mLastBid = ev.price
    Else
        Exit Sub
    End If
End If
    
calcTradeProfit ev.price
checkCloseall

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub QuoteListener_Volume(ev As QuoteEvent)

End Sub

'@================================================================================
' mBalancingOrderPlex Event Handlers
'@================================================================================

Private Sub mBalancingOrderPlex_Completed()
' now need to adjust the other OrderPlexes to take account of the
' settling order

Const ProcName As String = "mBalancingOrderPlex_Completed"
Dim failpoint As String
On Error GoTo Err

Dim oc As OrderContext

For Each oc In mOrderContexts
    oc.NotifyPositionCloseout mBalancingOrderPlex.EntryPrice, _
                        mBalancingOrderPlex.StartTime
Next

Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mTicker Event Handlers
'@================================================================================

Private Sub mTicker_StateChange(ev As TWUtilities30.StateChangeEvent)
Dim lContract As Contract

Const ProcName As String = "mTicker_StateChange"
Dim failpoint As String
On Error GoTo Err

If ev.State = TickerStates.TickerStateReady Then
    Set lContract = mTicker.Contract
    
    mTickValue = lContract.TickValue
    mTickSize = lContract.TickSize
    mProfitIncrement = mProfitIncrementTicks * mTickValue
    mcurrencyCode = lContract.specifier.currencyCode
ElseIf ev.State = TickerStates.TickerStateStopped Then
    Dim oc As OrderContext
    
    For Each oc In mOrderContexts
        oc.NotifyEndOfTradingSession mTicker.Timestamp
    Next
End If

Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mTicker_TradingSessionEnded()
Const ProcName As String = "mTicker_TradingSessionEnded"
Dim failpoint As Long
On Error GoTo Err

If mProfitIncrement = 0 Then Exit Sub
GLogProfitProfile FormatTimestamp(mProfitProfileBarStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & "," & _
                    mProfitProfileBarNumber & "," & _
                    mProfitProfileBarOpen & "," & _
                    mProfitProfileBarHigh & "," & _
                    mProfitProfileBarLow & "," & _
                    mProfitProfileBarClose, _
                    mIsSimulated, _
                    Me
GLogProfitProfile FormatTimestamp(mTicker.Timestamp, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & "," & _
                mProfitProfileBarNumber + 1 & "," & _
                mProfit & "," & _
                mProfit & "," & _
                mProfit & "," & _
                mProfit, _
                mIsSimulated, _
                Me

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub mTicker_TradingSessionStarted()
Const ProcName As String = "mTicker_TradingSessionStarted"
Dim failpoint As Long
On Error GoTo Err

mProfitProfileBarStartTime = mTicker.Timestamp

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Drawdown() As Currency
Const ProcName As String = "Drawdown"
Dim failpoint As Long
On Error GoTo Err

Drawdown = mMaxProfit - mProfit

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DrawdownlastTrade() As Currency
DrawdownlastTrade = mDrawdownLastTrade
End Property

Public Property Get DrawdownThisTrade() As Currency
Const ProcName As String = "DrawdownThisTrade"
Dim failpoint As Long
On Error GoTo Err

DrawdownThisTrade = mMaxProfitThisTrade - mProfitThisTrade

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Executions() As Executions
Set Executions = mExecutions
End Property

Public Property Get IndexApplication() As Long
IndexApplication = mIndexApplication
End Property

Public Property Get IsRiskUnlimited() As Boolean
IsRiskUnlimited = (Risk = MaxCurrency)
End Property

Friend Property Let IsSimulated(ByVal value As Boolean)
Const ProcName As String = "IsSimulated"
Dim failpoint As Long
On Error GoTo Err

mIsSimulated = value
mOrderContexts.IsSimulated = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Friend Property Get Key() As String
Key = mKey
End Property

Friend Property Let MoneyManager( _
                ByVal value As MoneyManager)
Set mMoneyManager = value
End Property

Public Property Get OrderContexts() As OrderContexts
Set OrderContexts = mOrderContexts
End Property

Public Property Get Orders() As Orders
Set Orders = mOrders
End Property

Friend Property Let orderSubmitter( _
                ByVal value As TradeBuildSP.IOrderSubmitter)
Const ProcName As String = "OrderSubmitter"
Dim failpoint As Long
On Error GoTo Err

Set mOrderSubmitter = value
mOrderContexts.orderSubmitter = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PendingPositionSize() As Integer
Dim OrderContext As OrderContext
Dim size As Long

Const ProcName As String = "PendingPositionSize"
Dim failpoint As Long
On Error GoTo Err

For Each OrderContext In mOrderContexts
    size = size + OrderContext.PendingSize
Next

PendingPositionSize = size

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PositionSize() As Integer
PositionSize = mQuantityOutstanding
End Property

Public Property Get Profit() As Currency
Profit = mProfit
End Property

Public Property Let ProfitProfileIncrementTicks(ByVal value As Long)
Const ProcName As String = "ProfitProfileIncrementTicks"
Dim failpoint As Long
On Error GoTo Err

mProfitIncrementTicks = value
If mTickValue <> 0 Then mProfitIncrement = mProfitIncrementTicks * mTickValue

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ProfitLastTrade() As Currency
ProfitLastTrade = mProfitLastTrade
End Property

Public Property Get ProfitThisTrade() As Currency
ProfitThisTrade = mProfitThisTrade
End Property

Public Property Get ProviderIsReady() As Boolean
ProviderIsReady = mProviderIsReady
End Property

Public Property Get Risk() As Currency
Dim OrderContext As OrderContext
Dim totRisk As Currency
Dim ocRisk As Currency

Const ProcName As String = "Risk"
Dim failpoint As Long
On Error GoTo Err

For Each OrderContext In mOrderContexts
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let Ticker(ByVal value As Ticker)

Const ProcName As String = "Ticker"
Dim failpoint As Long
On Error GoTo Err

Set mTicker = value

mOrderContexts.Ticker = value

mTicker.AddQuoteListener Me

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Property

Public Property Get Ticker() As Ticker
Set Ticker = mTicker
End Property

Public Property Get Workspace() As Workspace
Set Workspace = mWorkspaceRef.Target
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddChangeListener(ByVal value As ChangeListener)
mChangeListeners.Add value
End Sub

Public Sub AddProfitListener(ByVal value As ProfitListener)
mProfitListeners.Add value
End Sub

Public Sub AddStudy( _
                ByVal study As study, _
                ByRef valueNames() As String)
Const ProcName As String = "AddStudy"
Dim failpoint As Long
On Error GoTo Err

mStudies.Add study, valueNames

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Function AllowOrder(ByVal orderLong As Boolean, _
                            ByRef orderSize As Long, _
                            ByVal newRisk As Currency, _
                            ByRef Message As String) As Boolean

Const ProcName As String = "AllowOrder"
Dim failpoint As Long
On Error GoTo Err

If mClosingPosition Then
    AllowOrder = False
    Message = "Closing positions"
    Exit Function
End If

If Not mMoneyManager Is Nothing Then
    AllowOrder = mMoneyManager.AllowOrder(PositionSize, _
                                        PendingPositionSize, _
                                        orderLong, _
                                        orderSize, _
                                        newRisk + Risk, _
                                        mProfit, _
                                        mMaxProfit - mProfit, _
                                        mProfitThisTrade, _
                                        mMaxProfitThisTrade - mProfitThisTrade, _
                                        mTicker.Timestamp, _
                                        Message, _
                                        mEndSession)
Else
    AllowOrder = True
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub ClosePositions( _
                Optional ByVal flags As ClosePositionFlags = ClosePositionCancelOrders Or ClosePositionWaitForCancel)
Dim lOrderContext As OrderContext
Dim cancellingOrders As Boolean

Const ProcName As String = "ClosePositions"
Dim failpoint As Long
On Error GoTo Err

If mClosingPosition Then Exit Sub

If PositionSize = 0 And PendingPositionSize = 0 Then
    fireChange PositionClosed
    Exit Sub
End If

mClosingPosition = True
If (flags And ClosePositionCancelOrders) = ClosePositionCancelOrders Then
    ' first need to tell each OrderContext to cancel any outstanding stuff
    For Each lOrderContext In mOrderContexts
        If lOrderContext.Cancel And (flags And ClosePositionWaitForCancel) Then
            cancellingOrders = True
        End If
    Next
End If

mOrderContexts.Enabled = False  ' prevent any further orders being placed
                                ' while the position is being closed

If Not cancellingOrders And PositionSize <> 0 Then placeBalancingOrder

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Friend Sub Finish()

Const ProcName As String = "Finish"
Dim failpoint As Long
On Error GoTo Err

Set mTicker = Nothing

mOrderContexts.Finish

Set mOrderContexts = Nothing
Set mMoneyManager = Nothing
Set mBalancingOrderContext = Nothing
Set mBalancingOrderPlex = Nothing

If Not mOrderSubmitter Is Nothing Then mOrderSubmitter.Finish
Set mOrderSubmitter = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Initialise( _
                ByVal pKey As String, _
                ByVal pWorkspace As Workspace)
mKey = pKey
Set mWorkspaceRef = CreateWeakReference(pWorkspace)
End Sub

Friend Sub NotifyOrderContextInactive(ByVal pOrderContext As OrderContext)

Dim OrderContext As OrderContext
Dim size As Long


Const ProcName As String = "NotifyOrderContextInactive"
Dim failpoint As Long
On Error GoTo Err

If Not mClosingPosition Then
    For Each OrderContext In mOrderContexts
        If Not OrderContext.Complete Then Exit Sub
    Next
    ' all OrderContexts are complete
    checkCloseall
    If mEndSession Then
        GLogMoneyManagement "Ending Session", mIsSimulated, Me
    End If
ElseIf pOrderContext Is mBalancingOrderContext Then
    'should now be net flat
    mOrderContexts.Enabled = True
    
    If mEndSession Then
        GLogMoneyManagement "Ending Session", mIsSimulated, Me
    End If
    
    mClosingPosition = False
    fireChange PositionClosed

Else
    
    For Each OrderContext In mOrderContexts
        If Not OrderContext.Complete Then Exit Sub
        size = size + OrderContext.size
    Next
    
    ' If we get here, all OrderContexts must be complete.
    
    ' Now, if there is a net position, need to close it out
    
    If size = 0 Then
        If mEndSession Then
            GLogMoneyManagement "Ending Session", mIsSimulated, Me
        End If
        mClosingPosition = False
        fireChange PositionClosed
    Else
        placeBalancingOrder
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveChangeListener(ByVal value As ChangeListener)
Dim i As Long
Const ProcName As String = "RemoveChangeListener"
Dim failpoint As Long
On Error GoTo Err

For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveProfitListener(ByVal value As ProfitListener)
Dim i As Long
Const ProcName As String = "RemoveProfitListener"
Dim failpoint As Long
On Error GoTo Err

For i = mProfitListeners.Count To 1 Step -1
    If mProfitListeners.Item(i) Is value Then mProfitListeners.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcTradeProfit(Optional ByVal price As Double)

Dim Profit As Currency
Dim ev As StudyValueEvent

Const ProcName As String = "calcTradeProfit"
Dim failpoint As Long
On Error GoTo Err

If mTickSize = 0 Then Exit Sub

Profit = mTickValue * (mCumSellPrice + mQuantityOutstanding * price - mCumBuyPrice) / mTickSize

If Profit <> mProfit Then
    mProfit = Profit
    
    fireProfit mProfit, ProfitTypes.ProfitTypeSessionProfit
    
    Set ev.source = Me
    ev.sVal.Timestamp = mTicker.Timestamp
    ev.sVal.value = mProfit
    ev.valueName = "Profit"
    mStudies.notify ev
    
    GLogProfit mProfit, mIsSimulated, Me
    GLogDrawdown mMaxProfit - mProfit, mIsSimulated, Me
End If

If mProfit > mMaxProfit Then
    mMaxProfit = mProfit
    fireProfit mMaxProfit, ProfitTypes.ProfitTypeSessionMaxProfit
    GLogMaxProfit mMaxProfit, mIsSimulated, Me
End If
If mProfit < mMaxLoss Then
    mMaxLoss = mProfit
    fireProfit mMaxLoss, ProfitTypes.ProfitTypeSessionMaxProfit
    GLogMaxLoss mMaxLoss, mIsSimulated, Me
End If
If mProfit < mMaxProfit Then
    mDrawdown = mMaxProfit - mProfit
    fireProfit mDrawdown, ProfitTypes.ProfitTypeSessionDrawdown
ElseIf mDrawdown <> 0 Then
    mDrawdown = 0
    fireProfit mDrawdown, ProfitTypes.ProfitTypeSessionDrawdown
End If

mProfitThisTrade = mProfit - mProfitLastFlat
fireProfit mProfitThisTrade, ProfitTypes.ProfitTypeTradeProfit
If mProfitThisTrade > mMaxProfitThisTrade Then
    mMaxProfitThisTrade = mProfitThisTrade
    fireProfit mMaxProfitThisTrade, ProfitTypes.ProfitTypeTradeMaxProfit
End If
If mProfitThisTrade < mMaxLossThisTrade Then
    mMaxLossThisTrade = mProfitThisTrade
    fireProfit mMaxLossThisTrade, ProfitTypes.ProfitTypeTradeMaxLoss
End If
If (mMaxProfitThisTrade - mProfitThisTrade) < mDrawdownThisTrade Then
    mDrawdownThisTrade = mMaxProfitThisTrade - mProfitThisTrade
    fireProfit mDrawdownThisTrade, ProfitTypes.ProfitTypeTradeDrawdown
End If

If mProfitIncrement = 0 Then Exit Sub

If mProfit - mProfitProfileBarLow >= mProfitIncrement Or _
    mProfitProfileBarHigh - mProfit >= mProfitIncrement _
Then
    ' start a new Profit profile bar
    GLogProfitProfile mProfitProfileBarStartTime & "," & _
                        mProfitProfileBarNumber & "," & _
                        mProfitProfileBarOpen & "," & _
                        mProfitProfileBarHigh & "," & _
                        mProfitProfileBarLow & "," & _
                        mProfitProfileBarClose, _
                        mIsSimulated, _
                        Me
    mProfitProfileBarStartTime = mTicker.Timestamp
    mProfitProfileBarNumber = mProfitProfileBarNumber + 1
    mProfitProfileBarOpen = mProfit
    mProfitProfileBarHigh = mProfit
    mProfitProfileBarLow = mProfit
    mProfitProfileBarClose = mProfit
Else
    If mProfit > mProfitProfileBarHigh Then mProfitProfileBarHigh = mProfit
    If mProfit < mProfitProfileBarLow Then mProfitProfileBarLow = mProfit
    mProfitProfileBarClose = mProfit
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub checkCloseall()

Dim msg As String
Dim EndSession As Boolean

Const ProcName As String = "checkCloseall"
Dim failpoint As Long
On Error GoTo Err

If mIsSimulated Then Exit Sub

If Not mMoneyManager Is Nothing Then
    If mMoneyManager.CloseAll(mProfit, _
                                mMaxProfit - mProfit, _
                                mProfitThisTrade, _
                                mMaxProfitThisTrade - mProfitThisTrade, _
                                mTicker.Timestamp, _
                                msg, _
                                EndSession) _
    Then
        If Not mClosingPosition Then
            GLogMoneyManagement "Close all positions: " & msg, mIsSimulated, Me
            ClosePositions
        End If
    Else
        If EndSession And Not mClosingPosition Then
            If Not mEndSession Then
                mEndSession = True
                GLogMoneyManagement "Ending Session: " & msg, mIsSimulated, Me
            End If
        End If
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireChange( _
                ByVal changeType As PositionManagerChangeTypes)
Dim listener As ChangeListener
Dim Change As ChangeEvent
Const ProcName As String = "fireChange"
Dim failpoint As Long
On Error GoTo Err

Set Change.source = Me
Change.changeType = changeType
For Each listener In mChangeListeners
    listener.Change Change
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireProfit( _
                ByVal ProfitAmount As Double, _
                ByVal profitType As ProfitTypes)
Dim listener As ProfitListener
Dim Profit As ProfitEvent
Const ProcName As String = "fireProfit"
Dim failpoint As Long
On Error GoTo Err

Set Profit.source = Me
Profit.ProfitAmount = ProfitAmount
Profit.profitType = profitType
Profit.currencyCode = mcurrencyCode
For Each listener In mProfitListeners
    listener.ProfitAmount Profit
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub placeBalancingOrder()
Const ProcName As String = "placeBalancingOrder"
Dim failpoint As Long
On Error GoTo Err

GLogOrder _
        "Close positions at market: (Curr price=" & mTicker.FormatPrice(mTicker.TradePrice) & ")", _
        mIsSimulated, _
        Me

mBalancingOrderContext.Enabled = True
If mBalancingOrderContext.IsOrderTifValueSupported(TIFDay) Then mBalancingOrderContext.TimeInForce = TIFDay
Set mBalancingOrderPlex = mBalancingOrderContext.CreateOrderPlex( _
                            IIf(PositionSize < 0, ActionBuy, ActionSell), _
                            Abs(PositionSize), _
                            EntryOrderTypes.EntryOrderTypeMarket, _
                            0, _
                            0, _
                            0, _
                            StopOrderTypes.StopOrderTypeNone, _
                            0, _
                            0, _
                            0, _
                            TargetOrderTypes.TargetOrderTypeNone, _
                            0, _
                            0, _
                            0)
mBalancingOrderContext.ExecuteOrderPlex mBalancingOrderPlex

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub processExecution(ByVal pExec As Execution)

Const ProcName As String = "NotifyExecution"
Dim failpoint As Long
On Error GoTo Err

If mQuantityOutstanding = 0 Then
    mTradeStartTime = pExec.Time
    mTradeStartTimezone = pExec.Timezone
End If

mQuantityOutstanding = mQuantityOutstanding + pExec.SignedQuantity

If pExec.SignedQuantity > 0 Then
    mCumBuyPrice = mCumBuyPrice + (pExec.price * pExec.SignedQuantity)
Else
    mCumSellPrice = mCumSellPrice + (pExec.price * pExec.Quantity)
End If

If mQuantityOutstanding <> 0 Then
    'we need to notify the position change before the Profit
    GLogPosition mQuantityOutstanding, mIsSimulated, Me

    If mQuantityOutstanding < 0 Then
        If mLastAsk <> 0 Then
            calcTradeProfit mLastAsk
        End If
    Else        ' mQuantityOutstanding > 0
        If mLastBid <> 0 Then
            calcTradeProfit mLastBid
        End If
    End If
Else
    'we need to notify the position change after the final Profit
    calcTradeProfit
    
    GLogPosition mQuantityOutstanding, mIsSimulated, Me
    GLogTradeProfile mTradeStartTime & "," & _
            pExec.Time & "," & _
            mTradeStartTimezone & "," & _
            mProfitThisTrade & "," & _
            mMaxProfitThisTrade & "," & _
            mMaxLossThisTrade, _
            mIsSimulated, _
            Me
    mProfitLastTrade = mProfitThisTrade
    mDrawdownLastTrade = mMaxProfitThisTrade - mProfitThisTrade
    mProfitThisTrade = 0
    mMaxProfitThisTrade = 0
    mMaxLossThisTrade = 0
    mProfitLastFlat = mProfit
End If

If mQuantityOutstanding = 0 Then
    If Not mMoneyManager Is Nothing Then mMoneyManager.Flat
End If

fireChange PositionManagerChangeTypes.PositionSizeChanged
RaiseEvent PositionChanged

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub



