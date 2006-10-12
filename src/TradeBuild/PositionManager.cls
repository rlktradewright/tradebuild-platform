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

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

Implements IOrderSubmissionSrvcConsumer
Implements QuoteListener

'================================================================================
' Events
'================================================================================

Event Error(ByRef ev As ErrorEvent)

Event PositionChanged()

Event ProviderReady()

Event ProviderNotReady()

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

Public Enum ReasonTypes
    ReasonNone = 0
    ReasonEntry
    ReasonTarget
    ReasonStop
End Enum

'================================================================================
' Types
'================================================================================

Public Type OrderPlexProfile
    NextActiveIndex     As Long
    PrevActiveIndex     As Long
    key                 As String
    action              As OrderActions
    quantity            As Long
    EntryPrice          As Double
    ExitPrice           As Double
    QuantityOutstanding As Long
    TotalEntryPrice     As Double
    profit              As Currency
    MaxProfit           As Currency
    MaxLoss             As Currency
    risk                As Currency
    StartTime           As Date
    endTime             As Date
    Description         As String
    entryReason         As String
    stopReason          As String
    targetReason        As String
    closedOut           As Boolean
End Type

'================================================================================
' Member variables
'================================================================================

Private mIndexApplication

Private mTicker As Ticker

Private mOrders As Orders

Private mDummy As Boolean
Private mSimulated As Boolean

Private mOrderContexts As OrderContexts
Private mMoneyManager As MoneyManager
Private mStudies As studies

Private mPlexTable() As OrderPlexProfile
Private mNextPlexTableIndex As Long
Private mFirstActivePlexTableIndex As Long
Private mLastActivePlexTableIndex As Long

Private mSettlingOrderContext As OrderContext
Private mSettlingOrderPlex As orderPlex
Private mSettlingHandle As Long

Private mClosing As Boolean
    ' indicates that all OrderContexts are being closed, either as a result of
    ' a call to the closePosition functions, or because money management
    ' has told us to do so
Private mEndSession As Boolean
    ' indicates that money management has told us to end the session. Note
    ' that this can happen whilst there are still some OrderContexts in progress,
    ' for example if the daily profit has been met. In this case, the session
    ' is ended when all OrderContexts have completed.

Private mProfitIncrementTicks As Long
Private mProfitIncrement As Currency
    ' specifies the range of the bars in the profit profile
Private mProfitProfileBarStartTime As Date
Private mProfitProfileBarOpen As Currency
    ' the profit at the start of the current profit profile bar
Private mProfitProfileBarHigh As Currency
Private mProfitProfileBarLow As Currency
Private mProfitProfileBarClose As Currency
Private mProfitProfileBarNumber As Long

Private mTotalEntryPrice As Double
Private mQuantityOutstanding As Long
Private mProfit As Currency
Private mMaxProfit As Currency
Private mDrawdown As Currency

Private mProfitLastFlat As Currency
Private mProfitThisTrade As Currency
Private mMaxProfitThisTrade As Currency
Private mMinProfitThisTrade As Currency
Private mDrawdownThisTrade As Currency
Private mTradeStartTime As Date

Private mProfitLastTrade As Currency
Private mDrawdownLastTrade As Currency

Private mTickValue As Currency
Private mTickSize As Double
Private mcurrencyCode As String

Private mLastBid As Double
Private mLastAsk As Double

Private mProfitListeners As WeakCollection
Private mChangeListeners As WeakCollection

Private mOrderSubmitter As TradeBuildSP.IOrderSubmitter


'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mIndexApplication = GPositionManager.gNextApplicationIndex

Set mOrders = New Orders

Set mOrderContexts = New OrderContexts
mOrderContexts.PositionManager = Me

Set mStudies = New studies

ReDim mPlexTable(1023) As OrderPlexProfile
mFirstActivePlexTableIndex = 0
mLastActivePlexTableIndex = 1
mNextPlexTableIndex = 2
mPlexTable(mFirstActivePlexTableIndex).NextActiveIndex = mLastActivePlexTableIndex
mPlexTable(mLastActivePlexTableIndex).PrevActiveIndex = mFirstActivePlexTableIndex

Set mProfitListeners = New WeakCollection
Set mChangeListeners = New WeakCollection

mSimulated = True
End Sub

Private Sub Class_Terminate()
Debug.Print "PositionManager terminated"
End Sub

'================================================================================
' TradeBuildSP.IOrderSubmissionSrvcConsumer Interface Members
'================================================================================

Private Sub IOrderSubmissionSrvcConsumer_cancelTickData( _
                ByVal consumer As TradeBuildSP.IStreamingDataConsumer)
mTicker.removeDataConsumer consumer
End Sub

Private Sub IOrderSubmissionSrvcConsumer_Error( _
                ByVal errorCode As TradeBuildSP.StandardSPErrorCodes, _
                ByVal errorMessage As String, _
                ByVal spHandle As Long)
Dim ev As ErrorEvent
Select Case errorCode
Case LOCantConnectToBroker
    Set ev.source = Me
    ev.errorCode = ApiErrorCodes.ApiErrCantConnectToBroker
    ev.errorMsg = "Can't connect to broker: " & _
                    errorMessage
    RaiseEvent Error(ev)
Case LORetryConnectToBroker
    Set ev.source = Me
    ev.errorCode = ApiErrorCodes.ApiErrRetryConnectToBroker
    ev.errorMsg = "Trying again to connect to broker: " & _
                    errorMessage
    RaiseEvent Error(ev)
Case LOLostConnectionToBroker
    Set ev.source = Me
    ev.errorCode = ApiErrorCodes.ApiErrLostConnectionToBroker
    ev.errorMsg = "Lost connection to broker: " & _
                    errorMessage
    RaiseEvent Error(ev)
Case LOReConnectingToBroker
    Set ev.source = Me
    ev.errorCode = ApiErrorCodes.ApiErrReConnectingToBroker
    ev.errorMsg = "Reconnecting to broker: " & _
                    errorMessage
    RaiseEvent Error(ev)
Case LOUnknownContract
    Set ev.source = Me
    ev.errorCode = ApiErrorCodes.ApiErrContractSpecifierInvalid
    ev.errorMsg = "Unknown contract: " & _
                    errorMessage
    RaiseEvent Error(ev)
Case LOInvalidOrder
    Set ev.source = Me
    ev.errorCode = ApiErrorCodes.ApiErrInvalidOrder
    ev.errorMsg = "Invalid order: " & _
                    errorMessage
    RaiseEvent Error(ev)
Case LONotUniqueContract
    Set ev.source = Me
    ev.errorCode = ApiErrorCodes.ApiErrContractSpecifierAmbiguous
    ev.errorMsg = "Not unique contract: " & _
                    errorMessage
    RaiseEvent Error(ev)
Case LOInsufficientFunds
    Set ev.source = Me
    ev.errorCode = ApiErrorCodes.ApiErrInsufficientFunds
    ev.errorMsg = "Insufficient funds for this order: " & _
                    errorMessage
    RaiseEvent Error(ev)
Case LODisconnectedFromBroker
    Set ev.source = Me
    ev.errorCode = ApiErrorCodes.ApiErrDisconnectedFromBroker
    ev.errorMsg = "Disconnected from broker: " & _
                    errorMessage
    RaiseEvent Error(ev)
End Select
End Sub

Private Sub IOrderSubmissionSrvcConsumer_executionDetails( _
                ByVal pContractSpecifier As TradeBuildSP.IContractSpecifier, _
                ByVal exec As TradeBuildSP.IExecution)
Dim lExec As Execution
Dim lOrder As TradeBuild.order

If mDummy Then
    On Error Resume Next
    Set lExec = mTicker.dummyExecutions.Item(exec.ExecId)
    On Error GoTo 0
    
    If lExec Is Nothing Then
        mTicker.dummyExecutions.add exec
        Set lOrder = mOrders.Item(exec.orderTradeBuildId)
        lOrder.notifyFill exec
    End If
Else
    On Error Resume Next
    Set lExec = mTicker.executions.Item(exec.ExecId)
    On Error GoTo 0
    
    If lExec Is Nothing Then
        mTicker.executions.add exec
        Set lOrder = mOrders.Item(exec.orderTradeBuildId)
        lOrder.notifyFill exec
    End If
End If
End Sub

Private Sub IOrderSubmissionSrvcConsumer_notReady()
mOrderContexts.ProviderReady = False
RaiseEvent ProviderNotReady
End Sub

Private Sub IOrderSubmissionSrvcConsumer_orderStatus( _
                ByVal orderId As String, _
                ByVal status As TradeBuildSP.OrderStatuses)
Dim lOrder As TradeBuild.order
Set lOrder = mOrders.Item(orderId)
lOrder.updateStatus status
End Sub

Private Sub IOrderSubmissionSrvcConsumer_preFill( _
                ByVal orderId As String, _
                fillPrice As Double, _
                fillSize As Long)

End Sub

Private Sub IOrderSubmissionSrvcConsumer_ready()
mOrderContexts.ProviderReady = True
RaiseEvent ProviderReady
End Sub

Private Sub IOrderSubmissionSrvcConsumer_rejectOrder( _
                ByVal orderId As String, _
                ByVal reason As String)
Dim lOrder As TradeBuild.order
Set lOrder = mOrders.Item(orderId)
lOrder.updateStatus OrderStatuses.OrderStatusRejected
End Sub

Private Sub IOrderSubmissionSrvcConsumer_requestTickData( _
                ByVal consumer As TradeBuildSP.IStreamingDataConsumer, _
                ByVal includeMarketDepth As Boolean)
mTicker.addDataConsumer consumer
End Sub

'================================================================================
' QuoteListener Interface Members
'================================================================================

Private Sub QuoteListener_ask(ev As QuoteEvent)
Dim i As Long

If ev.price = 0 Then Exit Sub ' can occur at start of session

If ev.price = mLastAsk Then Exit Sub

mLastAsk = ev.price

i = mPlexTable(mFirstActivePlexTableIndex).NextActiveIndex
Do While i <> mLastActivePlexTableIndex
    calcPlexProfit i, ev.price
    i = mPlexTable(i).NextActiveIndex
Loop

If mQuantityOutstanding < 0 Then
    calcTradeProfit ev.price
    checkCloseall
End If

End Sub

Private Sub QuoteListener_bid(ev As QuoteEvent)
Dim i As Long

If ev.price = 0 Then Exit Sub  ' can occur at start of session

If ev.price = mLastBid Then Exit Sub

mLastBid = ev.price

i = mPlexTable(mFirstActivePlexTableIndex).NextActiveIndex
Do While i <> mLastActivePlexTableIndex
    calcPlexProfit i, ev.price
    i = mPlexTable(i).NextActiveIndex
Loop

If mQuantityOutstanding > 0 Then
    calcTradeProfit ev.price
    checkCloseall
End If

End Sub

Private Sub QuoteListener_High(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_Low(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_openInterest(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_PreviousClose(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_trade(ev As QuoteEvent)
Dim i As Long

If ev.price = 0 Then Exit Sub  ' can occur at start of session

If mQuantityOutstanding = 0 Then Exit Sub
If ev.price >= mLastAsk Then
    If mQuantityOutstanding < 0 Then
        mLastAsk = ev.price ' because the ask price must have gone at least that high
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
    
i = mPlexTable(mFirstActivePlexTableIndex).NextActiveIndex
Do While i <> mLastActivePlexTableIndex
    calcPlexProfit i, ev.price
    i = mPlexTable(i).NextActiveIndex
Loop

calcTradeProfit ev.price
checkCloseall

End Sub

Private Sub QuoteListener_volume(ev As QuoteEvent)

End Sub
'================================================================================
' Properties
'================================================================================

Public Property Get drawdown() As Currency
drawdown = mMaxProfit - mProfit
End Property

Public Property Get drawdownlastTrade() As Currency
drawdownlastTrade = mDrawdownLastTrade
End Property

Public Property Get drawdownThisTrade() As Currency
drawdownThisTrade = mMaxProfitThisTrade - mProfitThisTrade
End Property

Friend Property Let dummy(ByVal value As Boolean)
mDummy = value
mOrderContexts.dummy = value
End Property

Public Property Get indexApplication() As Long
indexApplication = mIndexApplication
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
Set mOrderSubmitter = value
mOrderContexts.orderSubmitter = value
End Property

Public Property Get pendingPositionSize() As Integer
Dim OrderContext As OrderContext
Dim Size As Long

For Each OrderContext In mOrderContexts
    Size = Size + OrderContext.pendingSize
Next

pendingPositionSize = Size
End Property

Public Property Get positionSize() As Integer
positionSize = mQuantityOutstanding
End Property

Public Property Get profit() As Currency
profit = mProfit
End Property

Public Property Let profitProfileIncrementTicks(ByVal value As Long)
mProfitIncrementTicks = value
If mTickValue <> 0 Then mProfitIncrement = mProfitIncrementTicks * mTickValue
End Property

Public Property Get profitLastTrade() As Currency
profitLastTrade = mProfitLastTrade
End Property

Public Property Get profitThisTrade() As Currency
profitThisTrade = mProfitThisTrade
End Property

Public Property Get risk() As Currency
Dim OrderContext As OrderContext
Dim totRisk As Currency

For Each OrderContext In mOrderContexts
    totRisk = totRisk + OrderContext.risk
Next

risk = totRisk
End Property

Friend Property Let simulated(ByVal value As Boolean)
mSimulated = value
mOrderContexts.simulated = value
End Property

Public Property Get simulated() As Boolean
simulated = mSimulated
End Property

Friend Property Let Ticker(ByVal value As Ticker)
Dim lContract As Contract

Set mTicker = value
Set lContract = mTicker.Contract

mOrderContexts.Ticker = value

mTicker.addQuoteListener Me

mTickValue = lContract.tickValue
mTickSize = lContract.TickSize
mProfitIncrement = mProfitIncrementTicks * mTickValue
mcurrencyCode = lContract.specifier.currencyCode

End Property

Public Property Get Ticker() As Ticker
Set Ticker = mTicker
End Property

'================================================================================
' Methods
'================================================================================

Public Sub addChangeListener(ByVal value As ChangeListener)
mChangeListeners.add value
End Sub

Public Sub addProfitListener(ByVal value As ProfitListener)
mProfitListeners.add value
End Sub

Public Sub AddStudy( _
                ByVal study As TradeBuild.study, _
                ByVal ValueName As String)
mStudies.add study, ValueName
End Sub

Friend Function allowOrder(ByVal orderLong As Boolean, _
                            ByRef orderSize As Long, _
                            ByVal newRisk As Currency, _
                            ByRef message As String) As Boolean

If mClosing Then
    allowOrder = False
    message = "Closing positions"
    Exit Function
End If

If Not mMoneyManager Is Nothing Then
    allowOrder = mMoneyManager.allowOrder(positionSize, _
                                        pendingPositionSize, _
                                        orderLong, _
                                        orderSize, _
                                        newRisk + risk, _
                                        mProfit, _
                                        mMaxProfit - mProfit, _
                                        mProfitThisTrade, _
                                        mMaxProfitThisTrade - mProfitThisTrade, _
                                        mTicker.timestamp, _
                                        message, _
                                        mEndSession)
Else
    allowOrder = True
End If
End Function

Public Sub closePositions()
Dim OrderContext As OrderContext

If mClosing Then Exit Sub

' first need to tell each OrderContext to cancel any outstanding stuff
For Each OrderContext In mOrderContexts
    If OrderContext.Cancel Then
        mClosing = True
        mOrderContexts.enabled = False
    End If
Next

If mClosing Then
Else
    If positionSize <> 0 Then
        
        mTicker.notifyListeners _
                "Close positions: (" & mTicker.TradePriceString & ") at market", _
                IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                Me
        
        Set mSettlingOrderContext = mOrderContexts.add("$settling")
        If mSettlingOrderContext.isOrderTifValueSupported(TIFDay) Then mSettlingOrderContext.timeInForce = TIFDay
        Set mSettlingOrderPlex = mSettlingOrderContext.CreateOrderPlex( _
                            IIf(positionSize < 0, ActionBuy, ActionSell), _
                            Abs(positionSize), _
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
        mSettlingOrderContext.executeOrderPlex mSettlingOrderPlex
    End If
End If

End Sub

Friend Sub endOfSession(ByVal timestamp As Date)
If mProfitIncrement = 0 Then Exit Sub
mTicker.notifyListeners mProfitProfileBarStartTime & "," & _
                mProfitProfileBarNumber & "," & _
                mProfitProfileBarOpen & "," & _
                mProfitProfileBarHigh & "," & _
                mProfitProfileBarLow & "," & _
                mProfitProfileBarClose, _
                IIf(mDummy, _
                    TradeBuildListenValueTypes.VTDummyProfitProfile, _
                    TradeBuildListenValueTypes.VTProfitProfile), _
                Me
mTicker.notifyListeners timestamp & "," & _
                mProfitProfileBarNumber + 1 & "," & _
                mProfit & "," & _
                mProfit & "," & _
                mProfit & "," & _
                mProfit, _
                IIf(mDummy, _
                    TradeBuildListenValueTypes.VTDummyProfitProfile, _
                    TradeBuildListenValueTypes.VTProfitProfile), _
                Me
End Sub

Friend Sub finish()

Set mTicker = Nothing

mOrderContexts.finish

Set mOrderContexts = Nothing
Set mMoneyManager = Nothing
Set mSettlingOrderContext = Nothing
Set mSettlingOrderPlex = Nothing

If Not mOrderSubmitter Is Nothing Then mOrderSubmitter.finish
Set mOrderSubmitter = Nothing
End Sub

Friend Sub notifyExecution(ByVal quantity As Long, _
                            ByVal price As Double, _
                            ByVal handle As Long, _
                            ByVal closedOut As Boolean)

If mQuantityOutstanding = 0 Then mTradeStartTime = mTicker.timestamp

mTotalEntryPrice = mTotalEntryPrice + (price * quantity)
mQuantityOutstanding = mQuantityOutstanding + quantity

If mQuantityOutstanding <> 0 Then
    'we need to notify the position change before the profit
    mTicker.notifyListeners _
                    mQuantityOutstanding, _
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyPosition, TradeBuildListenValueTypes.VTPosition), _
                    Me

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
    'we need to notify the position change after the final profit
    calcTradeProfit
    
    mTicker.notifyListeners _
                    mQuantityOutstanding, _
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyPosition, TradeBuildListenValueTypes.VTPosition), _
                    Me
    mTicker.notifyListeners _
                    mTradeStartTime & "," & _
                    mTicker.timestamp & "," & _
                    mProfitThisTrade & "," & _
                    mMaxProfitThisTrade & "," & _
                    mMinProfitThisTrade, _
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyTradeProfile, TradeBuildListenValueTypes.VTTradeProfile), _
                    Me
    mProfitLastTrade = mProfitThisTrade
    mDrawdownLastTrade = mMaxProfitThisTrade - mProfitThisTrade
    mProfitThisTrade = 0
    mMaxProfitThisTrade = 0
    mMinProfitThisTrade = 0
    mProfitLastFlat = mProfit
End If

With mPlexTable(handle)
    .closedOut = closedOut
    If .QuantityOutstanding = 0 Then .StartTime = mTicker.timestamp
    
    .TotalEntryPrice = .TotalEntryPrice + (price * quantity)
    
    .QuantityOutstanding = .QuantityOutstanding + quantity
    
    If Abs(.QuantityOutstanding) > Abs(.quantity) Then
        .quantity = .QuantityOutstanding
    End If
    
    If (.action = ActionBuy And quantity > 0) Or _
        (.action = ActionSell And quantity < 0) _
    Then
        .EntryPrice = .TotalEntryPrice / .quantity
    End If
        
    If .QuantityOutstanding < 0 Then
        If mLastAsk <> 0 Then
            calcPlexProfit handle, mLastAsk
        End If
    ElseIf .QuantityOutstanding > 0 Then
        If mLastBid <> 0 Then
            calcPlexProfit handle, mLastBid
        End If
    Else
        .endTime = mTicker.timestamp
        calcPlexProfit handle
        notifyPlexProfileListeners handle
        
        ' remove the relevant entry from the plex table
        mPlexTable(.PrevActiveIndex).NextActiveIndex = .NextActiveIndex
        mPlexTable(.NextActiveIndex).PrevActiveIndex = .PrevActiveIndex
        
    End If
End With

If mQuantityOutstanding = 0 Then
    If Not mMoneyManager Is Nothing Then mMoneyManager.flat
End If

fireChange PositionManagerChangeTypes.PositionSizeChanged
RaiseEvent PositionChanged
End Sub

Friend Sub OrderContextInactive(ByVal pOrderContext As OrderContext)

Dim OrderContext As OrderContext
Dim Size As Long


If Not mClosing Then
    For Each OrderContext In mOrderContexts
        If Not OrderContext.Complete Then Exit Sub
    Next
    ' all OrderContexts are complete
    checkCloseall
    If mEndSession Then
        mTicker.notifyListeners "Ending session", _
                        TradeBuildListenValueTypes.VTMoneyManagement, _
                        Me
    End If
    Exit Sub
End If

If Not mSettlingOrderContext Is Nothing Then
    If pOrderContext Is mSettlingOrderContext Then
        'should now be net flat
        Set mSettlingOrderContext = Nothing
        mClosing = False
        mOrderContexts.enabled = True
        
        ' now need to adjust the other orderPlex table
        ' entries' profit figure to take account of the
        ' settling order
        
        Dim i As Long
        Dim settlingPrice As Double
        
        ' first remove the settling entry from the plex table
        With mPlexTable(mSettlingHandle)
            mPlexTable(.PrevActiveIndex).NextActiveIndex = .NextActiveIndex
            mPlexTable(.NextActiveIndex).PrevActiveIndex = .PrevActiveIndex
        End With
        
        settlingPrice = mPlexTable(mSettlingHandle).TotalEntryPrice / mPlexTable(mSettlingHandle).QuantityOutstanding
        i = mPlexTable(mFirstActivePlexTableIndex).NextActiveIndex
        Do While i <> mLastActivePlexTableIndex
            With mPlexTable(i)
                .endTime = mTicker.timestamp
                .ExitPrice = settlingPrice
                .closedOut = True
                .TotalEntryPrice = .TotalEntryPrice - (settlingPrice * .QuantityOutstanding)
                .QuantityOutstanding = 0
                calcPlexProfit i
                notifyPlexProfileListeners i
            
                i = .NextActiveIndex
                
                ' remove the current entry from the plex table
                mPlexTable(.PrevActiveIndex).NextActiveIndex = .NextActiveIndex
                mPlexTable(.NextActiveIndex).PrevActiveIndex = .PrevActiveIndex
                
            End With
        Loop
        
        
        
        If mEndSession Then
            mTicker.notifyListeners "Ending session", _
                                TradeBuildListenValueTypes.VTMoneyManagement, _
                                Me
        End If
        Exit Sub
    End If
End If

For Each OrderContext In mOrderContexts
    If Not OrderContext.Complete Then Exit Sub
    Size = Size + OrderContext.Size
Next

' If we get here, all OrderContexts must be complete.

' Now, if there is a net position, need to close it out

If Size = 0 Then
    mClosing = False
    mOrderContexts.enabled = True
    If mEndSession Then
        mTicker.notifyListeners "Ending session", _
                            TradeBuildListenValueTypes.VTMoneyManagement, _
                            Me
    End If
Else
    
    mTicker.notifyListeners _
            "Close positions: (" & mTicker.TradePriceString & ") at market", _
            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
            Me
    
    Set mSettlingOrderContext = mOrderContexts.add("$settling")
    If mSettlingOrderContext.isOrderTifValueSupported(TIFDay) Then mSettlingOrderContext.timeInForce = TIFDay
    Set mSettlingOrderPlex = mSettlingOrderContext.CreateOrderPlex( _
                                IIf(positionSize < 0, ActionBuy, ActionSell), _
                                Abs(positionSize), _
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
    mSettlingOrderContext.executeOrderPlex mSettlingOrderPlex
End If

End Sub

Friend Function registerPlex(ByVal key As String, _
                            ByVal Description As String, _
                            ByVal action As OrderActions, _
                            ByVal risk As Currency) As Long
If mNextPlexTableIndex >= UBound(mPlexTable) Then
    ReDim Preserve mPlexTable(UBound(mPlexTable) + 1024) As OrderPlexProfile
End If
registerPlex = mNextPlexTableIndex
mNextPlexTableIndex = mNextPlexTableIndex + 1
With mPlexTable(registerPlex)
    .key = key
    .Description = Description
    .action = action
    .risk = risk
    .NextActiveIndex = mLastActivePlexTableIndex
    .PrevActiveIndex = mPlexTable(mLastActivePlexTableIndex).PrevActiveIndex
    mPlexTable(.PrevActiveIndex).NextActiveIndex = registerPlex
    mPlexTable(mLastActivePlexTableIndex).PrevActiveIndex = registerPlex
End With
If Not mSettlingOrderPlex Is Nothing Then
    If key = mSettlingOrderPlex.key Then
        mSettlingHandle = registerPlex
    End If
End If
End Function

Public Sub removeChangeListener(ByVal value As ChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.remove i
Next
End Sub

Public Sub removeProfitListener(ByVal value As ProfitListener)
Dim i As Long
For i = mProfitListeners.Count To 1 Step -1
    If mProfitListeners.Item(i) Is value Then mProfitListeners.remove i
Next
End Sub

Friend Function setPlexReason(ByVal handle As Long, _
                            ByVal reasonType As ReasonTypes, _
                            ByVal reason As String)

With mPlexTable(handle)
    Select Case reasonType
    Case ReasonNone
    Case ReasonEntry
        .entryReason = reason
    Case ReasonTarget
        .targetReason = reason
    Case ReasonStop
        .stopReason = reason
    End Select
End With
End Function

Friend Sub setPlexRisk(ByVal handle As Long, _
                            ByVal risk As Currency)
mPlexTable(handle).risk = risk
End Sub

Friend Sub startOfSession(ByVal timestamp As Date)
mProfitProfileBarStartTime = timestamp
End Sub

Friend Sub TickfileCompleted()
Dim i As Long
i = mPlexTable(mFirstActivePlexTableIndex).NextActiveIndex
Do While i <> mLastActivePlexTableIndex
    With mPlexTable(i)
        .endTime = mTicker.timestamp
        notifyPlexProfileListeners i
    
        i = .NextActiveIndex
        
        ' remove the current entry from the plex table
        mPlexTable(.PrevActiveIndex).NextActiveIndex = .NextActiveIndex
        mPlexTable(.NextActiveIndex).PrevActiveIndex = .PrevActiveIndex
        
    End With
Loop

End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcPlexProfit(ByVal handle As Long, _
                            Optional ByVal price As Double)

With mPlexTable(handle)
    If .QuantityOutstanding <> 0 Then
        .profit = mTickValue * .QuantityOutstanding * (price - .TotalEntryPrice / .QuantityOutstanding) / mTickSize
        .ExitPrice = price
    Else
        .profit = -1 * mTickValue * .TotalEntryPrice / mTickSize
        .ExitPrice = -(.TotalEntryPrice - .quantity * .EntryPrice) / .quantity
    End If
    
    If .profit > .MaxProfit Then .MaxProfit = .profit
    If .profit < .MaxLoss Then .MaxLoss = .profit
    
End With
End Sub

Private Sub calcTradeProfit(Optional ByVal price As Double)

Dim profit As Currency
Dim ev As TradeBuild.StudyValueEvent

If mQuantityOutstanding <> 0 Then
    profit = mTickValue * mQuantityOutstanding * (price - mTotalEntryPrice / mQuantityOutstanding) / mTickSize
Else
    profit = -1 * mTickValue * mTotalEntryPrice / mTickSize
End If

If profit <> mProfit Then
    mProfit = profit
    
    fireProfit mProfit, ProfitTypes.ProfitTypeSessionProfit
    
    Set ev.source = Me
    ev.timestamp = mTicker.timestamp
    ev.value = mProfit
    ev.ValueName = "profit"
    mStudies.notify ev
    
    mTicker.notifyListeners mProfit, _
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyProfit, TradeBuildListenValueTypes.VTProfit), _
                    Me
    mTicker.notifyListeners mMaxProfit - mProfit, _
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyDrawdown, TradeBuildListenValueTypes.VTDrawdown), _
                    Me
End If

If mProfit > mMaxProfit Then
    mMaxProfit = mProfit
    fireProfit mMaxProfit, ProfitTypes.ProfitTypeSessionMaxProfit
    mTicker.notifyListeners mMaxProfit, _
                        IIf(mDummy, TradeBuildListenValueTypes.VTDummyMaxProfit, TradeBuildListenValueTypes.VTMaxProfit), _
                        Me
ElseIf mProfit < mMaxProfit Then
    mDrawdown = mMaxProfit - mProfit
    fireProfit mDrawdown, ProfitTypes.ProfitTypeSessionDrawdown
ElseIf mDrawdown <> 0 Then
    mDrawdown = 0
    fireProfit mDrawdown, ProfitTypes.ProfitTypeSessionDrawdown
End If

mProfitThisTrade = mProfit - mProfitLastFlat
fireProfit mProfitThisTrade, ProfitTypes.ProfitTypetradeProfit
If mProfitThisTrade > mMaxProfitThisTrade Then
    mMaxProfitThisTrade = mProfitThisTrade
    fireProfit mMaxProfitThisTrade, ProfitTypes.ProfitTypeTradeMaxProfit
ElseIf (mMaxProfitThisTrade - mProfitThisTrade) < mDrawdownThisTrade Then
    mDrawdownThisTrade = mMaxProfitThisTrade - mProfitThisTrade
    fireProfit mDrawdownThisTrade, ProfitTypes.ProfitTypetradeDrawdown
End If
If mProfitThisTrade < mMinProfitThisTrade Then mMinProfitThisTrade = mProfitThisTrade

If mProfitIncrement = 0 Then Exit Sub

If mProfit - mProfitProfileBarLow >= mProfitIncrement Or _
    mProfitProfileBarHigh - mProfit >= mProfitIncrement _
Then
    ' start a new profit profile bar
    mTicker.notifyListeners mProfitProfileBarStartTime & "," & _
                        mProfitProfileBarNumber & "," & _
                        mProfitProfileBarOpen & "," & _
                        mProfitProfileBarHigh & "," & _
                        mProfitProfileBarLow & "," & _
                        mProfitProfileBarClose, _
                        IIf(mDummy, TradeBuildListenValueTypes.VTDummyProfitProfile, TradeBuildListenValueTypes.VTProfitProfile), _
                        Me
    mProfitProfileBarStartTime = mTicker.timestamp
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

End Sub

Private Sub checkCloseall()

Dim msg As String
Dim endSession As Boolean

If mDummy Then Exit Sub

If Not mMoneyManager Is Nothing Then
    If mMoneyManager.closeAll(mProfit, _
                                mMaxProfit - mProfit, _
                                mProfitThisTrade, _
                                mMaxProfitThisTrade - mProfitThisTrade, _
                                mTicker.timestamp, _
                                msg, _
                                endSession) _
    Then
        If Not mClosing Then
            mTicker.notifyListeners "Close all positions: " & msg, _
                            TradeBuildListenValueTypes.VTMoneyManagement, _
                            Me
            closePositions
        End If
    Else
        If endSession And Not mClosing Then
            If Not mEndSession Then
                mEndSession = True
                mTicker.notifyListeners "Ending session: " & msg, _
                                TradeBuildListenValueTypes.VTMoneyManagement, _
                                Me
            End If
        End If
    End If
End If
End Sub

Private Sub fireChange( _
                ByVal changeType As PositionManagerChangeTypes)
Dim listener As ChangeListener
Dim i As Long
Dim Change As ChangeEvent
Set Change.source = Me
Change.changeType = changeType
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change Change
Next
End Sub

Private Sub fireProfit( _
                ByVal profitAmount As Double, _
                ByVal profitType As ProfitTypes)
Dim listener As ProfitListener
Dim i As Long
Dim profit As ProfitEvent
Set profit.source = Me
profit.profitAmount = profitAmount
profit.profitType = profitType
profit.currencyCode = mcurrencyCode
For i = 1 To mProfitListeners.Count
    Set listener = mProfitListeners(i)
    listener.profitAmount profit
Next
End Sub

Private Sub notifyPlexProfileListeners(ByVal index As Long)
Dim data As Variant
data = mPlexTable(index)
mTicker.notifyListeners _
                data, _
                IIf(mDummy, _
                    TradeBuildListenValueTypes.VTDummyOrderPlexProfileStruct, _
                    TradeBuildListenValueTypes.VTOrderPlexProfileStruct), _
                Me

With mPlexTable(index)
    mTicker.notifyListeners _
                FormatDateTime(.StartTime) & "," & _
                FormatDateTime(.endTime) & "," & _
                .Description & "," & _
                .key & "," & _
                gOrderActionToString(.action) & "," & _
                .quantity & "," & _
                .EntryPrice & "," & _
                .ExitPrice & "," & _
                .entryReason & "," & _
                .targetReason & "," & _
                .stopReason & "," & _
                IIf(.closedOut, "1,", "0,") & _
                .profit & "," & _
                .MaxProfit & "," & _
                .MaxLoss & "," & _
                .risk & "," & _
                .QuantityOutstanding, _
                IIf(mDummy, _
                    TradeBuildListenValueTypes.VTDummyOrderPlexProfileString, _
                    TradeBuildListenValueTypes.VTOrderPlexProfileString), _
                Me
End With
End Sub



