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

Public Type OrderPlexProfile
    NextActiveIndex     As Long
    PrevActiveIndex     As Long
    Key                 As String
    action              As OrderActions
    quantity            As Long
    EntryPrice          As Double
    ExitPrice           As Double
    QuantityOutstanding As Long
    TotalEntryPrice     As Double
    Profit              As Currency
    MaxProfit           As Currency
    MaxLoss             As Currency
    Risk                As Currency
    StartTime           As Date
    endTime             As Date
    Description         As String
    entryReason         As String
    stopReason          As String
    targetReason        As String
    closedOut           As Boolean
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mIndexApplication As Long

Private mProviderIsReady As Boolean

Private mTicker As Ticker

Private mOrders As Orders

Private mIsSimulated As Boolean

Private mOrderContexts As OrderContexts
Private mMoneyManager As MoneyManager
Private mStudies As Studies

Private mPlexTable() As OrderPlexProfile
Private mNextPlexTableIndex As Long
Private mFirstActivePlexTableIndex As Long
Private mLastActivePlexTableIndex As Long

Private mBalancingOrderContext As OrderContext
Private mBalancingOrderPlex As orderPlex
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

Private mProfitListeners As Collection
Private mChangeListeners As Collection

Private mOrderSubmitter As TradeBuildSP.IOrderSubmitter


'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mIndexApplication = GPositionManager.gNextApplicationIndex

Set mOrders = New Orders

Set mOrderContexts = New OrderContexts
mOrderContexts.PositionManager = Me
Set mBalancingOrderContext = mOrderContexts.Add("$balancing")

Set mStudies = New Studies

ReDim mPlexTable(3) As OrderPlexProfile
mFirstActivePlexTableIndex = 0
mLastActivePlexTableIndex = 1
mNextPlexTableIndex = 4
mPlexTable(mFirstActivePlexTableIndex).NextActiveIndex = mLastActivePlexTableIndex
mPlexTable(mLastActivePlexTableIndex).PrevActiveIndex = mFirstActivePlexTableIndex

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
mTicker.RemoveDataConsumer consumer
End Sub

Private Sub IOrderSubmissionSrvcConsumer_executionDetails( _
                ByVal pContractSpecifier As contractSpecifier, _
                ByVal accountIdentifier As String, _
                ByVal action As TradeBuildSP.OrderActions, _
                ByVal execId As String, _
                ByVal fillingExchange As String, _
                ByVal isLiquidation As Boolean, _
                ByVal orderBrokerId As String, _
                ByVal orderTradeBuildId As String, _
                ByVal price As Double, _
                ByVal quantity As Long, _
                ByVal fillTime As Date)
Dim lExec As Execution
Dim lOrder As order
Dim isNew As Boolean

If mIsSimulated Then
    On Error Resume Next
    Set lExec = mTicker.ExecutionsSimulated.item(execId)
    On Error GoTo 0
Else
    On Error Resume Next
    Set lExec = mTicker.Executions.item(execId)
    On Error GoTo 0
End If

If lExec Is Nothing Then
    Set lExec = New Execution
    lExec.execId = execId
    isNew = True
Else
    ' we've alReady had this notification
End If

With lExec
    .accountIdentifier = accountIdentifier
    .action = action
    .contractSpecifier = pContractSpecifier
    If mIsSimulated Then .IsSimulated = True
    .fillingExchange = fillingExchange
    .isLiquidation = isLiquidation
    .orderBrokerId = orderBrokerId
    .orderTradeBuildId = orderTradeBuildId
    .price = price
    .quantity = quantity
    .time = fillTime
End With

If isNew Then
    If mIsSimulated Then
        mTicker.ExecutionsSimulated.Add lExec
    Else
        mTicker.Executions.Add lExec
    End If
    Set lOrder = mOrders.item(lExec.orderTradeBuildId)
    lOrder.notifyFill lExec
Else
    If mIsSimulated Then
        mTicker.ExecutionsSimulated.Change lExec
    Else
        mTicker.Executions.Change lExec
    End If
End If
End Sub

Private Sub IOrderSubmissionSrvcConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)
gTB.ServiceProviders.CommonServiceConsumerNotifyEvent eventCode, eventMessage, spHandle
End Sub

Private Sub IOrderSubmissionSrvcConsumer_notReady()
mProviderIsReady = False
mOrderContexts.ProviderReady = False
mClosingPosition = False
fireChange ProviderReadinessChanged
RaiseEvent ProviderNotReady
End Sub

Private Sub IOrderSubmissionSrvcConsumer_OrderReport(ByVal OrderReport As TradeBuildSP.IOrderReport)
Dim lOrder As order
Dim different As Boolean

Dim failpoint As Long
On Error GoTo Err

mOrders(OrderReport.tradeBuildId).SyncToOrderReport OrderReport

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "IOrderSubmissionSrvcConsumer_OrderReport" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Private Sub IOrderSubmissionSrvcConsumer_orderStatus( _
                ByVal orderId As String, _
                ByVal status As TradeBuildSP.OrderStatuses)
Dim lOrder As order
Set lOrder = mOrders.item(orderId)
lOrder.updateStatus status
End Sub

Private Sub IOrderSubmissionSrvcConsumer_preFill( _
                ByVal orderId As String, _
                fillPrice As Double, _
                fillSize As Long)

End Sub

Private Sub IOrderSubmissionSrvcConsumer_Ready()
mProviderIsReady = True
mOrderContexts.ProviderReady = True
fireChange ProviderReadinessChanged
RaiseEvent ProviderReady
End Sub

Private Sub IOrderSubmissionSrvcConsumer_rejectOrder( _
                ByVal orderId As String, _
                ByVal reason As String)
Dim lOrder As order
Set lOrder = mOrders.item(orderId)
lOrder.updateStatus OrderStatuses.OrderStatusRejected
End Sub

Private Sub IOrderSubmissionSrvcConsumer_requestTickData( _
                ByVal consumer As TradeBuildSP.IStreamingDataConsumer, _
                ByVal includeMarketDepth As Boolean)
mTicker.AddDataConsumer consumer, includeMarketDepth
End Sub

'@================================================================================
' QuoteListener Interface Members
'@================================================================================

Private Sub QuoteListener_Ask(ev As QuoteEvent)
Dim i As Long

If ev.price = 0 Then Exit Sub ' can occur at start of Session

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

Private Sub QuoteListener_Bid(ev As QuoteEvent)
Dim i As Long

If ev.price = 0 Then Exit Sub  ' can occur at start of Session

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

Private Sub QuoteListener_OpenInterest(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_PreviousClose(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_sessionOpen(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_Trade(ev As QuoteEvent)
Dim i As Long

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
    
i = mPlexTable(mFirstActivePlexTableIndex).NextActiveIndex
Do While i <> mLastActivePlexTableIndex
    calcPlexProfit i, ev.price
    i = mPlexTable(i).NextActiveIndex
Loop

calcTradeProfit ev.price
checkCloseall

End Sub

Private Sub QuoteListener_Volume(ev As QuoteEvent)

End Sub
'@================================================================================
' Properties
'@================================================================================

Public Property Get Drawdown() As Currency
Drawdown = mMaxProfit - mProfit
End Property

Public Property Get DrawdownlastTrade() As Currency
DrawdownlastTrade = mDrawdownLastTrade
End Property

Public Property Get DrawdownThisTrade() As Currency
DrawdownThisTrade = mMaxProfitThisTrade - mProfitThisTrade
End Property

Public Property Get IndexApplication() As Long
IndexApplication = mIndexApplication
End Property

Public Property Get IsRiskUnlimited() As Boolean
IsRiskUnlimited = (Risk = MaxCurrency)
End Property

Friend Property Let IsSimulated(ByVal value As Boolean)
mIsSimulated = value
mOrderContexts.IsSimulated = value
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
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

Public Property Get PendingPositionSize() As Integer
Dim OrderContext As OrderContext
Dim size As Long

For Each OrderContext In mOrderContexts
    size = size + OrderContext.pendingSize
Next

PendingPositionSize = size
End Property

Public Property Get PositionSize() As Integer
PositionSize = mQuantityOutstanding
End Property

Public Property Get Profit() As Currency
Profit = mProfit
End Property

Public Property Let ProfitProfileIncrementTicks(ByVal value As Long)
mProfitIncrementTicks = value
If mTickValue <> 0 Then mProfitIncrement = mProfitIncrementTicks * mTickValue
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

For Each OrderContext In mOrderContexts
    ocRisk = OrderContext.Risk
    If ocRisk = MaxCurrency Then
        totRisk = MaxCurrency
        Exit For
    End If
    totRisk = totRisk + ocRisk
Next

Risk = totRisk
End Property

Friend Property Let Ticker(ByVal value As Ticker)
Dim lContract As Contract

Set mTicker = value
Set lContract = mTicker.Contract

mOrderContexts.Ticker = value

mTicker.AddQuoteListener Me

mTickValue = lContract.tickValue
mTickSize = lContract.TickSize
mProfitIncrement = mProfitIncrementTicks * mTickValue
mcurrencyCode = lContract.specifier.currencyCode

End Property

Public Property Get Ticker() As Ticker
Set Ticker = mTicker
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
mStudies.Add study, valueNames
End Sub

Friend Function AllowOrder(ByVal orderLong As Boolean, _
                            ByRef orderSize As Long, _
                            ByVal newRisk As Currency, _
                            ByRef message As String) As Boolean

If mClosingPosition Then
    AllowOrder = False
    message = "Closing positions"
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
                                        message, _
                                        mEndSession)
Else
    AllowOrder = True
End If
End Function

Public Sub ClosePositions( _
                Optional ByVal flags As ClosePositionFlags = ClosePositionCancelOrders Or ClosePositionWaitForCancel)
Dim lOrderContext As OrderContext
Dim cancellingOrders As Boolean

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

mOrderContexts.enabled = False  ' prevent any further orders being placed
                                ' while the position is being closed

If Not cancellingOrders And PositionSize <> 0 Then placeBalancingOrder

End Sub

Friend Sub EndOfSession(ByVal Timestamp As Date)
If mProfitIncrement = 0 Then Exit Sub
gTB.notifyListeners mProfitProfileBarStartTime & "," & _
                mProfitProfileBarNumber & "," & _
                mProfitProfileBarOpen & "," & _
                mProfitProfileBarHigh & "," & _
                mProfitProfileBarLow & "," & _
                mProfitProfileBarClose, _
                IIf(mIsSimulated, _
                    TradeBuildListenValueTypes.VTSimulatedProfitProfile, _
                    TradeBuildListenValueTypes.VTProfitProfile), _
                Me
gTB.notifyListeners Timestamp & "," & _
                mProfitProfileBarNumber + 1 & "," & _
                mProfit & "," & _
                mProfit & "," & _
                mProfit & "," & _
                mProfit, _
                IIf(mIsSimulated, _
                    TradeBuildListenValueTypes.VTSimulatedProfitProfile, _
                    TradeBuildListenValueTypes.VTProfitProfile), _
                Me
End Sub

Friend Sub Finish()

Set mTicker = Nothing

mOrderContexts.Finish

Set mOrderContexts = Nothing
Set mMoneyManager = Nothing
Set mBalancingOrderContext = Nothing
Set mBalancingOrderPlex = Nothing

If Not mOrderSubmitter Is Nothing Then mOrderSubmitter.Finish
Set mOrderSubmitter = Nothing
End Sub

Friend Sub NotifyExecution(ByVal quantity As Long, _
                            ByVal price As Double, _
                            ByVal Handle As Long, _
                            ByVal closedOut As Boolean)

If mQuantityOutstanding = 0 Then mTradeStartTime = mTicker.Timestamp

mTotalEntryPrice = mTotalEntryPrice + (price * quantity)
mQuantityOutstanding = mQuantityOutstanding + quantity

If mQuantityOutstanding <> 0 Then
    'we need to notify the position change before the Profit
    gTB.notifyListeners _
                    mQuantityOutstanding, _
                    IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedPosition, TradeBuildListenValueTypes.VTPosition), _
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
    'we need to notify the position change after the final Profit
    calcTradeProfit
    
    gTB.notifyListeners _
                    mQuantityOutstanding, _
                    IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedPosition, TradeBuildListenValueTypes.VTPosition), _
                    Me
    gTB.notifyListeners _
                    mTradeStartTime & "," & _
                    mTicker.Timestamp & "," & _
                    mProfitThisTrade & "," & _
                    mMaxProfitThisTrade & "," & _
                    mMinProfitThisTrade, _
                    IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedTradeProfile, TradeBuildListenValueTypes.VTTradeProfile), _
                    Me
    mProfitLastTrade = mProfitThisTrade
    mDrawdownLastTrade = mMaxProfitThisTrade - mProfitThisTrade
    mProfitThisTrade = 0
    mMaxProfitThisTrade = 0
    mMinProfitThisTrade = 0
    mProfitLastFlat = mProfit
End If

With mPlexTable(Handle)
    .closedOut = closedOut
    If .QuantityOutstanding = 0 Then .StartTime = mTicker.Timestamp
    
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
            calcPlexProfit Handle, mLastAsk
        End If
    ElseIf .QuantityOutstanding > 0 Then
        If mLastBid <> 0 Then
            calcPlexProfit Handle, mLastBid
        End If
    Else
        .endTime = mTicker.Timestamp
        calcPlexProfit Handle
        notifyPlexProfileListeners Handle
        
        ' Remove the relevant entry from the plex table
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
Dim size As Long


If Not mClosingPosition Then
    For Each OrderContext In mOrderContexts
        If Not OrderContext.Complete Then Exit Sub
    Next
    ' all OrderContexts are complete
    checkCloseall
    If mEndSession Then
        gTB.notifyListeners "Ending Session", _
                        TradeBuildListenValueTypes.VTMoneyManagement, _
                        Me
    End If
ElseIf pOrderContext Is mBalancingOrderContext Then
    'should now be net flat
    mOrderContexts.enabled = True
    
    ' now need to adjust the other orderPlex table
    ' entries' Profit figure to take account of the
    ' settling order
    
    Dim i As Long
    Dim settlingPrice As Double
    
    ' first Remove the settling entry from the plex table
    With mPlexTable(mSettlingHandle)
        mPlexTable(.PrevActiveIndex).NextActiveIndex = .NextActiveIndex
        mPlexTable(.NextActiveIndex).PrevActiveIndex = .PrevActiveIndex
    End With
    
    settlingPrice = mPlexTable(mSettlingHandle).TotalEntryPrice / mPlexTable(mSettlingHandle).QuantityOutstanding
    i = mPlexTable(mFirstActivePlexTableIndex).NextActiveIndex
    Do While i <> mLastActivePlexTableIndex
        With mPlexTable(i)
            .endTime = mTicker.Timestamp
            .ExitPrice = settlingPrice
            .closedOut = True
            .TotalEntryPrice = .TotalEntryPrice - (settlingPrice * .QuantityOutstanding)
            .QuantityOutstanding = 0
            calcPlexProfit i
            notifyPlexProfileListeners i
        
            i = .NextActiveIndex
            
            ' Remove the current entry from the plex table
            mPlexTable(.PrevActiveIndex).NextActiveIndex = .NextActiveIndex
            mPlexTable(.NextActiveIndex).PrevActiveIndex = .PrevActiveIndex
            
        End With
    Loop
    
    
    
    If mEndSession Then
        gTB.notifyListeners "Ending Session", _
                            TradeBuildListenValueTypes.VTMoneyManagement, _
                            Me
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
            gTB.notifyListeners "Ending Session", _
                                TradeBuildListenValueTypes.VTMoneyManagement, _
                                Me
        End If
        mClosingPosition = False
        fireChange PositionClosed
    Else
        placeBalancingOrder
    End If
End If
End Sub

Friend Function RegisterPlex(ByVal Key As String, _
                            ByVal Description As String, _
                            ByVal action As OrderActions, _
                            ByVal Risk As Currency) As Long
If mNextPlexTableIndex >= UBound(mPlexTable) Then
    ReDim Preserve mPlexTable(2 * (UBound(mPlexTable) + 1) - 1) As OrderPlexProfile
End If
RegisterPlex = mNextPlexTableIndex
mNextPlexTableIndex = mNextPlexTableIndex + 1
With mPlexTable(RegisterPlex)
    .Key = Key
    .Description = Description
    .action = action
    .Risk = Risk
    .NextActiveIndex = mLastActivePlexTableIndex
    .PrevActiveIndex = mPlexTable(mLastActivePlexTableIndex).PrevActiveIndex
    mPlexTable(.PrevActiveIndex).NextActiveIndex = RegisterPlex
    mPlexTable(mLastActivePlexTableIndex).PrevActiveIndex = RegisterPlex
End With
If Not mBalancingOrderPlex Is Nothing Then
    If Key = mBalancingOrderPlex.Key Then
        mSettlingHandle = RegisterPlex
    End If
End If
End Function

Public Sub RemoveChangeListener(ByVal value As ChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.item(i) Is value Then mChangeListeners.Remove i
Next
End Sub

Public Sub RemoveProfitListener(ByVal value As ProfitListener)
Dim i As Long
For i = mProfitListeners.Count To 1 Step -1
    If mProfitListeners.item(i) Is value Then mProfitListeners.Remove i
Next
End Sub

Friend Function SetPlexReason(ByVal Handle As Long, _
                            ByVal reasonType As ReasonTypes, _
                            ByVal reason As String)

With mPlexTable(Handle)
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

Friend Sub SetPlexRisk(ByVal Handle As Long, _
                            ByVal Risk As Currency)
mPlexTable(Handle).Risk = Risk
End Sub

Friend Sub StartOfSession(ByVal Timestamp As Date)
mProfitProfileBarStartTime = Timestamp
End Sub

Friend Sub TickfileCompleted()
Dim i As Long
i = mPlexTable(mFirstActivePlexTableIndex).NextActiveIndex
Do While i <> mLastActivePlexTableIndex
    With mPlexTable(i)
        .endTime = mTicker.Timestamp
        notifyPlexProfileListeners i
    
        i = .NextActiveIndex
        
        ' Remove the current entry from the plex table
        mPlexTable(.PrevActiveIndex).NextActiveIndex = .NextActiveIndex
        mPlexTable(.NextActiveIndex).PrevActiveIndex = .PrevActiveIndex
        
    End With
Loop

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcPlexProfit(ByVal Handle As Long, _
                            Optional ByVal price As Double)

With mPlexTable(Handle)
    If .QuantityOutstanding <> 0 Then
        .Profit = mTickValue * .QuantityOutstanding * (price - .TotalEntryPrice / .QuantityOutstanding) / mTickSize
        .ExitPrice = price
    Else
        .Profit = -1 * mTickValue * .TotalEntryPrice / mTickSize
        .ExitPrice = -(.TotalEntryPrice - .quantity * .EntryPrice) / .quantity
    End If
    
    If .Profit > .MaxProfit Then .MaxProfit = .Profit
    If .Profit < .MaxLoss Then .MaxLoss = .Profit
    
End With
End Sub

Private Sub calcTradeProfit(Optional ByVal price As Double)

Dim Profit As Currency
Dim ev As StudyValueEvent

If mQuantityOutstanding <> 0 Then
    Profit = mTickValue * mQuantityOutstanding * (price - mTotalEntryPrice / mQuantityOutstanding) / mTickSize
Else
    Profit = -1 * mTickValue * mTotalEntryPrice / mTickSize
End If

If Profit <> mProfit Then
    mProfit = Profit
    
    fireProfit mProfit, ProfitTypes.ProfitTypeSessionProfit
    
    Set ev.source = Me
    ev.sVal.Timestamp = mTicker.Timestamp
    ev.sVal.value = mProfit
    ev.valueName = "Profit"
    mStudies.notify ev
    
    gTB.notifyListeners mProfit, _
                    IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedProfit, TradeBuildListenValueTypes.VTProfit), _
                    Me
    gTB.notifyListeners mMaxProfit - mProfit, _
                    IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedDrawdown, TradeBuildListenValueTypes.VTDrawdown), _
                    Me
End If

If mProfit > mMaxProfit Then
    mMaxProfit = mProfit
    fireProfit mMaxProfit, ProfitTypes.ProfitTypeSessionMaxProfit
    gTB.notifyListeners mMaxProfit, _
                        IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedMaxProfit, TradeBuildListenValueTypes.VTMaxProfit), _
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
    ' start a new Profit profile bar
    gTB.notifyListeners mProfitProfileBarStartTime & "," & _
                        mProfitProfileBarNumber & "," & _
                        mProfitProfileBarOpen & "," & _
                        mProfitProfileBarHigh & "," & _
                        mProfitProfileBarLow & "," & _
                        mProfitProfileBarClose, _
                        IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedProfitProfile, TradeBuildListenValueTypes.VTProfitProfile), _
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

End Sub

Private Sub checkCloseall()

Dim msg As String
Dim endSession As Boolean

If mIsSimulated Then Exit Sub

If Not mMoneyManager Is Nothing Then
    If mMoneyManager.closeAll(mProfit, _
                                mMaxProfit - mProfit, _
                                mProfitThisTrade, _
                                mMaxProfitThisTrade - mProfitThisTrade, _
                                mTicker.Timestamp, _
                                msg, _
                                endSession) _
    Then
        If Not mClosingPosition Then
            gTB.notifyListeners "Close all positions: " & msg, _
                            TradeBuildListenValueTypes.VTMoneyManagement, _
                            Me
            ClosePositions
        End If
    Else
        If endSession And Not mClosingPosition Then
            If Not mEndSession Then
                mEndSession = True
                gTB.notifyListeners "Ending Session: " & msg, _
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
Dim Change As ChangeEvent
Set Change.source = Me
Change.changeType = changeType
For Each listener In mChangeListeners
    listener.Change Change
Next
End Sub

Private Sub fireProfit( _
                ByVal profitAmount As Double, _
                ByVal profitType As ProfitTypes)
Dim listener As ProfitListener
Dim Profit As ProfitEvent
Set Profit.source = Me
Profit.profitAmount = profitAmount
Profit.profitType = profitType
Profit.currencyCode = mcurrencyCode
For Each listener In mProfitListeners
    listener.profitAmount Profit
Next
End Sub

Private Sub notifyPlexProfileListeners(ByVal index As Long)
Dim data As Variant
data = mPlexTable(index)
gTB.notifyListeners _
                data, _
                IIf(mIsSimulated, _
                    TradeBuildListenValueTypes.VTSimulatedOrderPlexProfileStruct, _
                    TradeBuildListenValueTypes.VTOrderPlexProfileStruct), _
                Me

With mPlexTable(index)
    gTB.notifyListeners _
                FormatDateTime(.StartTime) & "," & _
                FormatDateTime(.endTime) & "," & _
                .Description & "," & _
                .Key & "," & _
                gOrderActionToString(.action) & "," & _
                .quantity & "," & _
                .EntryPrice & "," & _
                .ExitPrice & "," & _
                .entryReason & "," & _
                .targetReason & "," & _
                .stopReason & "," & _
                IIf(.closedOut, "1,", "0,") & _
                .Profit & "," & _
                .MaxProfit & "," & _
                .MaxLoss & "," & _
                .Risk & "," & _
                .QuantityOutstanding, _
                IIf(mIsSimulated, _
                    TradeBuildListenValueTypes.VTSimulatedOrderPlexProfileString, _
                    TradeBuildListenValueTypes.VTOrderPlexProfileString), _
                Me
End With
End Sub

Private Sub placeBalancingOrder()
gTB.notifyListeners _
        "Close positions: (" & mTicker.TradePriceString & ") at market", _
        IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
        Me

mBalancingOrderContext.enabled = True
If mBalancingOrderContext.isOrderTifValueSupported(TIFDay) Then mBalancingOrderContext.timeInForce = TIFDay
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
mBalancingOrderContext.executeOrderPlex mBalancingOrderPlex
End Sub

