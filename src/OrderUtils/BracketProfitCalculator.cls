VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketProfitCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
Implements IGenericTickListener
Implements IExecutionListener

'@================================================================================
' Events
'@================================================================================

Event Change(ev As ChangeEventData)
Event Profit(ev As ProfitEventData)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "BracketProfitCalculator"

'@================================================================================
' Member variables
'@================================================================================

Private mBracketOrder                               As IBracketOrder
Attribute mBracketOrder.VB_VarHelpID = -1
Private mBracketOrderKey                            As String
Private mContractFuture                             As IFuture
Private mDataSource                                 As IMarketDataSource

Private mLastBid                                    As Double
Private mLastAsk                                    As Double
Private mLastTrade                                  As Double

Private mCumBuyPrice                                As BoxedDecimal
Private mCumSellPrice                               As BoxedDecimal

Private mProfit                                     As BoxedDecimal
Private mMaxProfit                                  As BoxedDecimal
Private mMaxLoss                                    As BoxedDecimal
Private mDrawdown                                   As BoxedDecimal

Private mProfitListeners                            As New Listeners
Private mChangeListeners                            As New Listeners

Private mProfitNotificationThreshold                As BoxedDecimal
Private mLossNotificationThreshold                  As BoxedDecimal
Private mDrawdownNotificationThreshold              As BoxedDecimal

Private mAnExecutionHasOccurred                     As Boolean

Private mCurrentTick                                As GenericTick

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mCumBuyPrice = DecimalZero
Set mCumSellPrice = DecimalZero

Set mProfit = DecimalZero
Set mMaxProfit = DecimalZero
Set mMaxLoss = DecimalZero
Set mDrawdown = DecimalZero

Set mProfitNotificationThreshold = DecimalZero
Set mLossNotificationThreshold = DecimalZero
Set mDrawdownNotificationThreshold = DecimalZero

Set mCurrentTick.Size = DecimalZero
End Sub

Private Sub Class_Terminate()
Const ProcName As String = "Class_Terminate"
On Error GoTo Err

GOrderUtils.Log "Profit calculator: " & GetObjectKey(Me) & _
        " terminated for bracket order: " & mBracketOrderKey, _
    ProcName, ModuleName, , LogLevelDetail

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
On Error GoTo Err

Select Case ev.changeType
Case BracketOrderChangeTypes.BracketOrderSelfCancelled
    mBracketOrder.RemoveExecutionListener Me
    mBracketOrder.RemoveChangeListener Me
    Set mBracketOrder = Nothing
    mDataSource.RemoveGenericTickListener Me
Case BracketOrderChangeTypes.BracketOrderCompleted
    If mAnExecutionHasOccurred And mBracketOrder.Size = 0 Then
        logBracketProfile
        mDataSource.RemoveGenericTickListener Me
    End If
Case BracketOrderChangeTypes.BracketOrderPositionCloseoutNotified
    Assert mBracketOrder.Size = 0, "Bracket order size is not zero after position closeout notification"
    
    logBracketProfile
    
    mBracketOrder.RemoveExecutionListener Me
    mBracketOrder.RemoveChangeListener Me
    Set mBracketOrder = Nothing
    mDataSource.RemoveGenericTickListener Me
End Select

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IExecutionListener Interface Members
'@================================================================================

Private Sub IExecutionListener_NotifyExecution(ev As ExecutionEventData)
Const ProcName As String = "IExecutionListener_NotifyExecution"
On Error GoTo Err

If Not mAnExecutionHasOccurred Then
    mAnExecutionHasOccurred = True
    listenToData
End If

If ev.ExecutionReport.Action = OrderActionBuy Then
    Set mCumBuyPrice = mCumBuyPrice.Add(ev.ExecutionReport.Quantity.Mult(CreateBoxedDecimal(ev.ExecutionReport.Price)))
Else
    Set mCumSellPrice = mCumSellPrice.Add(ev.ExecutionReport.Quantity.Mult(CreateBoxedDecimal(ev.ExecutionReport.Price)))
End If

calcProfitUsingLastReceivedPrices

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

If mBracketOrder Is Nothing Then Exit Sub

mCurrentTick = ev.Tick

Select Case ev.Tick.TickType
Case TickTypeBid
    If ev.Tick.Price = mLastBid Then Exit Sub
    ProcessBidPrice ev.Tick.Price
Case TickTypeAsk
    If ev.Tick.Price = mLastAsk Then Exit Sub
    ProcessAskPrice ev.Tick.Price
Case TickTypeTrade
    If ev.Tick.Price = mLastTrade Then Exit Sub
    ProcessTradePrice ev.Tick.Price
End Select

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get BracketOrder() As IBracketOrder
Set BracketOrder = mBracketOrder
End Property

Public Property Get Drawdown() As BoxedDecimal
Set Drawdown = mDrawdown
End Property

Public Property Let DrawdownNotificationThreshold(ByVal Value As BoxedDecimal)
Set mDrawdownNotificationThreshold = Value
End Property

Public Property Let LossNotificationThreshold(ByVal Value As BoxedDecimal)
Set mLossNotificationThreshold = Value
End Property

Public Property Get MaxLoss() As BoxedDecimal
Set MaxLoss = mMaxLoss
End Property

Public Property Get MaxProfit() As BoxedDecimal
Set MaxProfit = mMaxProfit
End Property

Public Property Get Profit() As BoxedDecimal
Set Profit = mProfit
End Property

Public Property Let ProfitNotificationThreshold(ByVal Value As BoxedDecimal)
Set mProfitNotificationThreshold = Value
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddProfitListener(ByVal pListener As IProfitListener)
Const ProcName As String = "AddProfitListener"
On Error GoTo Err

mProfitListeners.Add pListener

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pDataSource As IMarketDataSource)
Const ProcName As String = "Initialise"
On Error GoTo Err

mLastTrade = MaxDouble

Set mBracketOrder = pBracketOrder
mBracketOrderKey = mBracketOrder.Key
mBracketOrder.AddExecutionListener Me
mBracketOrder.AddChangeListener Me

GOrderUtils.Log "Profit calculator: " & GetObjectKey(Me) & _
        " initialised for bracket order: " & mBracketOrderKey, _
    ProcName, ModuleName, , LogLevelDetail

Set mContractFuture = pDataSource.ContractFuture
Set mDataSource = pDataSource

'If mBracketOrder.Size <> 0 Then
    Set mCumBuyPrice = pBracketOrder.CumBuyPrice
    Set mCumSellPrice = pBracketOrder.CumSellPrice
    If mCumBuyPrice <> 0 Or mCumSellPrice <> 0 Then mAnExecutionHasOccurred = True

    listenToData
    calcInitialProfit
'End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveChangeListener(ByVal pListener As IChangeListener)
Const ProcName As String = "RemoveChangeListener"
On Error GoTo Err

mChangeListeners.Remove pListener

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveProfitListener(ByVal pListener As IProfitListener)
Const ProcName As String = "RemoveProfitListener"
On Error GoTo Err

mProfitListeners.Remove pListener

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcInitialProfit()
Const ProcName As String = "calcInitialProfit"
On Error GoTo Err

If mDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
If Not mDataSource.HasCurrentTick(TickTypeAsk) Then Exit Sub
If Not mDataSource.HasCurrentTick(TickTypeBid) Then Exit Sub
If Not mDataSource.HasCurrentTick(TickTypeTrade) Then Exit Sub

Dim lAskTick As GenericTick
lAskTick = mDataSource.CurrentTick(TickTypeAsk)
If lAskTick.Timestamp <> 0# Then mLastAsk = lAskTick.Price

Dim lBidTick As GenericTick
lBidTick = mDataSource.CurrentTick(TickTypeBid)
If lBidTick.Timestamp <> 0# Then mLastBid = lBidTick.Price

Dim lTradeTick As GenericTick
lTradeTick = mDataSource.CurrentTick(TickTypeTrade)

mLastTrade = MaxDouble

If mLastBid <> 0# And _
    lTradeTick.Timestamp <> 0# And _
    lTradeTick.Timestamp >= lBidTick.Timestamp And _
    lTradeTick.Price < mLastBid Then
    mLastTrade = lTradeTick.Price
End If

If mLastAsk <> 0# And _
    lTradeTick.Timestamp <> 0# And _
    lTradeTick.Timestamp >= lAskTick.Timestamp And _
    lTradeTick.Price > mLastAsk Then
    mLastTrade = lTradeTick.Price
End If

calcProfitUsingLastReceivedPrices

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub calcProfitLoss(ByVal pPrice As Double)
Const ProcName As String = "calcProfitLoss"
On Error GoTo Err

Static sValueFactor As BoxedDecimal

If sValueFactor Is Nothing Then
    Assert mContractFuture.IsAvailable
    
    Dim lContract As IContract
    Set lContract = mContractFuture.Value
    Set sValueFactor = CreateBoxedDecimal(lContract.Specifier.Multiplier)
    If sValueFactor Is DecimalZero Then Set sValueFactor = DecimalOne
End If

If pPrice = 0 Then Exit Sub

Dim lProfit As BoxedDecimal
Set lProfit = mCumSellPrice.Add(mBracketOrder.Size.Mult(CreateBoxedDecimal(pPrice)).Subtr(mCumBuyPrice)).Mult(sValueFactor)

Dim lNotifyProfit As Boolean
If lProfit.NE(mProfit) Then
    Set mProfit = lProfit
    lNotifyProfit = True
End If

Dim lNotifyMaxProfit As Boolean
If lProfit.GT(mMaxProfit) Then
    Set mMaxProfit = lProfit
    lNotifyMaxProfit = True
    GOrderUtils.Log mBracketOrder.Key & ": Max profit increased: " & mMaxProfit.ToString & "; Tick=" & GenericTickToString(mCurrentTick), ProcName, ModuleName, , LogLevels.LogLevelDetail
End If

Dim lNotifyMaxLoss As Boolean
If lProfit.LT(mMaxLoss) Then
    Set mMaxLoss = lProfit
    lNotifyMaxLoss = True
End If

Dim lNotifyDrawdown As Boolean
If lProfit.LT(mMaxProfit) Then
    Set mDrawdown = mMaxProfit.Subtr(lProfit)
    lNotifyDrawdown = True
ElseIf mDrawdown <> 0 Then
    Set mDrawdown = DecimalZero
    lNotifyDrawdown = True
End If

Dim lProfitTypes As ProfitTypes
If lNotifyProfit Then lProfitTypes = ProfitTypes.ProfitTypeProfit
If lNotifyMaxProfit Then lProfitTypes = lProfitTypes Or ProfitTypes.ProfitTypeMaxProfit
If lNotifyMaxLoss Then lProfitTypes = lProfitTypes Or ProfitTypes.ProfitTypeMaxLoss
If lNotifyDrawdown Then lProfitTypes = lProfitTypes Or ProfitTypes.ProfitTypeDrawdown
fireProfit lProfitTypes

If lNotifyProfit Then
    If Not mProfitNotificationThreshold Is DecimalZero And _
        mProfit.GE(mProfitNotificationThreshold) _
    Then
        fireChange ProfitChangeProfitThresholdExceeded
        Set mProfitNotificationThreshold = DecimalZero
    End If
    
    If Not mLossNotificationThreshold Is DecimalZero And _
        mProfit.LE(mLossNotificationThreshold) _
    Then
        fireChange ProfitChangeLossThresholdExceeded
        Set mLossNotificationThreshold = DecimalZero
    End If

    If Not mDrawdownNotificationThreshold Is DecimalZero And _
        (mMaxProfit.Subtr(lProfit).GE(mDrawdownNotificationThreshold)) _
    Then
        fireChange ProfitChangeDrawdownThresholdExceeded
        Set mDrawdownNotificationThreshold = DecimalZero
    End If
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub calcProfitUsingLastReceivedPrices()
Const ProcName As String = "calcProfitUsingLastReceivedPrices"
On Error GoTo Err

If mLastTrade <> MaxDouble Then
    ProcessTradePrice mLastTrade
ElseIf mBracketOrder.LongPosition Then
    ProcessBidPrice mLastBid
Else
    ProcessAskPrice mLastAsk
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireChange( _
                ByVal changeType As ProfitChangeTypes)
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireProfit( _
                ByVal pProfitTypes As ProfitTypes)
Const ProcName As String = "fireProfit"
On Error GoTo Err

Static sCurrencyCode As String
Dim ev As ProfitEventData

If sCurrencyCode = "" Then
    Assert mContractFuture.IsAvailable
    Dim lContract As IContract
    Set lContract = mContractFuture.Value
    sCurrencyCode = lContract.Specifier.CurrencyCode
End If

Set ev.Source = Me
ev.ProfitTypes = pProfitTypes
ev.CurrencyCode = sCurrencyCode

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

RaiseEvent Profit(ev)

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub listenToData()
Const ProcName As String = "listenToData"
On Error GoTo Err

mDataSource.AddGenericTickListener Me

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub logBracketProfile()
Const ProcName As String = "logBracketProfile"
On Error GoTo Err

Static sContractSpec As IContractSpecifier
Static sContractTimezoneName As String

If sContractSpec Is Nothing Then
    Dim lContract As IContract
    Set lContract = mContractFuture.Value
    Set sContractSpec = lContract.Specifier
    sContractTimezoneName = lContract.TimezoneName
End If

Dim profile As New BracketOrderProfile
With profile
    .Action = mBracketOrder.EntryOrder.Action
    .ContractSpec = sContractSpec
    .Description = mBracketOrder.Description
    .EndTime = mBracketOrder.EndTime
    .EntryReason = mBracketOrder.EntryReason
    .EntryPrice = mBracketOrder.EntryPrice
    .ExitPrice = mBracketOrder.ExitPrice
    .Key = mBracketOrder.Key
    .MaxLoss = mMaxLoss
    .MaxProfit = mMaxProfit
    .Profit = mProfit
    .Quantity = mBracketOrder.EntryOrder.Quantity
    .Risk = mBracketOrder.Risk
    .StartTime = mBracketOrder.StartTime
    .StopReason = mBracketOrder.StopReason
    .TargetReason = mBracketOrder.TargetReason
    .TimezoneName = sContractTimezoneName
End With
GOrderUtils.LogBracketOrderProfileObject profile, mBracketOrder.IsSimulated, Me

GOrderUtils.LogBracketOrderProfileString _
            mBracketOrder.GroupName & "," & _
            GOrderUtils.GetContractName(sContractSpec) & "," & _
            FormatTimestamp(mBracketOrder.StartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & "," & _
            FormatTimestamp(mBracketOrder.EndTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & "," & _
            sContractTimezoneName & "," & _
            mBracketOrder.Description & "," & _
            mBracketOrder.Key & "," & _
            OrderActionToString(mBracketOrder.EntryOrder.Action) & "," & _
            mBracketOrder.EntryOrder.Quantity & "," & _
            mBracketOrder.EntryPrice & "," & _
            mBracketOrder.ExitPrice & "," & _
            mBracketOrder.EntryReason & "," & _
            mBracketOrder.TargetReason & "," & _
            mBracketOrder.StopReason & "," & _
            mProfit & "," & _
            mMaxProfit & "," & _
            mMaxLoss & "," & _
            IIf(mBracketOrder.Risk = MaxCurrency, "Max", mBracketOrder.Risk), _
            mBracketOrder.IsSimulated, _
            Me

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ProcessAskPrice(ByVal pPrice As Double)
Const ProcName As String = "ProcessAskPrice"
On Error GoTo Err

If Not mBracketOrder.LongPosition Then calcProfitLoss pPrice
mLastAsk = pPrice
mLastTrade = MaxDouble

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ProcessBidPrice(ByVal pPrice As Double)
Const ProcName As String = "ProcessBidPrice"
On Error GoTo Err

If mBracketOrder.LongPosition Then calcProfitLoss pPrice
mLastBid = pPrice
mLastTrade = MaxDouble

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ProcessTradePrice(ByVal pPrice As Double)
Const ProcName As String = "ProcessTradePrice"
On Error GoTo Err

If pPrice > mLastAsk Then
    ProcessAskPrice pPrice     ' because the Ask Price must have gone at least that high
    mLastTrade = pPrice
ElseIf pPrice < mLastBid Then
    ProcessBidPrice pPrice   ' because the Bid Price must have gone at least that low
    mLastTrade = pPrice
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub




