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

Implements ChangeListener
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
Private mContractFuture                             As IFuture
Private mDataSource                                 As IMarketDataSource

Private mLastBid                                    As Double
Private mLastAsk                                    As Double

Private mCumBuyPrice                                As Double
Private mCumSellPrice                               As Double

Private mProfit                                     As Currency
Private mMaxProfit                                  As Currency
Private mMaxLoss                                    As Currency
Private mDrawdown                                   As Currency

Private mProfitListeners                            As New Listeners
Private mChangeListeners                            As New Listeners

Private mProfitNotificationThreshold                As Currency
Private mLossNotificationThreshold                  As Currency
Private mDrawdownNotificationThreshold              As Currency

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "ChangeListener_Change"
On Error GoTo Err

Select Case ev.changeType
Case BracketOrderChangeTypes.BracketOrderCompleted
    If mBracketOrder.Size = 0 Then
        logBracketProfile
        mDataSource.RemoveGenericTickListener Me
    End If
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IExecutionListener Interface Members
'@================================================================================

Private Sub IExecutionListener_NotifyExecution(ev As ExecutionEventData)
Const ProcName As String = "IExecutionListener_NotifyExecution"
On Error GoTo Err

If ev.ExecutionReport.Action = OrderActionBuy Then
    mCumBuyPrice = mCumBuyPrice + ev.ExecutionReport.Quantity * ev.ExecutionReport.Price
Else
    mCumSellPrice = mCumSellPrice + ev.ExecutionReport.Quantity * ev.ExecutionReport.Price
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
    ProcessBidPrice ev.Tick.Price
Case TickTypeAsk
    ProcessAskPrice ev.Tick.Price
Case TickTypeTrade
    ProcessTradePrice ev.Tick.Price
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get BracketOrder() As IBracketOrder
Set BracketOrder = mBracketOrder
End Property

Public Property Let DrawdownNotificationThreshold(ByVal Value As Currency)
mDrawdownNotificationThreshold = Value
End Property

Public Property Let LossNotificationThreshold(ByVal Value As Currency)
mLossNotificationThreshold = Value
End Property

Public Property Let ProfitNotificationThreshold(ByVal Value As Currency)
mProfitNotificationThreshold = Value
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddChangeListener(ByVal pListener As ChangeListener)
Const ProcName As String = "AddChangeListener"

On Error GoTo Err

mChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddProfitListener(ByVal pListener As IProfitListener)
Const ProcName As String = "AddProfitListener"

On Error GoTo Err

mProfitListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pDataSource As IMarketDataSource)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mBracketOrder = pBracketOrder
mBracketOrder.AddExecutionListener Me
mBracketOrder.AddChangeListener Me

mCumBuyPrice = pBracketOrder.CumBuyPrice
mCumSellPrice = pBracketOrder.CumSellPrice

Set mContractFuture = pDataSource.ContractFuture

Set mDataSource = pDataSource
mDataSource.AddGenericTickListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveChangeListener(ByVal pListener As ChangeListener)
Const ProcName As String = "RemoveChangeListener"

On Error GoTo Err

mChangeListeners.Remove pListener

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

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcProfitLoss(ByVal Value As Double)
Const ProcName As String = "calcProfitLoss"
On Error GoTo Err

Static sValueFactor As Double

If sValueFactor = 0 Then
    Assert mContractFuture.IsAvailable
    
    Dim lContract As IContract
    Set lContract = mContractFuture.Value
    sValueFactor = lContract.Multiplier
End If

If Value = 0 Then Exit Function

Dim lProfit As Currency
lProfit = (mCumSellPrice + mBracketOrder.Size * Value - mCumBuyPrice) * sValueFactor

Dim lNotifyProfit As Boolean
If lProfit <> mProfit Then
    mProfit = lProfit
    lNotifyProfit = True
End If

Dim lNotifyMaxProfit As Boolean
If lProfit > mMaxProfit Then
    mMaxProfit = lProfit
    lNotifyMaxProfit = True
End If

Dim lNotifyMaxLoss As Boolean
If lProfit < mMaxLoss Then
    mMaxLoss = lProfit
    lNotifyMaxLoss = True
End If

Dim lNotifyDrawdown As Boolean
If lProfit < mMaxProfit Then
    mDrawdown = mMaxProfit - lProfit
    lNotifyDrawdown = True
ElseIf mDrawdown <> 0 Then
    mDrawdown = 0
    lNotifyDrawdown = True
End If

If lNotifyProfit Then fireProfit mProfit, ProfitTypes.ProfitTypeProfit
If lNotifyMaxProfit Then fireProfit mMaxProfit, ProfitTypes.ProfitTypeMaxProfit
If lNotifyMaxLoss Then fireProfit mMaxLoss, ProfitTypes.ProfitTypeMaxLoss
If lNotifyDrawdown Then fireProfit mDrawdown, ProfitTypes.ProfitTypeDrawdown

If lNotifyProfit Then
    If mProfitNotificationThreshold <> 0 And _
        mProfit >= mProfitNotificationThreshold _
    Then
        fireChange ProfitChangeProfitThresholdExceeded
        mProfitNotificationThreshold = 0
    End If
    
    If mLossNotificationThreshold <> 0 And _
        mProfit <= mLossNotificationThreshold _
    Then
        fireChange ProfitChangeLossThresholdExceeded
        mLossNotificationThreshold = 0
    End If

    If mDrawdownNotificationThreshold <> 0 And _
        (mMaxProfit - lProfit >= mDrawdownNotificationThreshold) _
    Then
        fireChange ProfitChangeDrawdownThresholdExceeded
        mDrawdownNotificationThreshold = 0
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub fireChange( _
                ByVal changeType As ProfitChangeTypes)
Const ProcName As String = "fireChange"
On Error GoTo Err

Dim ev As ChangeEventData
Set ev.Source = Me
ev.changeType = changeType

mChangeListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mChangeListeners.Count
    Dim lListener As ChangeListener
    Set lListener = mChangeListeners.GetListener(i)
    lListener.Change ev
Next

RaiseEvent Change(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireProfit( _
                ByVal ProfitAmount As Double, _
                ByVal ProfitType As ProfitTypes)
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
ev.ProfitAmount = ProfitAmount
ev.ProfitType = ProfitType
ev.CurrencyCode = sCurrencyCode

mProfitListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mProfitListeners.Count
    Dim lListener As IProfitListener
    Set lListener = mProfitListeners.GetListener(i)
    lListener.NotifyProfit ev
Next

RaiseEvent Profit(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub logBracketProfile()
Const ProcName As String = "logBracketProfile"
On Error GoTo Err

Dim profile As BracketOrderProfile
With profile
    .Action = mBracketOrder.EntryOrder.Action
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
End With
gLogBracketOrderProfileStruct profile, mBracketOrder.IsSimulated, Me

gLogBracketOrderProfileString FormatDateTime(mBracketOrder.StartTime) & "," & _
            FormatDateTime(mBracketOrder.EndTime) & "," & _
            mBracketOrder.Description & "," & _
            mBracketOrder.Key & "," & _
            gOrderActionToString(mBracketOrder.EntryOrder.Action) & "," & _
            mBracketOrder.EntryOrder.Quantity & "," & _
            mBracketOrder.EntryPrice & "," & _
            mBracketOrder.ExitPrice & "," & _
            mBracketOrder.EntryReason & "," & _
            mBracketOrder.TargetReason & "," & _
            mBracketOrder.StopReason & "," & _
            mProfit & "," & _
            mMaxProfit & "," & _
            mMaxLoss & "," & _
            mBracketOrder.Risk, _
            mBracketOrder.IsSimulated, _
            Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ProcessAskPrice(ByVal pPrice As Double)
Const ProcName As String = "ProcessAskPrice"
On Error GoTo Err

If pPrice = mLastAsk Then Exit Sub
If Not mBracketOrder.LongPosition Then calcProfitLoss pPrice
mLastAsk = pPrice

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ProcessBidPrice(ByVal pPrice As Double)
Const ProcName As String = "ProcessBidPrice"
On Error GoTo Err

If pPrice = mLastBid Then Exit Sub
If mBracketOrder.LongPosition Then calcProfitLoss pPrice
mLastBid = pPrice

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ProcessTradePrice(ByVal pPrice As Double)
Const ProcName As String = "ProcessTradePrice"
On Error GoTo Err

If pPrice > mLastAsk Then ProcessAskPrice pPrice     ' because the Ask Price must have gone at least that high
If pPrice < mLastBid Then ProcessBidPrice pPrice   ' because the Bid Price must have gone at least that low

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




