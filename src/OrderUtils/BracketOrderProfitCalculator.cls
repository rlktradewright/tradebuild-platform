VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketOrderProfitCalculator"
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

Implements IGenericTickListener

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

Private Const ModuleName                            As String = "BracketOrderProfitCalculator"

'@================================================================================
' Member variables
'@================================================================================

Private mBracketOrder                               As BracketOrder
Private mDataSource                                 As IMarketDataSource

Private mLastBid                                    As Double
Private mLastAsk                                    As Double

Private mValueFactor                                As Double
Private mProfit                                     As Currency
Private mMaxProfit                                  As Currency
Private mMaxLoss                                    As Currency
Private mDrawdown                                   As Currency
Private mProfitNotificationThreshold                As Currency
Private mLossNotificationThreshold                  As Currency
Private mDrawdownNotificationThreshold              As Currency


'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Select Case ev.Tick.TickType
Case TickTypeBid
    processBidTick ev.Tick
Case TickTypeAsk
    processAskTick ev.Tick
Case TickTypeTrade
    If ev.price > mLastAsk Then processAskTick ev.Tick  ' because the Ask price must have gone at least that high
    If ev.price < mLastBid Then processBidTick ev.Tick  ' because the Bid price must have gone at least that low
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

mDataSource.RemoveGenericTickListener Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Initialise(ByVal pBracketOrder As BracketOrder, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pTickSize As Double, _
                ByVal pTickValue As Double)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mBracketOrder = pBracketOrder
Set mDataSource = pDataSource
mDataSource.AddGenericTickListener Me
mValueFactor = lContract.Multiplier


Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcProfitLoss(ByVal value As Double)
Const ProcName As String = "calcProfitLoss"
Dim lProfit As Currency
Dim lNotifyProfit As Boolean
Dim lNotifyMaxProfit As Boolean
Dim lNotifyMaxLoss As Boolean
Dim lNotifyDrawdown As Boolean


On Error GoTo Err

If value = 0 Then Exit Function

lProfit = (mCumSellPrice + size * value - mCumBuyPrice) * mValueFactor

If lProfit <> mProfit Then
    mProfit = lProfit
    lNotifyProfit = True
End If

If lProfit > mMaxProfit Then
    mMaxProfit = lProfit
    lNotifyMaxProfit = True
End If

If lProfit < mMaxLoss Then
    mMaxLoss = lProfit
    lNotifyMaxLoss = True
End If

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
        fireChange BracketOrderProfitThresholdExceeded
        RaiseEvent ProfitThresholdExceeded
        mProfitNotificationThreshold = 0
        mSummary.ProfitNotificationThreshold = 0
    End If
    
    If mLossNotificationThreshold <> 0 And _
        mProfit <= mLossNotificationThreshold _
    Then
        fireChange BracketOrderLossThresholdExceeded
        RaiseEvent LossThresholdExceeded
        mLossNotificationThreshold = 0
        mSummary.LossNotificationThreshold = 0
    End If

    If mDrawdownNotificationThreshold <> 0 And _
        (mMaxProfit - lProfit >= mDrawdownNotificationThreshold) _
    Then
        fireChange BracketOrderDrawdownThresholdExceeded
        RaiseEvent DrawdownThresholdExceeded
        mDrawdownNotificationThreshold = 0
        mSummary.DrawdownNotificationThreshold = 0
    End If
End If

SaveRecoveryInfo
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub processAskTick(ByRef pTick As GenericTick)
Const ProcName As String = "processAskTick"
On Error GoTo Err

If pTick.price = mLastAsk Then Exit Sub
If Not mBracketOrder.LongPosition Then calcProfitLoss pTick.price
mLastAsk = pTick.price

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub processBidTick(ByRef pTick As GenericTick)
Const ProcName As String = "processBidTick"
On Error GoTo Err

If pTick.price = mLastBid Then Exit Sub
If mBracketOrder.LongPosition Then calcProfitLoss pTick.price
mLastBid = pTick.price

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub



