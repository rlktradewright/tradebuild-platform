VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScaleStrategy1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements IScaleStrategy

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                            As String = "ScaleStrategy1"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mContext As TradingContext
Private mOrderContext As OrderContext

Private WithEvents mInitialBracketOrder As BracketOrder
Attribute mInitialBracketOrder.VB_VarHelpID = -1
Private WithEvents mAdditionalBracketOrder As BracketOrder   ' use this to catch the
Attribute mAdditionalBracketOrder.VB_VarHelpID = -1
                                                        ' fill event
Private mAdditionalBracketOrders As New Collection

Private S As ScaleStrategy1Static

Private mInitialEntryPrice As Double
Private mInitialATR As Double

Private mTickSize As Double
Private mTickValue As Double

Private mNumberOfIncrements As Long

Private mNextPriceThreshold As Double

Private mUnitSize As Long   ' number of contracts such that a price movement of 1 ATR (at time
                            ' of entry) represents mRiskUnitPercent of account balance
                            
Private mLongPosition As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Debug.Print "CScaleStrategy1 terminated"
End Sub

'================================================================================
' IScaleStrategy Interface Members
'================================================================================

Private Property Get IScaleStrategy_Name() As String
IScaleStrategy_Name = "ScaleStrategy1"
End Property

Private Sub IScaleStrategy_NotifyBracketOrderCompletion(ByVal BracketOrder As ResourceIdentifier)

End Sub

Private Sub IScaleStrategy_NotifyBracketOrderFill(ByVal BracketOrder As ResourceIdentifier)

End Sub

Private Sub IScaleStrategy_NotifyBracketOrderStopLossAdjusted(ByVal BracketOrder As ResourceIdentifier)

End Sub

Private Sub IScaleStrategy_NotifyTick(ByVal pType As TickTypes)
Const ProcName As String = "IScaleStrategy_NotifyTick"
On Error GoTo Err

Dim newValue As Double
Dim newStopOffsetTicks As Double
Dim stopIncrement As Double
Dim plex As IBracketOrder

If pType <> TickTypes.TickTypeTrade Then Exit Sub

newValue = mContext.TradeQuote.price

If mLongPosition Then
    If newValue < mNextPriceThreshold Then Exit Sub
    mNextPriceThreshold = 99999999
Else
    If newValue > mNextPriceThreshold Then Exit Sub
    mNextPriceThreshold = 0
End If



mNumberOfIncrements = mNumberOfIncrements + 1

' first change the stop loss for all existing BracketOrderes
stopIncrement = Int((S.StopIncrementFactor * mInitialATR / mTickSize) + 0.9999) * mTickSize

mInitialBracketOrder.AdjustStop mInitialBracketOrder.StopLossOrder.TriggerPrice + IIf(mLongPosition, stopIncrement, -stopIncrement)

On Error Resume Next
For Each plex In mAdditionalBracketOrders
    plex.AdjustStop plex.StopLossOrder.TriggerPrice + IIf(mLongPosition, stopIncrement, -stopIncrement)
Next
On Error GoTo 0

' now buy or sell the required amount
newStopOffsetTicks = Int((S.InitialStopFactor * mInitialATR / mTickSize) + 0.9999)
If mLongPosition Then
    Set mAdditionalBracketOrder = mOrderContext.Buy(Number:=mUnitSize, _
                    BracketEntryType:=BracketEntryTypeMarket, _
                    EntryPrice:=0, _
                    EntryOffset:=0, _
                    EntryTriggerPrice:=0, _
                    BracketStopLossType:=BracketStopLossTypeAuto, _
                    StopLossPrice:=0, _
                    StopLossOffset:=-newStopOffsetTicks, _
                    StopLossLimitPrice:=0, _
                    BracketTargetType:=BracketTargetTypeNone, _
                    TargetPrice:=0, _
                    TargetOffset:=0, _
                    TargetTriggerPrice:=0)
    If Not mAdditionalBracketOrder Is Nothing Then
        mAdditionalBracketOrders.Add mAdditionalBracketOrder
        mAdditionalBracketOrder.EntryReason = "A" & mNumberOfIncrements
    End If
Else
    Set mAdditionalBracketOrder = mOrderContext.Sell(Number:=mUnitSize, _
                    BracketEntryType:=BracketEntryTypeMarket, _
                    EntryPrice:=0, _
                    EntryOffset:=0, _
                    EntryTriggerPrice:=0, _
                    BracketStopLossType:=BracketStopLossTypeAuto, _
                    StopLossPrice:=0, _
                    StopLossOffset:=newStopOffsetTicks, _
                    StopLossLimitPrice:=0, _
                    BracketTargetType:=BracketTargetTypeNone, _
                    TargetPrice:=0, _
                    TargetOffset:=0, _
                    TargetTriggerPrice:=0)
    If Not mAdditionalBracketOrder Is Nothing Then
        mAdditionalBracketOrders.Add mAdditionalBracketOrder
        mAdditionalBracketOrder.EntryReason = "A" & mNumberOfIncrements
    End If
End If

If mNumberOfIncrements = S.MaxIncrements Then mContext.StopTickData Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IScaleStrategy_Start( _
                ByVal pContext As TradingContext, _
                ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "IScaleStrategy_Start"
On Error GoTo Err

Set mContext = pContext
Set mInitialBracketOrder = pBracketOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' mAdditionalBracketOrder Event Handlers
'================================================================================

Private Sub mAdditionalBracketOrder_EntryOrderPartiallyFilled()
Dim priceIncrement As Double

If mNumberOfIncrements = S.MaxIncrements Then mContext.StopTickData Me
priceIncrement = Int((S.ScaleThresholdFactor * mInitialATR / mTickSize) + 0.9999) * mTickSize
mNextPriceThreshold = mAdditionalBracketOrder.EntryOrder.AveragePrice + IIf(mLongPosition, priceIncrement, -priceIncrement)
End Sub

'================================================================================
' mInitialBracketOrder Event Handlers
'================================================================================

Private Sub mInitialBracketOrder_Completed()
Dim i As Long

If S.MaxIncrements = 0 Then Exit Sub     ' nothing to do

Set mInitialBracketOrder = Nothing
Set mAdditionalBracketOrder = Nothing
For i = mAdditionalBracketOrders.Count To 1 Step -1
    mAdditionalBracketOrders.Remove i
Next
mContext.StopTickData Me
End Sub

Private Sub mInitialBracketOrder_EntryOrderPartiallyFilled()
Dim priceIncrement As Double

Const ProcName As String = "mInitialBracketOrder_EntryOrderPartiallyFilled"
On Error GoTo Err

If S.MaxIncrements = 0 Then Exit Sub     ' nothing to do

mLongPosition = IIf(mInitialBracketOrder.EntryOrder.Action = OrderActionBuy, True, False)

mInitialEntryPrice = mInitialBracketOrder.EntryOrder.AveragePrice

mContext.StartTickData Me

mInitialATR = S.ATR.ATR(1).Value
mUnitSize = Int((S.RiskIncrementPercent * AccountBalance / 100) / (mTickValue * mInitialATR / mTickSize))
If mUnitSize = 0 Then mUnitSize = 1

priceIncrement = Int((S.ScaleThresholdFactor * mInitialATR / mTickSize) + 0.9999) * mTickSize
mNextPriceThreshold = mInitialEntryPrice + IIf(mLongPosition, priceIncrement, -priceIncrement)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mInitialBracketOrder_StopAdjusted()
Const ProcName As String = "mInitialBracketOrder_StopAdjusted"
On Error GoTo Err

Dim newStop As Double: newStop = mInitialBracketOrder.StopLossOrder.TriggerPrice

On Error Resume Next    ' it's possible some of the order plexes have already been stopped out

Dim plex As IBracketOrder
For Each plex In mAdditionalBracketOrders
    plex.AdjustStop newStop
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================




