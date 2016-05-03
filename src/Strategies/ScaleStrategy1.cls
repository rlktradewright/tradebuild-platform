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

Implements IBracketOrderEventSink
Implements IPositionMgmtStrategy

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
' IBracketOrderEventSink Interface Members
'================================================================================

Private Sub IBracketOrderEventSink_NotifyBracketOrderCompletion(ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "IBracketOrderEventSink_NotifyBracketOrderCompletion"
On Error GoTo Err

If Not BracketOrder Is PrimaryBracketOrder Then Exit Sub
If S.MaxIncrements = 0 Then Exit Sub     ' nothing to do

ClearPrimaryBracketOrder
Set mAdditionalBracketOrder = Nothing
Set mAdditionalBracketOrders = Nothing
mContext.StopTickData Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrderEventSink_NotifyBracketOrderFill(ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "IBracketOrderEventSink_NotifyBracketOrderFill"
On Error GoTo Err

If S.MaxIncrements = 0 Then Exit Sub     ' nothing to do

Dim priceIncrement As Double

If BracketOrder Is PrimaryBracketOrder Then
    mLongPosition = IsBuy
    
    mInitialEntryPrice = AveragePrice
    
    mContext.StartTickData Me
    
    mInitialATR = mContext.GetStudyValue(S.ATR)
    mUnitSize = Int((S.RiskIncrementPercent * AccountBalance / 100) / (mTickValue * mInitialATR / mTickSize))
    If mUnitSize = 0 Then mUnitSize = 1
    
    priceIncrement = Int((S.ScaleThresholdFactor * mInitialATR / mTickSize) + 0.9999) * mTickSize
    mNextPriceThreshold = mInitialEntryPrice + IIf(mLongPosition, priceIncrement, -priceIncrement)
Else
    If mNumberOfIncrements = S.MaxIncrements Then mContext.StopTickData Me
    priceIncrement = Int((S.ScaleThresholdFactor * mInitialATR / mTickSize) + 0.9999) * mTickSize
    mNextPriceThreshold = AveragePrice(mAdditionalBracketOrder) + IIf(mLongPosition, priceIncrement, -priceIncrement)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrderEventSink_NotifyBracketOrderStopLossAdjusted(ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "IBracketOrderEventSink_NotifyBracketOrderStopLossAdjusted"
On Error GoTo Err

If Not CanTrade Then Exit Sub

If Not BracketOrder Is PrimaryBracketOrder Then Exit Sub

Dim bo As ResourceIdentifier
For Each bo In mAdditionalBracketOrders
    If Not IsCompleted(bo) Then AdjustStop StopLossTriggerPrice, , bo
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' IPositionMgmtStrategy Interface Members
'================================================================================

Private Property Get IPositionMgmtStrategy_Name() As String
IPositionMgmtStrategy_Name = "ScaleStrategy1"
End Property

Private Sub IPositionMgmtStrategy_NotifyTick(ByVal pType As TickTypes)
Const ProcName As String = "IPositionMgmtStrategy_NotifyTick"
On Error GoTo Err

If pType <> TickTypes.TickTypeTrade Then Exit Sub
If Not CanTrade Then Exit Sub

Dim newValue As Double
newValue = mContext.TradeQuote.price

If mLongPosition Then
    If newValue < mNextPriceThreshold Then Exit Sub
    mNextPriceThreshold = 99999999
Else
    If newValue > mNextPriceThreshold Then Exit Sub
    mNextPriceThreshold = 0
End If



mNumberOfIncrements = mNumberOfIncrements + 1

' first change the stop loss for all existing BracketOrders
Dim stopIncrement As Double
stopIncrement = Int((S.StopIncrementFactor * mInitialATR / mTickSize) + 0.9999) * mTickSize

AdjustStop StopLossTriggerPrice + IIf(mLongPosition, stopIncrement, -stopIncrement)

Dim bo As ResourceIdentifier
For Each bo In mAdditionalBracketOrders
    If Not IsCompleted(bo) Then AdjustStop StopLossTriggerPrice(bo) + IIf(mLongPosition, stopIncrement, -stopIncrement), , bo
Next
On Error GoTo 0

' now buy or sell the required amount
Dim newStopOffsetTicks As Double
newStopOffsetTicks = Int((S.InitialStopFactor * mInitialATR / mTickSize) + 0.9999)
If mLongPosition Then
    Set mAdditionalBracketOrder = Buy( _
                    Number:=mUnitSize, _
                    EntryOrder:=DeclareEntryOrder(EntryOrderTypeMarket), _
                    StopLossOrder:=DeclareStopLossOrder(StopLossOrderTypeAuto, Offset:=-newStopOffsetTicks), _
                    NotifyCompletion:=True)
    If Not mAdditionalBracketOrder Is Nothing Then
        mAdditionalBracketOrders.Add mAdditionalBracketOrder
        SetEntryReason "A" & mNumberOfIncrements, mAdditionalBracketOrder
    End If
Else
    Set mAdditionalBracketOrder = Sell( _
                    Number:=mUnitSize, _
                    EntryOrder:=DeclareEntryOrder(EntryOrderTypeMarket), _
                    StopLossOrder:=DeclareStopLossOrder(StopLossOrderTypeAuto, Offset:=newStopOffsetTicks), _
                    NotifyCompletion:=True)
    If Not mAdditionalBracketOrder Is Nothing Then
        mAdditionalBracketOrders.Add mAdditionalBracketOrder
        SetEntryReason "A" & mNumberOfIncrements, mAdditionalBracketOrder
    End If
End If

If mNumberOfIncrements = S.MaxIncrements Then mContext.StopTickData Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IPositionMgmtStrategy_Start( _
                ByVal TradingContext As TradingContext, _
                ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "IPositionMgmtStrategy_Start"
On Error GoTo Err

Set mContext = TradingContext
SetPrimaryBracketOrder BracketOrder

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




