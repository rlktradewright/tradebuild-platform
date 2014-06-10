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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

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
Private mOrderContext As OrderContext

Private WithEvents mInitialBracketOrder As BracketOrder
Attribute mInitialBracketOrder.VB_VarHelpID = -1
Private WithEvents mAdditionalBracketOrder As BracketOrder   ' use this to catch the
Attribute mAdditionalBracketOrder.VB_VarHelpID = -1
                                                        ' fill event
Private mAdditionalBracketOrders As Collection

Private mInitialEntryPrice As Double
Private mInitialATR As Double

Private mTimePeriod As Integer
Private mTimePeriodUnit As String
Private mBars As Bars
Private mATRStudy As ATR
Private mATRPeriods As Long
Private mListenerKey As String

Private mTickSize As Double
Private mTickValue As Double

Private mNumberOfIncrements As Long
Private mMaxIncrements As Long
Private mRiskUnitPercent As Double          ' the percentage of account balance when 1 unit of contracts
                                            ' changes price by 1 ATR
Private mRiskIncrementPercent As Double     ' the amount by which risk is to be increased at each scaling point
Private mScaleThresholdFactor As Double     ' the number of ATRs to the next scaling point
Private mStopIncrementFactor As Double      ' the number of ATRs to raise the stop losses at each scaling point
Private mInitialStopFactor As Double        ' the number of ATR to set the initial stop loss for each order

Private mNextPriceThreshold As Double

Private mUnitSize As Long   ' number of contracts such that a price movement of 1 ATR (at time
                            ' of entry) represents mRiskUnitPercent of account balance
                            
Private mLongPosition As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mAdditionalBracketOrders = New Collection
mATRPeriods = 20
mMaxIncrements = 3
mRiskUnitPercent = 1#
mRiskIncrementPercent = 1#
mScaleThresholdFactor = 0.5
mStopIncrementFactor = 0.5
mInitialStopFactor = 2#
End Sub

Private Sub Class_Terminate()
Debug.Print "CScaleStrategy1 terminated"
End Sub

'================================================================================
' IPositionMgmtStrategy Interface Members
'================================================================================

Private Sub IPositionMgmtStrategy_NotifyTick(ByVal pType As TickTypes)
Const ProcName As String = "IPositionMgmtStrategy_NotifyTick"
On Error GoTo Err

Dim newValue As Double
Dim newStopOffsetTicks As Double
Dim stopIncrement As Double
Dim plex As IBracketOrder

newValue = newval

If mLongPosition Then
    If newValue < mNextPriceThreshold Then Exit Sub
    mNextPriceThreshold = 99999999
Else
    If newValue > mNextPriceThreshold Then Exit Sub
    mNextPriceThreshold = 0
End If



mNumberOfIncrements = mNumberOfIncrements + 1

' first change the stop loss for all existing BracketOrderes
stopIncrement = Int((mStopIncrementFactor * mInitialATR / mTickSize) + 0.9999) * mTickSize

mInitialBracketOrder.AdjustStop mInitialBracketOrder.stoporder.triggerPrice + IIf(mLongPosition, stopIncrement, -stopIncrement)

On Error Resume Next
For Each plex In mAdditionalBracketOrders
    plex.AdjustStop plex.stoporder.triggerPrice + IIf(mLongPosition, stopIncrement, -stopIncrement)
Next
On Error GoTo 0

' now buy or sell the required amount
newStopOffsetTicks = Int((mInitialStopFactor * mInitialATR / mTickSize) + 0.9999)
If mLongPosition Then
    Set mAdditionalBracketOrder = mOrderContext.buy(Number:=mUnitSize, _
                    entryType:=EntryTypeMarket, _
                    buyPrice:=0, _
                    buyOffset:=0, _
                    triggerPrice:=0, _
                    stopType:=StopTypeAuto, _
                    stopPrice:=0, _
                    stopOffset:=-newStopOffsetTicks, _
                    stopLimitPrice:=0, _
                    targetType:=TargetTypeNone, _
                    targetPrice:=0, _
                    targetOffset:=0, _
                    targetTriggerPrice:=0, _
                    cancelPrice:=0, _
                    stopStrategy:=Nothing)
    If Not mAdditionalBracketOrder Is Nothing Then
        mAdditionalBracketOrders.Add mAdditionalBracketOrder
        mAdditionalBracketOrder.EntryReason = "A" & mNumberOfIncrements
    End If
Else
    Set mAdditionalBracketOrder = mOrderContext.sell(Number:=mUnitSize, _
                    entryType:=EntryTypeMarket, _
                    sellPrice:=0, _
                    sellOffset:=0, _
                    triggerPrice:=0, _
                    stopType:=StopTypeAuto, _
                    stopPrice:=0, _
                    stopOffset:=newStopOffsetTicks, _
                    stopLimitPrice:=0, _
                    targetType:=TargetTypeNone, _
                    targetPrice:=0, _
                    targetOffset:=0, targetTriggerPrice:=0, _
                    cancelPrice:=0, _
                    stopStrategy:=Nothing)
    If Not mAdditionalBracketOrder Is Nothing Then
        mAdditionalBracketOrders.Add mAdditionalBracketOrder
        mAdditionalBracketOrder.EntryReason = "A" & mNumberOfIncrements
    End If
End If

If mNumberOfIncrements = mMaxIncrements Then
    If mListenerKey <> "" Then
        mBars.removeListener mListenerKey   ' don't need price updates any more
        mListenerKey = ""
        Exit Sub
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IPositionMgmtStrategy_Start(ByVal pBracketOrder As OrderUtils27.IBracketOrder)

End Sub

'================================================================================
' mAdditionalBracketOrder Event Handlers
'================================================================================

Private Sub mAdditionalBracketOrder_EntryOrderPartiallyFilled()
Dim priceIncrement As Double

If mNumberOfIncrements = mMaxIncrements Then
    If mListenerKey <> "" Then
        mBars.removeListener mListenerKey   ' don't need price updates any more
        mListenerKey = ""
        Exit Sub
    End If
End If
priceIncrement = Int((mScaleThresholdFactor * mInitialATR / mTickSize) + 0.9999) * mTickSize
mNextPriceThreshold = mAdditionalBracketOrder.EntryOrder.AveragePrice + IIf(mLongPosition, priceIncrement, -priceIncrement)
End Sub

'================================================================================
' mInitialBracketOrder Event Handlers
'================================================================================

Private Sub mInitialBracketOrder_Completed()
Dim i As Long

If mMaxIncrements = 0 Then Exit Sub     ' nothing to do

Set mInitialBracketOrder = Nothing
Set mAdditionalBracketOrder = Nothing
For i = mAdditionalBracketOrders.Count To 1 Step -1
    mAdditionalBracketOrders.Remove i
Next
If mListenerKey <> "" Then
    mBars.removeListener mListenerKey   ' don't need price updates any more
    mListenerKey = ""
End If
End Sub

Private Sub mInitialBracketOrder_EntryOrderPartiallyFilled()
Dim priceIncrement As Double

If mMaxIncrements = 0 Then Exit Sub     ' nothing to do

mLongPosition = IIf(mInitialBracketOrder.EntryOrder.Action = ActionBuy, True, False)

mInitialEntryPrice = mInitialBracketOrder.EntryOrder.AveragePrice

If mListenerKey = "" Then
    mListenerKey = mBars.addListener(Me, "lastprice")
End If

mInitialATR = mATRStudy.ATR(1)
mUnitSize = Int((mRiskIncrementPercent * mOrderContext.GetAccountBalance / 100) / (mTickValue * mInitialATR / mTickSize))
If mUnitSize = 0 Then mUnitSize = 1

priceIncrement = Int((mScaleThresholdFactor * mInitialATR / mTickSize) + 0.9999) * mTickSize
mNextPriceThreshold = mInitialEntryPrice + IIf(mLongPosition, priceIncrement, -priceIncrement)
End Sub

Private Sub mInitialBracketOrder_StopAdjusted()
Dim plex As IBracketOrder
Dim newStop As Double

'If Not internal Then Exit Sub

newStop = mInitialBracketOrder.stoporder.triggerPrice
On Error Resume Next    ' it's possible some of the order plexes have already been stopped out
For Each plex In mAdditionalBracketOrders
    plex.AdjustStop newStop
Next
On Error GoTo 0
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let ATRPeriods(ByVal value As Long)
mATRPeriods = value
End Property

Public Property Let InitialStopFactor(ByVal value As Double)
mInitialStopFactor = value
End Property

Public Property Let MaxIncrements(ByVal value As Long)
mMaxIncrements = value
End Property

Public Property Let RiskIncrementPercent(ByVal value As Double)
mRiskIncrementPercent = value
End Property

Public Property Let RiskUnitPercent(ByVal value As Double)
mRiskUnitPercent = value
End Property

Public Property Let ScaleThresholdFactor(ByVal value As Double)
mScaleThresholdFactor = value
End Property

Public Property Let StopIncrementFactor(ByVal value As Double)
mStopIncrementFactor = value
End Property

Public Property Let timePeriod(ByVal newValue As Integer)
mTimePeriod = newValue
End Property

Public Property Let TimePeriodUnit(ByVal newValue As String)
mTimePeriodUnit = newValue
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Initialise( _
                ByVal pContext As TradingContext)
Const ProcName As String = "nitialise"
On Error GoTo Err

Dim barsID As String
Dim studyID As String
Dim params As cParameters

Set mContext = pContext
mTickSize = mContext.Contract.TickSize
mTickValue = mContext.Contract.TickValue

Dim lTFD As TimeframeDesignator
Set lTFD = mContext.AddTimeframe(GetTimePeriod(mTimePeriod, mTimePeriodUnit), barsID)

Set params = New cParameters
params.setParam "valuetype", "lastprice"
params.setParam "periods", mATRPeriods
Set mATRStudy = mContext.AddStudy("studies2.CAverageTrueRangeStudy", _
                        studyID, _
                        barsID, _
                        params)
mATRStudy.numberOfHistoricValues = 2

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================




