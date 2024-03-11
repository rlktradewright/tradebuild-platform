VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradingContext"
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

Private Const ModuleName                            As String = "TradingContext"

'@================================================================================
' Member variables
'@================================================================================

Private mStrategyRunner                             As StrategyRunner

Private mTicker                                     As Ticker
Private mContract                                   As IContract
Private mSession                                    As Session

Private mPositionManager                            As PositionManager
Private WithEvents mBracketOrders                   As BracketOrders
Attribute mBracketOrders.VB_VarHelpID = -1

Private mOrderContextBuy                            As OrderContext
Private mOrderContextBuySimulated                   As OrderContext
Private mOrderContextSell                           As OrderContext
Private mOrderContextSellSimulated                  As OrderContext

Private mPositionManagerSimulated                   As PositionManager
Private mTimeframes                                 As Timeframes

Private mTradeReasonLogger                          As Logger

Private mTickNumber                                 As Long

Private mStrategy                                   As IStrategy

Private mActivePositionMgmgtStrategies              As Collection
Private mPositionMgmgtResourceContexts              As Collection

Private mClock                                      As Clock

Private mReadyForTicks                              As Boolean

Private mCurrentTick                                As GenericTick

Private mFutureBuilder                              As New FutureBuilder

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mTradeReasonLogger = GetLogger("strategy.tradereason")
mTradeReasonLogger.LogToParent = False

Set mActivePositionMgmgtStrategies = New Collection
Set mPositionMgmgtResourceContexts = New Collection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
On Error GoTo Err

If ev.ChangeType = OrderContextChangeTypes.OrderContextReadyStateChanged Then
    Dim lOrderContext As OrderContext
    Set lOrderContext = ev.Source
    If lOrderContext.IsReadyTickSize Then
        lOrderContext.RemoveChangeListener Me
        setReady
    Else
        gLogger.Log "order context " & lOrderContext.Name & _
                        " not ready: " & lOrderContext.NotReadyMessage, _
                    ProcName, ModuleName, LogLevelDetail
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IGenericTickListener_NotifyTick Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Static sGotFirstAsk As Boolean
Static sGotFirstBid As Boolean
   
If ev.Tick.TickType = TickTypeAsk Then sGotFirstAsk = True
If ev.Tick.TickType = TickTypeBid Then sGotFirstBid = True

If Not (sGotFirstAsk And sGotFirstBid And mReadyForTicks) Then Exit Sub

Select Case ev.Tick.TickType
Case TickTypes.TickTypeAsk, _
        TickTypes.TickTypeBid, _
        TickTypes.TickTypeTrade, _
        TickTypes.TickTypeVolume
Case Else
    Exit Sub
End Select

mTickNumber = mTickNumber + 1
mCurrentTick = ev.Tick

Dim lPMStrategy As IPositionMgmtStrategy
For Each lPMStrategy In mActivePositionMgmgtStrategies
    mStrategyRunner.SetCurrent lPMStrategy, mPositionMgmgtResourceContexts.Item(GetObjectKey(lPMStrategy))
    lPMStrategy.NotifyTick ev.Tick.TickType
    mStrategyRunner.SetNotCurrent
Next


mStrategyRunner.SetCurrent mStrategy, Nothing
mStrategy.NotifyTick mCurrentTick.TickType
mStrategyRunner.SetNotCurrent

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

If ev.ChangeType <> CollItemAdded Then Exit Sub

CreateBracketProfitCalculator ev.AffectedItem, mTicker

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let AllowUnprotectedPositions(ByVal Value As Boolean)
mOrderContextBuy.PreventUnprotectedPositions = Not Value
mOrderContextSell.PreventUnprotectedPositions = Not Value
End Property

Public Property Get AskQuote() As Quote
Const ProcName As String = "AskQuote"
On Error GoTo Err

AskQuote = mTicker.CurrentQuote(TickTypeAsk)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BidQuote() As Quote
Const ProcName As String = "BidQuote"
On Error GoTo Err

BidQuote = mTicker.CurrentQuote(TickTypeBid)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CloseQuote() As Quote
Const ProcName As String = "CloseQuote"
On Error GoTo Err

CloseQuote = mTicker.CurrentQuote(TickTypeClosePrice)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Contract() As IContract
Const ProcName As String = "Contract"
On Error GoTo Err

Set Contract = mContract

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentTick() As GenericTick
CurrentTick = mCurrentTick
End Property

Public Property Get DefaultBuyOrderContext() As OrderContext
Const ProcName As String = "DefaultBuyOrderContext"
On Error GoTo Err

Set DefaultBuyOrderContext = mOrderContextBuy

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultBuyOrderContextSimulated() As OrderContext
Const ProcName As String = "DefaultBuyOrderContextSimulated"
On Error GoTo Err

Set DefaultBuyOrderContextSimulated = mOrderContextBuySimulated

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultSellOrderContext() As OrderContext
Const ProcName As String = "DefaultSellOrderContext"
On Error GoTo Err

Set DefaultSellOrderContext = mOrderContextSell

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultSellOrderContextSimulated() As OrderContext
Const ProcName As String = "DefaultSellOrderContextSimulated"
On Error GoTo Err

Set DefaultSellOrderContextSimulated = mOrderContextSellSimulated

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HighQuote() As Quote
Const ProcName As String = "HighQuote"
On Error GoTo Err

HighQuote = mTicker.CurrentQuote(TickTypeHighPrice)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LowQuote() As Quote
Const ProcName As String = "LowQuote"
On Error GoTo Err

LowQuote = mTicker.CurrentQuote(TickTypeLowPrice)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OpenQuote() As Quote
Const ProcName As String = "OpenQuote"
On Error GoTo Err

OpenQuote = mTicker.CurrentQuote(TickTypeOpenPrice)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PositionManager() As PositionManager
Const ProcName As String = "PositionManager"
On Error GoTo Err

Set PositionManager = mPositionManager

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PositionManagerSimulated() As PositionManager
Const ProcName As String = "PositionManagerSimulated"
On Error GoTo Err

Set PositionManagerSimulated = mPositionManagerSimulated

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Session() As Session
Set Session = mSession
End Property

Public Property Get TickNumber() As Long
Const ProcName As String = "TickNumber"
TickNumber = mTickNumber
End Property

Public Property Get Timestamp() As Date
Const ProcName As String = "Timestamp"
On Error GoTo Err

Timestamp = mClock.Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TradeQuote() As Quote
Const ProcName As String = "TradeQuote"
On Error GoTo Err

TradeQuote = mTicker.CurrentQuote(TickTypeTrade)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get VolumeQuote() As Quote
Const ProcName As String = "VolumeQuote"
On Error GoTo Err

VolumeQuote = mTicker.CurrentQuote(TickTypeVolume)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ApplyPositionManagementStrategy( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pStrategy As IPositionMgmtStrategy, _
                ByVal pResourceContext As ResourceContext)
Const ProcName As String = "ApplyPositionManagementStrategy"
On Error GoTo Err

mPositionMgmgtResourceContexts.Add pResourceContext, GetObjectKey(pStrategy)

mStrategyRunner.SetCurrent pStrategy, pResourceContext
mStrategyRunner.RequestBracketOrderNotification pBracketOrder, pStrategy, pResourceContext
pStrategy.Start Me, mStrategyRunner.GetResourceIdForBracketOrder(pBracketOrder)
mStrategyRunner.SetNotCurrent

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function getBars( _
                ByVal pTimeframe As Timeframe) As BarUtils27.Bars
Const ProcName As String = "GetBars"
On Error GoTo Err

Set getBars = pTimeframe.BarStudy.BarsFuture.Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetStudyValue( _
                ByVal Study As ResourceIdentifier, _
                Optional ByVal ValueName As String, _
                Optional ByVal Ref As Long) As Variant
Const ProcName As String = "GetStudyValue"
On Error GoTo Err

AssertArgument Not Study Is Nothing, "Study is nothing"
AssertArgument TypeOf Study.Resource Is IStudy, "Study does not refer to a Study object"

Dim lStudy As IStudy: Set lStudy = Study.Resource
If ValueName = "" Then ValueName = lStudy.StudyDefinition.DefaultValueName
GetStudyValue = lStudy.GetStudyValue(ValueName, Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function Initialise( _
                ByVal pStrategyRunner As StrategyRunner, _
                ByVal pStrategy As IStrategy, _
                ByVal pTicker As Ticker) As IFuture
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mStrategyRunner = pStrategyRunner
Set mStrategy = pStrategy

Set mTicker = pTicker
Set mClock = pTicker.ClockFuture.Value
Set mContract = mTicker.ContractFuture.Value
Set mSession = mTicker.SessionFuture.Value
mTicker.AddGenericTickListener Me

Set mPositionManager = mTicker.PositionManager
Set mBracketOrders = mPositionManager.BracketOrders
Set mPositionManagerSimulated = mTicker.PositionManagerSimulated

Set mOrderContextBuy = mPositionManager.OrderContexts.Add("BUY")
gLogger.Log "mOrderContextBuy: id=" & GetObjectKey(mOrderContextBuy), ProcName, ModuleName, LogLevelDetail

If Not mOrderContextBuy.IsReadyTickSize Then
    mOrderContextBuy.AddChangeListener Me
Else
    setReady
End If

Set mOrderContextSell = mPositionManager.OrderContexts.Add("SELL")
gLogger.Log "mOrderContextSell: id=" & GetObjectKey(mOrderContextSell), ProcName, ModuleName, LogLevelDetail

Set mOrderContextBuySimulated = mPositionManagerSimulated.OrderContexts.Add("BUY-Simulated")
gLogger.Log "mOrderContextBuySimulated: id=" & GetObjectKey(mOrderContextBuySimulated), ProcName, ModuleName, LogLevelDetail

Set mOrderContextSellSimulated = mPositionManagerSimulated.OrderContexts.Add("SELL-Simulated")
gLogger.Log "mOrderContextSellSimulated: id=" & GetObjectKey(mOrderContextSellSimulated), ProcName, ModuleName, LogLevelDetail

Set Initialise = mFutureBuilder.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub LogTradeReason(ByVal pReason As String, Optional ByVal pLogLevel As LogLevels = LogLevels.LogLevelNormal)
If pReason = "" Then Exit Sub

Dim lTimepart As String
lTimepart = FormatTimestamp(Timestamp, TimestampDateAndTimeISO8601) & "  "

If InStr(1, pReason, vbCrLf) = 0 Then
    mTradeReasonLogger.Log pLogLevel, lTimepart & pReason
Else
    Dim lSpacer As String
    lSpacer = Space(Len(lTimepart))
    mTradeReasonLogger.Log pLogLevel, lTimepart & Replace$(pReason, vbCrLf, vbCrLf & lSpacer)
End If

End Sub

Public Sub StartTickData(ByVal pStrategy As IPositionMgmtStrategy)
Const ProcName As String = "StartTickData"
On Error GoTo Err

mActivePositionMgmgtStrategies.Add pStrategy, GetObjectKey(pStrategy)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StopTickData(ByVal pStrategy As IPositionMgmtStrategy)
Const ProcName As String = "StopTickData"
On Error GoTo Err

mActivePositionMgmgtStrategies.Remove GetObjectKey(pStrategy)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub setReady()
Const ProcName As String = "setReady"
On Error GoTo Err

mReadyForTicks = True
mFutureBuilder.Value = CreateBoxedValue("READY!")
mFutureBuilder.Complete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

