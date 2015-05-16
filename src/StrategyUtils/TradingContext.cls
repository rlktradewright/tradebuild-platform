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

Private mTicker                                     As Ticker

Private mPositionManager                            As PositionManager
Private WithEvents mBracketOrders                   As BracketOrders
Attribute mBracketOrders.VB_VarHelpID = -1

Private mPositionManagerDummy                       As PositionManager
Private mTimeframes                                 As Timeframes

Private mTradeReasonLogger                          As Logger

Private mTickNumber                                 As Long

Private mStrategy                                   As IStrategy

Private mPositionMgmgtStrategies                    As Collection
Private mActivePositionMgmgtStrategies              As Collection

Private mReceivedFirstBidAndAsk                     As Boolean

Private mClock                                      As Clock

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mTradeReasonLogger = GetLogger("strategy.tradereason")
mTradeReasonLogger.LogToParent = True

Set mPositionMgmgtStrategies = New Collection
Set mActivePositionMgmgtStrategies = New Collection

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

mTickNumber = mTickNumber + 1

Static sGotFirstAsk As Boolean
Static sGotFirstBid As Boolean
   
If ev.Tick.TickType = TickTypeAsk Then sGotFirstAsk = True
If ev.Tick.TickType = TickTypeBid Then sGotFirstBid = True
If sGotFirstAsk And sGotFirstBid Then mReceivedFirstBidAndAsk = True

Dim lPMStrategy As IPositionMgmtStrategy
For Each lPMStrategy In mActivePositionMgmgtStrategies
    lPMStrategy.NotifyTick ev.Tick.TickType
Next

mStrategy.NotifyTick ev.Tick.TickType

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

Set Contract = mTicker.ContractFuture.Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultOrderContext() As OrderContext
Const ProcName As String = "DefaultOrderContext"
On Error GoTo Err

Set DefaultOrderContext = mPositionManager.OrderContexts.DefaultOrderContext

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultOrderContextDummy() As OrderContext
Const ProcName As String = "DefaultOrderContextDummy"
On Error GoTo Err

Set DefaultOrderContextDummy = mPositionManagerDummy.OrderContexts.DefaultOrderContext

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

Public Property Get PositionManagerDummy() As PositionManager
Const ProcName As String = "PositionManagerDummy"
On Error GoTo Err

Set PositionManagerDummy = mPositionManagerDummy

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ReceivedFirstBidAndAsk() As Boolean
ReceivedFirstBidAndAsk = mReceivedFirstBidAndAsk
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
                ByVal pStrategy As IPositionMgmtStrategy)
Const ProcName As String = "ApplyPositionManagementStrategy"
On Error GoTo Err

mPositionMgmgtStrategies.Add pStrategy
pStrategy.Start Me, pBracketOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetBars( _
                ByVal pTimeframe As Timeframe) As Bars
Const ProcName As String = "GetBars"
On Error GoTo Err

Set GetBars = pTimeframe.BarStudy.BarsFuture.Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pStrategy As IStrategy, _
                ByVal pTicker As Ticker)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mStrategy = pStrategy
Set mTicker = pTicker
Set mClock = pTicker.ClockFuture.Value
mTicker.AddGenericTickListener Me
Set mPositionManager = mTicker.PositionManager
Set mBracketOrders = mPositionManager.BracketOrders
Set mPositionManagerDummy = mTicker.PositionManagerSimulated

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LogTradeReason(ByVal pReason As String, Optional ByVal pLogLevel As LogLevels = LogLevels.LogLevelNormal)
mTradeReasonLogger.Log pLogLevel, FormatTimestamp(Timestamp, TimestampDateAndTimeISO8601) & "  " & pReason
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



