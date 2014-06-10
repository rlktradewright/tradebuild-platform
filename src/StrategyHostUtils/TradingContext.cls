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

Private mDataSource                                 As IMarketDataSource
Private mPositionManager                            As PositionManager
Private mPositionManagerSimulated                   As PositionManager
Private mTimeframes                                 As Timeframes

Private mHostController                             As HostController

Private mTradeReasonLogger                          As Logger

Private mTickNumber                                 As Long

Private mStrategy                                   As IStrategy

Private mReceivedFirstBidAndAsk                     As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mTradeReasonLogger = GetLogger("tradereason")
mTradeReasonLogger.LogToParent = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IGenericTickListener_NotifyTick Interface Members
'@================================================================================

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

mTickNumber = mTickNumber + 1

Static sGotFirstAsk As Boolean
Static sGotFirstBid As Boolean
   
If ev.Tick.TickType = TickTypeAsk Then sGotFirstAsk = True
If ev.Tick.TickType = TickTypeBid Then sGotFirstBid = True
If sGotFirstAsk And sGotFirstBid Then mReceivedFirstBidAndAsk = True

mStrategy.NotifyTick ev.Tick.TickType

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get AskQuote() As Quote
Const ProcName As String = "AskQuote"
On Error GoTo Err

AskQuote = mDataSource.CurrentQuote(TickTypeAsk)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BidQuote() As Quote
Const ProcName As String = "BidQuote"
On Error GoTo Err

BidQuote = mDataSource.CurrentQuote(TickTypeBid)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CloseQuote() As Quote
Const ProcName As String = "CloseQuote"
On Error GoTo Err

CloseQuote = mDataSource.CurrentQuote(TickTypeClosePrice)

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

Public Property Get DefaultOrderContextSimulated() As OrderContext
Const ProcName As String = "DefaultOrderContextSimulated"
On Error GoTo Err

Set DefaultOrderContextSimulated = mPositionManagerSimulated.OrderContexts.DefaultOrderContext

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HighQuote() As Quote
Const ProcName As String = "HighQuote"
On Error GoTo Err

HighQuote = mDataSource.CurrentQuote(TickTypeHighPrice)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LowQuote() As Quote
Const ProcName As String = "LowQuote"
On Error GoTo Err

LowQuote = mDataSource.CurrentQuote(TickTypeLowPrice)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OpenQuote() As Quote
Const ProcName As String = "OpenQuote"
On Error GoTo Err

OpenQuote = mDataSource.CurrentQuote(TickTypeOpenPrice)

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

Timestamp = mDataSource.Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TradeQuote() As Quote
Const ProcName As String = "TradeQuote"
On Error GoTo Err

TradeQuote = mDataSource.CurrentQuote(TickTypeTrade)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get VolumeQuote() As Quote
Const ProcName As String = "VolumeQuote"
On Error GoTo Err

VolumeQuote = mDataSource.CurrentQuote(TickTypeVolume)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function GetBars( _
                ByVal pTimeframe As Timeframe) As Bars
Const ProcName As String = "GetBars"
On Error GoTo Err

Set GetBars = pTimeframe.BarStudy.BarsFuture.Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub LogTradeReason(ByVal pReason As String, Optional ByVal pLogLevel As LogLevels = LogLevels.LogLevelNormal)
mTradeReasonLogger.Log pLogLevel, pReason
End Sub

'@================================================================================
' Methods
'@================================================================================

Public Sub ApplyPositionManagementStrategy( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pStrategy As IPositionMgmtStrategy)
Const ProcName As String = "ApplyPositionManagementStrategy"
On Error GoTo Err

pStrategy.Start pBracketOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pHostController As HostController, _
                ByVal pStrategy As IStrategy, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pPositionManager As PositionManager, _
                ByVal pPositionManagerSimulated As PositionManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mHostController = pHostController
Set mStrategy = pStrategy
Set mDataSource = pDataSource
mDataSource.AddGenericTickListener Me
Set mPositionManager = pPositionManager
Set mPositionManagerSimulated = pPositionManagerSimulated

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



