VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Quotes"
Attribute VB_GlobalNameSpace = True
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

Private Const ModuleName                            As String = "Quotes"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get AskPrice() As Double
Const ProcName As String = "AskPrice"
On Error GoTo Err

AskPrice = gTradingContext.AskQuote.Price

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get AskSize() As Long
Const ProcName As String = "AskSize"
On Error GoTo Err

AskSize = gTradingContext.AskQuote.Size

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get AskTimestamp() As Date
Const ProcName As String = "AskTimestamp"
On Error GoTo Err

AskTimestamp = gTradingContext.AskQuote.Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BidPrice() As Double
Const ProcName As String = "BidPrice"
On Error GoTo Err

BidPrice = gTradingContext.BidQuote.Price

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BidSize() As Long
Const ProcName As String = "BidSize"
On Error GoTo Err

BidSize = gTradingContext.BidQuote.Size

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BidTimestamp() As Date
Const ProcName As String = "BidTimestamp"
On Error GoTo Err

BidTimestamp = gTradingContext.BidQuote.Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ClosePrice() As Double
Const ProcName As String = "ClosePrice"
On Error GoTo Err

ClosePrice = gTradingContext.CloseQuote.Price

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CloseTimestamp() As Date
Const ProcName As String = "CloseTimestamp"
On Error GoTo Err

CloseTimestamp = gTradingContext.CloseQuote.Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HighPrice() As Double
Const ProcName As String = "HighPrice"
On Error GoTo Err

HighPrice = gTradingContext.HighQuote.Price

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HighTimestamp() As Date
Const ProcName As String = "HighTimestamp"
On Error GoTo Err

HighTimestamp = gTradingContext.HighQuote.Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LowPrice() As Double
Const ProcName As String = "LowPrice"
On Error GoTo Err

LowPrice = gTradingContext.LowQuote.Price

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LowTimestamp() As Date
Const ProcName As String = "LowTimestamp"
On Error GoTo Err

LowTimestamp = gTradingContext.LowQuote.Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OpenPrice() As Double
Const ProcName As String = "OpenPrice"
On Error GoTo Err

OpenPrice = gTradingContext.OpenQuote.Price

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OpenTimestamp() As Date
Const ProcName As String = "OpenTimestamp"
On Error GoTo Err

OpenTimestamp = gTradingContext.OpenQuote.Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsCurrentTickAsk() As Boolean
IsCurrentTickAsk = (gTradingContext.CurrentTick.TickType = TickTypeAsk)
End Property

Public Property Get IsCurrentTickBid() As Boolean
IsCurrentTickBid = (gTradingContext.CurrentTick.TickType = TickTypeBid)
End Property

Public Property Get IsCurrentTickTrade() As Boolean
IsCurrentTickTrade = (gTradingContext.CurrentTick.TickType = TickTypeTrade)
End Property

Public Property Get IsCurrentTickVolume() As Boolean
IsCurrentTickVolume = (gTradingContext.CurrentTick.TickType = TickTypeVolume)
End Property

Public Property Get TradePrice() As Double
Const ProcName As String = "TradePrice"
On Error GoTo Err

TradePrice = gTradingContext.TradeQuote.Price

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TradeSize() As Long
Const ProcName As String = "TradeSize"
On Error GoTo Err

TradeSize = gTradingContext.TradeQuote.Size

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TradeTimestamp() As Date
Const ProcName As String = "TradeTimestamp"
On Error GoTo Err

TradeTimestamp = gTradingContext.TradeQuote.Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Volume() As Long
Const ProcName As String = "Volume"
On Error GoTo Err

Volume = gTradingContext.VolumeQuote.Size

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get VolumeTimestamp() As Date
Const ProcName As String = "VolumeTimestamp"
On Error GoTo Err

VolumeTimestamp = gTradingContext.VolumeQuote.Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub StartTickData(ByVal pStrategy As IPositionMgmtStrategy)
Const ProcName As String = "StartTickData"
On Error GoTo Err

gTradingContext.StartTickData pStrategy

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StopTickData(ByVal pStrategy As IPositionMgmtStrategy)
Const ProcName As String = "StopTickData"
On Error GoTo Err

gTradingContext.StopTickData pStrategy

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




