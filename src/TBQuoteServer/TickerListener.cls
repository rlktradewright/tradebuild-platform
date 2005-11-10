VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickerListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements QuoteListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mContract As contract

Private mQuoteListeners             As Collection
Private mDataSignalListeners        As Collection

Private mIndex As Long

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mQuoteListeners = New Collection
Set mDataSignalListeners = New Collection
Set mTimer = New IntervalTimer
mTimer.RepeatNotifications = False
mTimer.TimerIntervalMillisecs = 200
End Sub

'================================================================================
' QuoteListener Interface Members
'================================================================================

Private Sub QuoteListener_ask(ev As TradeBuild.QuoteEvent)
fireSignalOn
mTimer.StartTimer
fireAsk ev.price, ev.Size
End Sub

Private Sub QuoteListener_bid(ev As TradeBuild.QuoteEvent)
fireSignalOn
mTimer.StartTimer
fireBid ev.price, ev.Size
End Sub

Private Sub QuoteListener_high(ev As TradeBuild.QuoteEvent)
fireSignalOn
mTimer.StartTimer
fireHigh ev.price
End Sub

Private Sub QuoteListener_Low(ev As TradeBuild.QuoteEvent)
fireSignalOn
mTimer.StartTimer
fireLow ev.price
End Sub

Private Sub QuoteListener_openInterest(ev As TradeBuild.QuoteEvent)
fireSignalOn
mTimer.StartTimer
fireOpenInterest ev.Size
End Sub

Private Sub QuoteListener_previousClose(ev As TradeBuild.QuoteEvent)
fireSignalOn
mTimer.StartTimer
firePreviousClose ev.price
End Sub

Private Sub QuoteListener_trade(ev As TradeBuild.QuoteEvent)
fireSignalOn
mTimer.StartTimer
fireTrade ev.price, ev.Size
End Sub

Private Sub QuoteListener_volume(ev As TradeBuild.QuoteEvent)
fireSignalOn
mTimer.StartTimer
fireVolume ev.Size
End Sub

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
fireSignalOff
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let contract(ByVal value As contract)
Set mContract = value
End Property

Public Property Get contract() As contract
Set contract = mContract
End Property

Public Property Let index(ByVal value As Long)
mIndex = value
End Property

Public Property Get index() As Long
index = mIndex
End Property

'================================================================================
' Methods
'================================================================================

Public Sub addDataSignalListener(ByVal value As DataSignalListener)
mDataSignalListeners.Add value
End Sub

Public Sub addQuoteListener(ByVal value As QuoteListener)
mQuoteListeners.Add value
End Sub

Public Sub removeDataSignalListener(ByVal value As DataSignalListener)
Dim i As Long
For i = mDataSignalListeners.Count To 1 Step -1
    If mDataSignalListeners.Item(i) Is value Then mDataSignalListeners.Remove i
Next
End Sub

Public Sub removeQuoteListener(ByVal value As QuoteListener)
Dim i As Long
For i = mQuoteListeners.Count To 1 Step -1
    If mQuoteListeners.Item(i) Is value Then mQuoteListeners.Remove i
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub fireAsk(ByVal price As Double, _
                ByVal Size As Long)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = price
quote.Size = Size
For Each listener In mQuoteListeners
    listener.ask quote
Next
End Sub

Private Sub fireBid(ByVal price As Double, _
                ByVal Size As Long)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = price
quote.Size = Size
For Each listener In mQuoteListeners
    listener.bid quote
Next
End Sub

Private Sub fireHigh(ByVal price As Double)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = price
quote.Size = 0
For Each listener In mQuoteListeners
    listener.high quote
Next
End Sub

Private Sub fireLow(ByVal price As Double)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = price
quote.Size = 0
For Each listener In mQuoteListeners
    listener.Low quote
Next
End Sub

Private Sub fireOpenInterest(ByVal Size As Long)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = 0
quote.Size = Size
For Each listener In mQuoteListeners
    listener.openInterest quote
Next
End Sub

Private Sub firePreviousClose(ByVal price As Double)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = price
quote.Size = 0
For Each listener In mQuoteListeners
    listener.previousClose quote
Next
End Sub

Private Sub fireSignalOff()
Dim listener As DataSignalListener
Dim datasignal As DataSignalEvent
Set datasignal.Source = Me
For Each listener In mDataSignalListeners
    listener.signalOff datasignal
Next
End Sub

Private Sub fireSignalOn()
Dim listener As DataSignalListener
Dim datasignal As DataSignalEvent
Set datasignal.Source = Me
For Each listener In mDataSignalListeners
    listener.signalOn datasignal
Next
End Sub

Private Sub fireTrade(ByVal price As Double, _
                ByVal Size As Long)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = price
quote.Size = Size
For Each listener In mQuoteListeners
    listener.trade quote
Next
End Sub

Private Sub fireVolume(ByVal Size As Long)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = 0
quote.Size = Size
For Each listener In mQuoteListeners
    listener.volume quote
Next
End Sub



