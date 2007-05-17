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

Private WithEvents mTimerList  As TimerList
Attribute mTimerList.VB_VarHelpID = -1
Private WithEvents mTimerListItem As TimerListItem
Attribute mTimerListItem.VB_VarHelpID = -1

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mQuoteListeners = New Collection
Set mDataSignalListeners = New Collection
End Sub

'================================================================================
' QuoteListener Interface Members
'================================================================================

Private Sub QuoteListener_ask(ev As QuoteEvent)
fireSignalOn
startTimer
fireAsk ev.price, ev.Size
End Sub

Private Sub QuoteListener_bid(ev As QuoteEvent)
fireSignalOn
startTimer
fireBid ev.price, ev.Size
End Sub

Private Sub QuoteListener_high(ev As QuoteEvent)
fireSignalOn
startTimer
fireHigh ev.price
End Sub

Private Sub QuoteListener_Low(ev As QuoteEvent)
fireSignalOn
startTimer
fireLow ev.price
End Sub

Private Sub QuoteListener_openInterest(ev As QuoteEvent)
fireSignalOn
startTimer
fireOpenInterest ev.Size
End Sub

Private Sub QuoteListener_previousClose(ev As QuoteEvent)
fireSignalOn
startTimer
firePreviousClose ev.price
End Sub

Private Sub QuoteListener_trade(ev As QuoteEvent)
fireSignalOn
startTimer
fireTrade ev.price, ev.Size
End Sub

Private Sub QuoteListener_volume(ev As QuoteEvent)
fireSignalOn
startTimer
fireVolume ev.Size
End Sub

'================================================================================
' mTimerListItem Event Handlers
'================================================================================

Private Sub mTimerListItem_StateChange(ev As TWUtilities.StateChangeEvent)
If ev.state <> TimerListItemStates.TimerListItemStateExpired Then Exit Sub
Set mTimerListItem = Nothing
fireSignalOff
End Sub

'================================================================================
' Properties
'================================================================================

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

Friend Sub initialise( _
                ByVal contract As contract, _
                ByVal timer As TimerList)
Set mContract = contract
Set mTimerList = timer
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
    listener.low quote
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

Private Sub startTimer()
If Not mTimerListItem Is Nothing Then
    If (mTimerListItem.expiryTime - GetTimestamp) * 8640000 > 100 Then Exit Sub
    mTimerList.Remove mTimerListItem
End If
Set mTimerListItem = mTimerList.Add(Me, 200, ExpiryTimeUnitMilliseconds)
fireSignalOn
End Sub

