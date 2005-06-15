VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bar"
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

Private mBarType As TradeBuild.BarTypes
Private mBarLength As Long
Private mOpen As Double
Private mHigh As Double
Private mLow As Double
Private mClose As Double
Private mVolume As Long
Private mTickVolume As Long
Private mOpenInterest As Long
Private mDateTime As Date
Private mBlank As Boolean   ' indicates that there were no ticks during the bar

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mBlank = True
mOpen = 0
mLow = MaxDouble
mHigh = MinDouble
mClose = 0
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get BarLength() As Long
BarLength = mBarLength
End Property

Public Property Let BarLength(ByVal value As Long)
mBarLength = value
End Property

Public Property Get BarType() As TradeBuild.BarTypes
BarType = mBarType
End Property

Public Property Let BarType(ByVal value As TradeBuild.BarTypes)
mBarType = value
End Property

Public Property Get Blank() As Boolean
Blank = mBlank
End Property

Public Property Let Blank(ByVal newValue As Boolean)
mBlank = newValue
End Property

Public Property Get ClosePrice() As Double
ClosePrice = mClose
End Property

Public Property Let ClosePrice(newValue As Double)
mBlank = False
mClose = newValue
End Property

Public Property Get DateTime() As Date
DateTime = mDateTime
End Property

Public Property Let DateTime(ByVal newValue As Date)
mDateTime = newValue
End Property

Public Property Get HighPrice() As Double
HighPrice = mHigh
End Property

Public Property Let HighPrice(newValue As Double)
mBlank = False
mHigh = newValue
End Property

Public Property Get IsUpBar() As Boolean
IsUpBar = (mClose > mOpen)
End Property

Public Property Get IsDownBar() As Boolean
IsDownBar = (mClose <= mOpen)
End Property

Public Property Get LowPrice() As Double
LowPrice = mLow
End Property

Public Property Let LowPrice(newValue As Double)
mBlank = False
mLow = newValue
End Property

Public Property Get OpenInterest() As Long
OpenInterest = mOpenInterest
End Property

Public Property Let OpenInterest(newValue As Long)
mBlank = False
mOpenInterest = newValue
End Property

Public Property Get OpenPrice() As Double
OpenPrice = mOpen
End Property

Public Property Let OpenPrice(newValue As Double)
mBlank = False
mOpen = newValue
End Property

Public Property Get TickVolume() As Long
TickVolume = mTickVolume
End Property

Public Property Let TickVolume(newValue As Long)
mBlank = False
mTickVolume = newValue
End Property

Public Property Get Volume() As Long
Volume = mVolume
End Property

Public Property Let Volume(newValue As Long)
If newValue = 0 Then Exit Property
mBlank = False
mVolume = newValue
End Property

'================================================================================
' Methods
'================================================================================

Public Sub Tick(ByVal price As Double)
mBlank = False
If mOpen = 0 Then mOpen = price
If price > mHigh Then mHigh = price
If price < mLow Then mLow = price
mClose = price
mTickVolume = mTickVolume + 1
End Sub

Public Sub TickSize(ByVal Size As Long)
Volume = Volume + Size
End Sub

'================================================================================
' Helper Functions
'================================================================================


