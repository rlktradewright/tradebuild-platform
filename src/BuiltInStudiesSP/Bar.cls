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
Attribute VB_Exposed = False
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

Private mOpen As Double
Private mHigh As Double
Private mLow As Double
Private mClose As Double
Private mTickVolume As Long
Private mDateTime As Date
Private mBlank As Boolean   ' indicates that there were no ticks during the bar

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mBlank = True
mOpen = DummyLow
mLow = DummyLow
mHigh = DummyHigh
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

Friend Property Get Blank() As Boolean
Blank = mBlank
End Property

Friend Property Let Blank(ByVal newValue As Boolean)
mBlank = newValue
End Property

Friend Property Get closeValue() As Double
closeValue = mClose
End Property

Friend Property Get datetime() As Date
datetime = mDateTime
End Property

Friend Property Let datetime(ByVal newValue As Date)
mDateTime = newValue
End Property

Friend Property Get highValue() As Double
highValue = mHigh
End Property

Friend Property Get IsUpBar() As Boolean
IsUpBar = (mClose > mOpen)
End Property

Friend Property Get IsDownBar() As Boolean
IsDownBar = (mClose <= mOpen)
End Property

Friend Property Get lowValue() As Double
lowValue = mLow
End Property

Friend Property Get openValue() As Double
openValue = mOpen
End Property

Friend Property Get tickVolume() As Long
tickVolume = mTickVolume
End Property

'================================================================================
' Methods
'================================================================================

Public Sub Tick(ByVal value As Double)
If mBlank Then
    mBlank = False
    mOpen = value
End If

If value > mHigh Then mHigh = value

If value < mLow Then mLow = value

mClose = value
mTickVolume = mTickVolume + 1
End Sub

'================================================================================
' Helper Functions
'================================================================================



