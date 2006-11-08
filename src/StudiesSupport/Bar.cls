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

Public Property Get Blank() As Boolean
Blank = mBlank
End Property

Friend Property Let Blank(ByVal newValue As Boolean)
mBlank = newValue
End Property

Public Property Get closeValue() As Double
closeValue = mClose
End Property

Public Property Get datetime() As Date
datetime = mDateTime
End Property

Friend Property Let datetime(ByVal newValue As Date)
mDateTime = newValue
End Property

Public Property Get gain() As Double
gain = mClose - mOpen
If gain < 0 Then gain = 0
End Property

Public Property Get highValue() As Double
highValue = mHigh
End Property

Public Property Get IsDownBar() As Boolean
IsDownBar = (mClose < mOpen)
End Property

Public Property Get IsUpBar() As Boolean
IsUpBar = (mClose > mOpen)
End Property

Public Property Get loss() As Double
loss = mOpen - mClose
If loss < 0 Then loss = 0
End Property

Public Property Get lowValue() As Double
lowValue = mLow
End Property

Public Property Get openValue() As Double
openValue = mOpen
End Property

Public Property Get tickVolume() As Long
tickVolume = mTickVolume
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Tick(ByVal value As Double)
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




