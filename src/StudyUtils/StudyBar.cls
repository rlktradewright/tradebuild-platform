VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mBarChanged As Boolean

Private mBarNumber As Long

Private mOpen As Double
Private mOpenChanged As Boolean

Private mHigh As Double
Private mHighChanged As Boolean

Private mLow As Double
Private mLowChanged As Boolean

Private mClose As Double
Private mCloseChanged As Boolean

Private mTickVolume As Long
Private mDateTime As Date
Private mBlank As Boolean   ' indicates that there were no ticks during the bar

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mBlank = True
mOpen = DummyLow
mLow = DummyLow
mHigh = DummyHigh
mClose = DummyLow
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get barChanged() As Boolean
barChanged = mBarChanged
End Property

Friend Property Let barNumber(ByVal value As Long)
mBarNumber = value
End Property

Public Property Get barNumber() As Long
barNumber = mBarNumber
End Property

Public Property Get Blank() As Boolean
Blank = mBlank
End Property

Friend Property Let Blank(ByVal newvalue As Boolean)
mBlank = newvalue
End Property

Public Property Get closeValue() As Double
closeValue = mClose
End Property

Public Property Get closeChanged() As Boolean
closeChanged = mCloseChanged
End Property

Public Property Get datetime() As Date
datetime = mDateTime
End Property

Friend Property Let datetime(ByVal newvalue As Date)
mDateTime = newvalue
End Property

Public Property Get gain() As Double
gain = mClose - mOpen
If gain < 0 Then gain = 0
End Property

Public Property Get highValue() As Double
highValue = mHigh
End Property

Public Property Get highChanged() As Boolean
highChanged = mHighChanged
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

Public Property Get lowChanged() As Boolean
lowChanged = mLowChanged
End Property

Public Property Get openValue() As Double
openValue = mOpen
End Property

Public Property Get openChanged() As Boolean
openChanged = mOpenChanged
End Property

Public Property Get tickVolume() As Long
tickVolume = mTickVolume
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub incrementTickVolume( _
                ByVal increment As Long)
mTickVolume = mTickVolume + increment
End Sub

Friend Sub Tick(ByVal value As Double)
mBarChanged = False
mOpenChanged = False
mHighChanged = False
mLowChanged = False
mCloseChanged = False

If mBlank Then
    mBlank = False
    mOpen = value
    mBarChanged = True
    mOpenChanged = True
End If

If value > mHigh Then
    mHigh = value
    mBarChanged = True
    mHighChanged = True
End If

If value < mLow Then
    mLow = value
    mBarChanged = True
    mLowChanged = True
End If

If value <> mClose Then
    mClose = value
    mBarChanged = True
    mCloseChanged = True
End If
mTickVolume = mTickVolume + 1
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




