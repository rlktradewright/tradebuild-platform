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

Event CloseChanged()
Event HighChanged()
Event LowChanged()
Event OpenChanged()
Event TickVolumeChanged()
Event VolumeChanged()

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

Private mBarType As Long
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

Private mDirty As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mBlank = True
mOpen = MinDouble
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

Public Property Get barLength() As Long
barLength = mBarLength
End Property

Friend Property Let barLength(ByVal value As Long)
mBarLength = value
End Property

Public Property Get barType() As Long
barType = mBarType
End Property

Friend Property Let barType(ByVal value As Long)
mBarType = value
End Property

Public Property Get Blank() As Boolean
Blank = mBlank
End Property

Friend Property Let Blank(ByVal newValue As Boolean)
mBlank = newValue
End Property

Public Property Get closeValue() As Double
closeValue = mClose
End Property

Friend Property Let closeValue(newValue As Double)
mBlank = False
mDirty = True
mClose = newValue
End Property

Public Property Get datetime() As Date
datetime = mDateTime
End Property

Friend Property Let datetime(ByVal newValue As Date)
mDateTime = newValue
End Property

Friend Property Get Dirty() As Boolean
Dirty = mDirty
End Property

Friend Property Let Dirty(ByVal newValue As Boolean)
mDirty = newValue
End Property

Public Property Get highValue() As Double
highValue = mHigh
End Property

Friend Property Let highValue(newValue As Double)
mBlank = False
mDirty = True
mHigh = newValue
End Property

Public Property Get IsUpBar() As Boolean
IsUpBar = (mClose > mOpen)
End Property

Public Property Get IsDownBar() As Boolean
IsDownBar = (mClose <= mOpen)
End Property

Public Property Get lowValue() As Double
lowValue = mLow
End Property

Friend Property Let lowValue(newValue As Double)
mBlank = False
mDirty = True
mLow = newValue
End Property

Public Property Get openInterest() As Long
openInterest = mOpenInterest
End Property

Friend Property Let openInterest(newValue As Long)
mBlank = False
mDirty = True
mOpenInterest = newValue
End Property

Public Property Get openValue() As Double
openValue = mOpen
End Property

Friend Property Let openValue(newValue As Double)
mBlank = False
mDirty = True
mOpen = newValue
End Property

Public Property Get tickVolume() As Long
tickVolume = mTickVolume
End Property

Friend Property Let tickVolume(newValue As Long)
mBlank = False
mDirty = True
mTickVolume = newValue
End Property

Public Property Get volume() As Long
volume = mVolume
End Property

Friend Property Let volume(newValue As Long)
If newValue = 0 Then Exit Property
' NB: don't set mBlank to false, because a bar with only volume
' should be considered blank
mDirty = True
mVolume = newValue
RaiseEvent VolumeChanged
End Property

'================================================================================
' Methods
'================================================================================

Public Sub Tick(ByVal value As Double)
Dim OpenChanged As Boolean
Dim HighChanged As Boolean
Dim LowChanged As Boolean

If mBlank Then
    mBlank = False
    mOpen = value
    OpenChanged = True
End If
If value > mHigh Then
    mHigh = value
    HighChanged = True
End If
If value < mLow Then
    mLow = value
    LowChanged = True
End If
mClose = value
mTickVolume = mTickVolume + 1
    
If OpenChanged Then RaiseEvent OpenChanged
If HighChanged Then RaiseEvent HighChanged
If LowChanged Then RaiseEvent LowChanged
RaiseEvent CloseChanged
RaiseEvent TickVolumeChanged
End Sub

Public Sub TickSize(ByVal size As Long)
volume = volume + size
'RaiseEvent VolumeChanged
End Sub

'================================================================================
' Helper Functions
'================================================================================


