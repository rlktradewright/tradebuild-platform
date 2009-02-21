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

Event Changed( _
                ByRef ev As ChangeEvent)

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "TimeframeUtils26"
Private Const ModuleName                    As String = "Bar"

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

Private mVolume As Long
Private mVolumeChanged As Boolean

Private mOpenInterest As Long
Private mOpenInterestChanged As Boolean

Private mTickVolume As Long
Private mTickVolumeChanged As Boolean

Private mTimestamp As Date
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

Public Property Get timestamp() As Date
timestamp = mTimestamp
End Property

Friend Property Let timestamp(ByVal newvalue As Date)
mTimestamp = newvalue
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

Public Property Get HL2() As Double
HL2 = (mHigh + mLow) / 2
End Property

Public Property Get HLC3() As Double
HLC3 = (mHigh + mLow + mClose) / 3
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

Public Property Get OHLC4() As Double
OHLC4 = (mOpen + mHigh + mLow + mClose) / 4
End Property

Public Property Get openChanged() As Boolean
openChanged = mOpenChanged
End Property

Public Property Get openInterest() As Long
openInterest = mOpenInterest
End Property

Public Property Get openInterestChanged() As Long
openInterestChanged = mOpenInterestChanged
End Property

Public Property Get openValue() As Double
openValue = mOpen
End Property

Public Property Get tickVolume() As Long
tickVolume = mTickVolume
End Property

Public Property Get volume() As Long
volume = mVolume
End Property

Public Property Get volumeChanged() As Long
volumeChanged = mVolumeChanged
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub incrementTickVolume( _
                ByVal increment As Long)
mTickVolume = mTickVolume + increment
fireChange BarTickVolumeChanged
End Sub

Friend Sub initialise( _
                ByVal timestamp As Date, _
                ByVal openValue As Double, _
                ByVal highValue As Double, _
                ByVal lowValue As Double, _
                ByVal closeValue As Double, _
                Optional ByVal volume As Long, _
                Optional ByVal tickVolume As Long, _
                Optional ByVal openInterest As Long)
mTimestamp = timestamp
mOpen = openValue
mHigh = highValue
mLow = lowValue
mClose = closeValue
mVolume = volume
mTickVolume = tickVolume
mOpenInterest = openInterest

mBarChanged = True
mOpenChanged = True
mHighChanged = True
mLowChanged = True
mCloseChanged = True
If mVolume <> 0 Then mVolumeChanged = True

If mOpenInterest <> 0 Then mOpenInterestChanged = True
If mTickVolume <> 0 Then mTickVolumeChanged = True

fireChange BarPriceChanged
If mVolume <> 0 Then fireChange BarVolumeChanged
If mOpenInterest <> 0 Then fireChange BarOpenInterestChanged
If mTickVolume <> 0 Then fireChange BarTickVolumeChanged
End Sub

Friend Sub notifyOpenInterest( _
                ByVal openInterest As Long)
clean
' note that we don't set mBlank to false because the bar is blank until
' the first value is notified
mOpenInterest = openInterest
mBarChanged = True
mOpenInterestChanged = True
fireChange BarOpenInterestChanged
End Sub

Friend Sub notifyValue(ByVal value As Double)
clean

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
If mBarChanged Then fireChange BarPriceChanged
End Sub

Friend Sub notifyVolume( _
                ByVal volume As Long)
clean
' note that we don't set mBlank to false because the bar is blank until
' the first value is notified
mVolume = volume
mBarChanged = True
mVolumeChanged = True
fireChange BarVolumeChanged
End Sub

Public Sub resetChangeIndicators()
clean
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub clean()
mBarChanged = False
mOpenChanged = False
mHighChanged = False
mLowChanged = False
mCloseChanged = False
mVolumeChanged = False
mOpenInterestChanged = False
End Sub

Private Sub fireChange( _
                ByVal changeType As BarChangeTypes)
Dim ev As ChangeEvent

Set ev.source = Me
ev.changeType = changeType
RaiseEvent Changed(ev)
End Sub




