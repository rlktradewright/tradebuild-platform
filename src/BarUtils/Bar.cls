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
Attribute VB_Creatable = True
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

Implements IStringable

'@================================================================================
' Events
'@================================================================================

Event Changed( _
                ByRef ev As ChangeEventData)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "Bar"

'@================================================================================
' Enums
'@================================================================================

Private Enum BarChanges
    BarChangeOpenChanged = 1&
    BarChangeHighChanged = 2&
    BarChangeLowChanged = 4&
    BarChangeCloseChanged = 8&
    BarChangeVolumeChanged = 16&
    BarChangeOpenInterestChanged = 32&
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mBarChanges As BarChanges

Private mBarNumber As Long

Private mOpen As Double

Private mHigh As Double

Private mLow As Double

Private mClose As Double

Private mVolume As Long

Private mOpenInterest As Long

Private mTickVolume As Long

Private mTimestamp As Date

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mOpen = MaxDouble
mLow = MaxDouble
mHigh = MinDouble
mClose = MaxDouble
End Sub

'@================================================================================
' IStringable Interface Members
'@================================================================================

Private Function IStringable_ToString() As String
IStringable_ToString = ToString
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get BarChanged() As Boolean
BarChanged = (mBarChanges <> 0)
End Property

Friend Property Let BarNumber(ByVal Value As Long)
mBarNumber = Value
End Property

Public Property Get BarNumber() As Long
BarNumber = mBarNumber
End Property

Public Property Get Blank() As Boolean
Blank = ((mOpen = MaxDouble) And (mLow = MaxDouble) And (mHigh = MinDouble) And (mClose = MaxDouble))
End Property

Public Property Get CloseValue() As Double
CloseValue = mClose
End Property

Public Property Get CloseChanged() As Boolean
CloseChanged = CBool(mBarChanges And BarChangeCloseChanged)
End Property

Public Property Get Timestamp() As Date
Timestamp = mTimestamp
End Property

Friend Property Let Timestamp(ByVal newvalue As Date)
mTimestamp = newvalue
End Property

Public Property Get Gain() As Double
Gain = mClose - mOpen
If Gain < 0 Then Gain = 0
End Property

Public Property Get HighValue() As Double
HighValue = mHigh
End Property

Public Property Get HighChanged() As Boolean
HighChanged = CBool(mBarChanges And BarChangeHighChanged)
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

Public Property Get Loss() As Double
Loss = mOpen - mClose
If Loss < 0 Then Loss = 0
End Property

Public Property Get LowValue() As Double
LowValue = mLow
End Property

Public Property Get LowChanged() As Boolean
LowChanged = CBool(mBarChanges And BarChangeLowChanged)
End Property

Public Property Get OHLC4() As Double
OHLC4 = (mOpen + mHigh + mLow + mClose) / 4
End Property

Public Property Get OpenChanged() As Boolean
OpenChanged = CBool(mBarChanges And BarChangeOpenChanged)
End Property

Public Property Get OpenInterest() As Long
OpenInterest = mOpenInterest
End Property

Public Property Get OpenInterestChanged() As Long
OpenInterestChanged = CBool(mBarChanges And BarChangeOpenInterestChanged)
End Property

Public Property Get OpenValue() As Double
OpenValue = mOpen
End Property

Public Property Get TickVolume() As Long
TickVolume = mTickVolume
End Property

Public Property Get Volume() As Long
Volume = mVolume
End Property

Public Property Get VolumeChanged() As Long
VolumeChanged = CBool(mBarChanges And BarChangeVolumeChanged)
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub IncrementTickVolume( _
                ByVal increment As Long)
Const ProcName As String = "IncrementTickVolume"
On Error GoTo Err

mTickVolume = mTickVolume + increment
fireChange BarTickVolumeChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise( _
                ByVal Timestamp As Date, _
                ByVal OpenValue As Double, _
                ByVal HighValue As Double, _
                ByVal LowValue As Double, _
                ByVal CloseValue As Double, _
                Optional ByVal Volume As Long, _
                Optional ByVal TickVolume As Long, _
                Optional ByVal OpenInterest As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

mTimestamp = Timestamp
mOpen = OpenValue
mHigh = HighValue
mLow = LowValue
mClose = CloseValue
mVolume = Volume
mTickVolume = TickVolume
mOpenInterest = OpenInterest

mBarChanges = (BarChangeOpenChanged Or BarChangeHighChanged Or BarChangeLowChanged Or BarChangeCloseChanged)
If mVolume <> 0 Then mBarChanges = (mBarChanges Or BarChangeVolumeChanged)
If mOpenInterest <> 0 Then mBarChanges = (mBarChanges Or BarChangeOpenInterestChanged)

fireChange BarPriceChanged Or _
        IIf(mVolume <> 0, BarVolumeChanged, 0) Or _
        IIf(mOpenInterest <> 0, BarOpenInterestChanged, 0) Or _
        IIf(mTickVolume <> 0, BarTickVolumeChanged, 0)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyOpenInterest( _
                ByVal OpenInterest As Long)
Const ProcName As String = "NotifyOpenInterest"
On Error GoTo Err

' note that we don't set mBlank to false because the Bar is blank until
' the first Value is notified
mOpenInterest = OpenInterest
mBarChanges = BarChangeOpenInterestChanged
fireChange BarOpenInterestChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyValue(ByVal Value As Double)
Const ProcName As String = "NotifyValue"
On Error GoTo Err

mBarChanges = 0

If Blank Then
    mOpen = Value
    mBarChanges = (mBarChanges Or BarChangeOpenChanged)
End If

If Value > mHigh Then
    mHigh = Value
    mBarChanges = (mBarChanges Or BarChangeHighChanged)
End If

If Value < mLow Then
    mLow = Value
    mBarChanges = (mBarChanges Or BarChangeLowChanged)
End If

If Value <> mClose Then
    mClose = Value
    mBarChanges = (mBarChanges Or BarChangeCloseChanged)
End If

mTickVolume = mTickVolume + 1
If mBarChanges Then fireChange BarPriceChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyVolume( _
                ByVal Volume As Long)
Const ProcName As String = "NotifyVolume"
On Error GoTo Err

' note that we don't set mBlank to false because the Bar is blank until
' the first Value is notified
mVolume = Volume
mBarChanges = BarChangeVolumeChanged
fireChange BarVolumeChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetBarChanged()
mBarChanges = (BarChangeOpenChanged Or BarChangeHighChanged Or BarChangeLowChanged Or BarChangeCloseChanged Or BarChangeVolumeChanged Or BarChangeOpenInterestChanged)
End Sub

Public Function ToString() As String
ToString = "Bar{Time:" & FormatTimestamp(mTimestamp, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & _
            ", Open:" & mOpen & _
            ", High:" & mHigh & _
            ", Low:" & mLow & _
            ", Close:" & mClose & _
            ", Volume:" & mVolume & _
            ", TickVolume:" & mTickVolume & _
            ", OpenInterest:" & mOpenInterest & "}"
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal changeType As BarChangeTypes)
Const ProcName As String = "fireChange"
On Error GoTo Err

Dim ev As ChangeEventData
Set ev.Source = Me
ev.changeType = changeType
RaiseEvent Changed(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




