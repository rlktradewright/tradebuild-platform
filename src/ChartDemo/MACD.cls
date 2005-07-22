VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Private mMAShort As ExponentialMovingAverage
Private mMALong As ExponentialMovingAverage
Private mMASignal As ExponentialMovingAverage


'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let LongPeriods(ByVal value As Long)
Set mMALong = New ExponentialMovingAverage
mMALong.periods = value
End Property

Public Property Get MACDHistValue() As Variant
If Not IsEmpty(MACDValue) And Not IsEmpty(MACDSignalValue) Then
    MACDHistValue = MACDValue - MACDSignalValue
End If
End Property

Public Property Get MACDSignalValue() As Variant
MACDSignalValue = mMASignal.maValue
End Property

Public Property Get MACDValue() As Variant
If Not IsEmpty(mMAShort.maValue) And Not IsEmpty(mMALong.maValue) Then
    MACDValue = mMAShort.maValue - mMALong.maValue
End If
End Property

Public Property Let ShortPeriods(ByVal value As Long)
Set mMAShort = New ExponentialMovingAverage
mMAShort.periods = value
End Property

Public Property Let SignalPeriods(ByVal value As Long)
Set mMASignal = New ExponentialMovingAverage
mMASignal.periods = value
End Property

'================================================================================
' Methods
'================================================================================

Public Sub datavalue(ByVal value As Double)
mMAShort.datavalue value
mMALong.datavalue value
If Not IsEmpty(mMAShort.maValue) And Not IsEmpty(mMALong.maValue) Then
    mMASignal.datavalue mMAShort.maValue - mMALong.maValue
End If
End Sub

Public Sub newPeriod()
If Not IsEmpty(mMAShort.maValue) And Not IsEmpty(mMALong.maValue) Then
    mMASignal.newPeriod
End If
mMAShort.newPeriod
mMALong.newPeriod
End Sub

'================================================================================
' Helper Functions
'================================================================================

