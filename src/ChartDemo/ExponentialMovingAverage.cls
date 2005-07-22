VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExponentialMovingAverage"
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

Private mExpFactor As Double
Private mPrevPeriodMa As Double
Private mMa As Double
Private mPeriods As Long
Private mPeriodCount As Long
Private mFirstValueReceived As Boolean

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

Public Property Get maValue() As Variant
If mPeriodCount < mPeriods Then
    maValue = Empty
Else
    maValue = mMa
End If
End Property

Public Property Let periods(ByVal value As Long)
mPeriods = value
mExpFactor = 2 / (value + 1)
End Property

'================================================================================
' Methods
'================================================================================

Public Sub datavalue(ByVal value As Double)
If Not mFirstValueReceived Then
    mPrevPeriodMa = value
    mFirstValueReceived = True
End If
mMa = (mExpFactor * value) + mPrevPeriodMa * (1 - mExpFactor)
End Sub

Public Sub newPeriod()
mPeriodCount = mPeriodCount + 1
mPrevPeriodMa = mMa
End Sub

'================================================================================
' Helper Functions
'================================================================================

