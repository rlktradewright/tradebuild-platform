VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Period"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private mPeriodNumber As Long
Private mTimestamp As Date

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get PeriodNumber() As Long
PeriodNumber = mPeriodNumber
End Property

Friend Property Let PeriodNumber(ByVal value As Long)
mPeriodNumber = value
End Property

Public Property Get Timestamp() As Date
Timestamp = mTimestamp
End Property

Friend Property Let Timestamp(ByVal value As Date)
mTimestamp = value
End Property


'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


