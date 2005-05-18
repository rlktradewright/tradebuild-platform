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
Private mBackColor As Long

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

Public Property Get backColor() As Long
backColor = mBackColor
End Property

Public Property Let backColor(ByVal value As Long)
mBackColor = value
End Property

Public Property Get periodNumber() As Long
periodNumber = mPeriodNumber
End Property

Friend Property Let periodNumber(ByVal value As Long)
mPeriodNumber = value
End Property

Public Property Get timestamp() As Date
timestamp = mTimestamp
End Property

Friend Property Let timestamp(ByVal value As Date)
mTimestamp = value
End Property


'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


