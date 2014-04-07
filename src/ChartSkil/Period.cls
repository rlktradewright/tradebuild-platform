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

Private mPeriodNumber                               As Long
Private mTimestamp                                  As Date
Private mVerticalGridLineType                       As VerticalGridlineTypes
Private mLabels                                     As PeriodLabels

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

Friend Property Get Labels() As PeriodLabels
Labels = mLabels
End Property

Friend Property Let Labels(ByRef Value As PeriodLabels)
mLabels = Value
End Property

Public Property Get PeriodNumber() As Long
PeriodNumber = mPeriodNumber
End Property

Friend Property Let PeriodNumber(ByVal Value As Long)
mPeriodNumber = Value
End Property

Public Property Get Timestamp() As Date
Timestamp = mTimestamp
End Property

Friend Property Let Timestamp(ByVal Value As Date)
mTimestamp = Value
End Property

Friend Property Let VerticalGridLineType(ByVal Value As VerticalGridlineTypes)
mVerticalGridLineType = Value
End Property

Friend Property Get VerticalGridLineType() As VerticalGridlineTypes
VerticalGridLineType = mVerticalGridLineType
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


