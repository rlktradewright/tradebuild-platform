VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parameter"
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
' Interfaces
'================================================================================

Implements TradeBuildSP.IParameter

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

Private mParameter As Parameters2.parameter

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IParameter Interface Members
'================================================================================

Private Property Let IParameter_name(ByVal RHS As String)
name = RHS
End Property

Private Property Get IParameter_name() As String
IParameter_name = name
End Property

Private Property Let IParameter_Value(ByVal RHS As String)
Value = RHS
End Property

Private Property Get IParameter_Value() As String
IParameter_Value = Value
End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let name(ByVal val As String)
mParameter.name = val
End Property

Public Property Get name() As String
name = mParameter.name
End Property

Friend Property Let parameter(ByVal Value As Parameters2.parameter)
Set mParameter = Value
End Property

Friend Property Get parameter() As Parameters2.parameter
Set parameter = mParameter
End Property

Public Property Let Value(ByVal val As String)
mParameter.Value = val
End Property

Public Property Get Value() As String
Value = mParameter.Value
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


