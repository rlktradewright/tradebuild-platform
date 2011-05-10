VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IExecutionReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "IExecutionReport"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get AccountIdentifier() As String
End Property

Public Property Get Action() As OrderActions
End Property

Public Property Get Id() As String
End Property

Public Property Get FillingExchange() As String
End Property

Public Property Get FillTime() As Date
End Property

Public Property Get IsLiquidation() As Boolean
End Property

Public Property Get BrokerId() As String
End Property

Public Property Get OrderId() As String
End Property

Public Property Get Price() As Double
End Property

Public Property Get Quantity() As Long
End Property

Public Property Get SecurityName() As String
End Property

Public Property Get TimezoneName() As String
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




