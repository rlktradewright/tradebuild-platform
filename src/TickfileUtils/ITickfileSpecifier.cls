VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITickfileSpecifier"
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

Private Const ModuleName                            As String = "ITickfileSpecifier"

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

Public Property Let Contract(ByVal RHS As IContract)
End Property

Public Property Get Contract() As IContract
End Property

Public Property Let EntireSession(ByVal RHS As Boolean)
End Property

Public Property Get EntireSession() As Boolean
End Property

Public Property Let Filename(ByVal RHS As String)
End Property

Public Property Get Filename() As String
End Property

Public Property Let FromDate(ByVal RHS As Date)
End Property

Public Property Get FromDate() As Date
End Property

Public Property Let TickfileFormatID(ByVal RHS As String)
End Property

Public Property Get TickfileFormatID() As String
End Property

Public Property Let ToDate(ByVal RHS As Date)
End Property

Public Property Get ToDate() As Date
End Property

Public Property Let UseFullSession(ByVal RHS As Boolean)
End Property

Public Property Get UseFullSession() As Boolean
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function ToString() As String
End Function

'@================================================================================
' Helper Functions
'@================================================================================




