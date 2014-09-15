VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITickStream"
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

Private Const ModuleName                            As String = "ITickStream"

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

Public Property Get ClockFuture() As IFuture
End Property

Public Property Get ContractFuture() As IFuture
End Property

Public Property Get State() As TickStreamStates
End Property

Public Property Get StreamId() As Long
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddGenericTickListener(ByVal pListener As IGenericTickListener)
End Sub

Public Sub AddStateChangeListener(ByVal pListener As StateChangeListener)
End Sub

Public Sub Finish()
End Sub

Public Sub RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
End Sub

Public Sub RemoveStateChangeListener(ByVal pListener As StateChangeListener)
End Sub

Public Sub Start()
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




