VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITickfileReader"
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

Private Const ModuleName                            As String = "ITickfileReader"

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

'Public Property Get FirstTickFileTime() As Date
'End Property

Public Property Get ReplayProgressPercent() As Long
End Property

Public Property Get ReplayProgressTicks() As Long
End Property

Public Property Get TickfileSizeBytes() As Long
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub CloseInputFile()
End Sub

Public Function FireNextTick() As Boolean
End Function

Public Function PrepareTickfile(ByVal pTickfileSpec As ITickfileSpecifier) As Boolean
End Function

Public Function Supports( _
                ByVal Capabilities As Long, _
                Optional ByVal FormatIdentifier As String) As Boolean
End Function

'@================================================================================
' Helper Functions
'@================================================================================




