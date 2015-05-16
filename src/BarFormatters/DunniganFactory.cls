VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DunniganFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements IBarFormatterFactory

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

Private Const ModuleName                    As String = "DunniganFactory"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' BarFormatterFactory Interface Members
'@================================================================================

Private Property Let IBarFormatterFactory_ConfigurationSection( _
                ByVal RHS As ConfigurationSection)
'Nothing we need to record in the config
Const ProcName As String = "IBarFormatterFactory_ConfigurationSection"
On Error GoTo Err



Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IBarFormatterFactory_CreateBarFormatter( _
                ByVal pBarsFuture As IFuture) As IBarFormatter
Const ProcName As String = "IBarFormatterFactory_CreateBarFormatter"
On Error GoTo Err

Dim bf As New DunniganBarFormatter
bf.BarsFuture = pBarsFuture
Set IBarFormatterFactory_CreateBarFormatter = bf

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IBarFormatterFactory_LoadFromConfig( _
                ByVal config As ConfigurationSection)
'Nothing to load from the config
Const ProcName As String = "IBarFormatterFactory_LoadFromConfig"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


