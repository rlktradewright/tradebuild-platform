VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketOrderHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "BracketOrderHandler"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mBracketOrder                    As BracketOrder
Private mStrategy                                   As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mBracketOrder Event Handlers
'@================================================================================

Private Sub mBracketOrder_Completed()
Const ProcName As String = "mBracketOrder_Completed"
On Error GoTo Err

If TypeOf mStrategy Is IStrategy Then

ElseIf TypeOf mStrategy Is IStopStrategy Then

ElseIf TypeOf mStrategy Is IScaleStrategy Then

End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mBracketOrder_EntryOrderPartiallyFilled()
Const ProcName As String = "mBracketOrder_EntryOrderPartiallyFilled"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mBracketOrder_StopAdjusted()
Const ProcName As String = "mBracketOrder_StopAdjusted"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pStrategy As Object)
Set mBracketOrder = pBracketOrder
Set mStrategy = pStrategy
End Sub


'@================================================================================
' Helper Functions
'@================================================================================




