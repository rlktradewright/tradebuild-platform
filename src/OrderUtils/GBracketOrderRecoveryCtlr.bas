Attribute VB_Name = "GBracketOrderRecoveryCtlr"
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

Private Const ModuleName                            As String = "GBracketOrderRecoveryCtlr"

'@================================================================================
' Member variables
'@================================================================================

Private mSessionName                                As String
Private mRecoveryCOntrollers                        As EnumerableCollection

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

Public Function gGetBracketOrderRecoveryController(ByVal pScopeName As String) As BracketOrderRecoveryCtlr
Const ProcName As String = "gGetBracketOrderRecoveryController"
On Error GoTo Err

Assert mSessionName <> "", "An order recovery session has not yet been started"

If mRecoveryCOntrollers.Contains(pScopeName) Then
    Set gGetBracketOrderRecoveryController = mRecoveryCOntrollers(pScopeName)
Else
    Set gGetBracketOrderRecoveryController = New BracketOrderRecoveryCtlr
    mRecoveryCOntrollers.Add gGetBracketOrderRecoveryController, pScopeName
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Property Get gSessionName() As String
gSessionName = mSessionName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub gStartOrderRecoverySession(ByVal pSessionName As String)
Const ProcName As String = "gStartOrderRecoverySession"
On Error GoTo Err

Assert pSessionName <> "", "Invalid session name"
Assert pSessionName <> mSessionName, "This session has already been started"

mSessionName = pSessionName
Set mRecoveryCOntrollers = New EnumerableCollection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




