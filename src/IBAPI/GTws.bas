Attribute VB_Name = "GTws"
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

Private Const ModuleName                            As String = "GTws"

'@================================================================================
' Member variables
'@================================================================================

Private mTwsCollection                              As New EnumerableCollection

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

'@================================================================================
' Methods
'@================================================================================

Public Function gGetTws( _
                ByVal pServer As String, _
                ByVal pPort As Long) As Tws
Const ProcName As String = "gGetTws"
On Error GoTo Err

If pServer = "" Then pServer = "127.0.0.1"

If Not mTwsCollection.Contains(generateTwsKey(pServer, pPort)) Then
    Set gGetTws = New Tws
    mTwsCollection.Add New Tws, generateTwsKey(pServer, pPort)
    
    gGetTws.Initialise pServer, pPort
Else
    Set gGetTws = mTwsCollection(generateTwsKey(pServer, pPort))
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub gReleaseTws( _
                ByVal pTws As Tws)
Const ProcName As String = "gReleaseTws"
On Error GoTo Err

If mTwsCollection.Contains(generateTwsKey(pTws.Server, pTws.Port)) Then mTwsCollection.Remove generateTwsKey(pTws.Server, pTws.Port)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function generateTwsKey( _
                ByVal pServer As String, _
                ByVal pPort As Long) As String
generateTwsKey = pServer & vbNullChar & pPort
End Function


