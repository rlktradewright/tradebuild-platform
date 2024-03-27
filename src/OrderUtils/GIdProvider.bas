Attribute VB_Name = "GIdProvider"
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

Private Const ModuleName                            As String = "GIdProvider"

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

'@================================================================================
' Methods
'@================================================================================

Public Sub gAddExistingId(ByVal pKey As String)
Const ProcName As String = "gAddExistingId"
On Error GoTo Err

init

If allocatedIds.Contains(pKey) Then Exit Sub
allocatedIds.Add Nothing, pKey

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gNextId() As String
Const ProcName As String = "gNextId"
On Error GoTo Err

init

Do
    gNextId = Hex(&H80000000 + Rnd * (&H7FFFFFF0))
    
    If Not allocatedIds.Contains(gNextId) Then Exit Do
Loop

allocatedIds.Add Nothing, gNextId

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function allocatedIds() As SortedDictionary
Const ProcName As String = "allocatedIds"
On Error GoTo Err

Static sKeys As SortedDictionary
If sKeys Is Nothing Then
    Set sKeys = CreateSortedDictionary(KeyTypeString)
End If
Set allocatedIds = sKeys

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub init()
Static sInitialised As Boolean
If sInitialised Then Exit Sub
sInitialised = True
'Randomize Right$(Format(GetTimestamp, "0.0000000000"), 6)
Randomize
End Sub

