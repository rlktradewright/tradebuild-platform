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

Public Sub gAddExistingKey(ByVal pKey As String)
Const ProcName As String = "gAddExistingKey"
On Error GoTo Err

allocatedKeys.Add Nothing, pKey

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gNextKey() As String
Const ProcName As String = "gNextKey"
On Error GoTo Err


Do
    Dim lTimestamp As Double: lTimestamp = GetTimestamp
    gNextKey = Hex((CLng(Int(lTimestamp)) Mod 100) * 10000000 + CLng(Right$(Int(lTimestamp * 886400000), 8)))
Loop While allocatedKeys.Contains(gNextKey)

allocatedKeys.Add Nothing, gNextKey

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function allocatedKeys() As SortedDictionary
Const ProcName As String = "allocatedKeys"
On Error GoTo Err

Static sKeys As SortedDictionary
If sKeys Is Nothing Then Set sKeys = CreateSortedDictionary(KeyTypeString)
Set allocatedKeys = sKeys

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



