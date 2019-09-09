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

If allocatedIds.Contains(pKey) Then Exit Sub
allocatedIds.Add Nothing, pKey

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gNextId() As String
Const ProcName As String = "gNextId"
On Error GoTo Err

Const MillisecsPerDay As Long = 24& * 60& * 60& * 1000&
Const OneMillisec As Double = 1# / MillisecsPerDay

Dim lTimestamp As Double: lTimestamp = GetTimestamp

Do
    gNextId = CStr(Hex(&H10000000 + (CLng(Int(lTimestamp)) Mod 21) * MillisecsPerDay + CLng((lTimestamp - Int(lTimestamp)) * MillisecsPerDay)))
    gNextId = Right$(gNextId, 4) & Left$(gNextId, 4)
    If Not allocatedIds.Contains(gNextId) Then Exit Do
    ' note that hitting a previously used id is most likely to
    ' occur when calling this function very rapidly
    lTimestamp = lTimestamp + OneMillisec
Loop

allocatedIds.Add Nothing, gNextId

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function allocatedIds() As SortedDictionary
Const ProcName As String = "allocatedIds"
On Error GoTo Err

Static sKeys As SortedDictionary
If sKeys Is Nothing Then Set sKeys = CreateSortedDictionary(KeyTypeString)
Set allocatedIds = sKeys

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



