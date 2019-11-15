VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NextCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Private Const ModuleName                            As String = "NextCommands"

'@================================================================================
' Member variables
'@================================================================================

Private mNextCommandLists()                         As CommandList

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

Public Property Get ValidCommandNames() As String
Const ProcName As String = "ValidCommandNames"
On Error GoTo Err

Dim s As String
Dim i As Long
For i = 0 To UBound(mNextCommandLists)
    If i <> 0 Then s = s & " "
    s = s & mNextCommandLists(i).CommandNames
Next
ValidCommandNames = s

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function IsCommandValid(ByVal pCommand As Command) As Boolean
Const ProcName As String = "IsCommandValid"
On Error GoTo Err

If pCommand Is Nothing Then Exit Function

Dim i As Long
For i = 0 To UBound(mNextCommandLists)
    If mNextCommandLists(i).IsValidCommandName(pCommand.Name) Then
        IsCommandValid = True
        Exit Function
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub SetValidNextCommandLists(ByRef pCommandLists() As CommandList)
Const ProcName As String = "gSetValidNextCommands"
On Error GoTo Err

mNextCommandLists = pCommandLists
'ReDim mNextCommandLists(UBound(pCommandLists)) As CommandList
'Dim i As Long
'For i = 0 To UBound(Values)
'    Dim lCommandList As CommandList
'    Set lCommandList = pCommandLists(i)
'    Set mNextCommandLists(i) = lCommandList
'Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




