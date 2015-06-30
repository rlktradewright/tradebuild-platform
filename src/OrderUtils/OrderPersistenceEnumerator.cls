VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPersistenceEnumerator"
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

Implements IEnumerable

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

Private Const ModuleName                            As String = "OrderPersistenceEnumerator"

'@================================================================================
' Member variables
'@================================================================================

Private mTS                                         As TextStream

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IEnumerable Interface Members
'@================================================================================

Private Property Get IEnumerable_Enumerator() As Enumerator
Const ProcName As String = "IEnumerable_Enumerator"
On Error GoTo Err

Set IEnumerable_Enumerator = getEnumerator(Me)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IEnumerable_NextItem(Item As Variant, cookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_NextItem"
On Error GoTo Err

Dim lLine As String
If mTS Is Nothing Then
    IEnumerable_NextItem = False
ElseIf getNextLine(lLine) Then
    Item = lLine
    IEnumerable_NextItem = True
Else
    IEnumerable_NextItem = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IEnumerable_Remove(cookie As Variant)
Const ProcName As String = "IEnumerable_Remove"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Remove not supported"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IEnumerable_Skip(ByVal numToSkip As Long, cookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_Skip"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Skip not supported"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IEnumerable_UpdateSequence() As Long

End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Function getEnumerator() As Enumerator
Const ProcName As String = "GetEnumerator"
On Error GoTo Err

Set getEnumerator = CreateEnumerator(Me)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise(ByVal pTS As TextStream)
Set mTS = pTS
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getNextLine(ByRef result As String) As Boolean
Const ProcName As String = "getNextLine"
On Error GoTo Err

Do While Not mTS.AtEndOfStream
    result = mTS.ReadLine
    If result <> "" And Left$(result, 2) <> "//" Then
        getNextLine = True
        Exit Do
    End If
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




