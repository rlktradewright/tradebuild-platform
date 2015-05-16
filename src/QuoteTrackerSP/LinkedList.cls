VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinkedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "LinkedList"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type ListEntry
    Next                    As Long
    prev                    As Long
    key                     As Variant
    item                    As Variant
End Type


'================================================================================
' Member variables
'================================================================================

Private mEntries() As ListEntry

Private mNextFreeIndex As Long
Private mSize As Long
Private mFirst As Long
Private mLast As Long

Private mCapacityIncrement As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
ReDim mEntries(2) As ListEntry
mFirst = 0
mLast = 1
mEntries(mFirst).Next = 1
mEntries(mLast).prev = 0
mEntries(mLast).key = Null
mNextFreeIndex = 2
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Get Capacity() As Long
Const ProcName As String = "Capacity"
On Error GoTo Err

Capacity = UBound(mEntries) - 1 ' allow 2 for first and last entries

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let CapacityIncrement(ByVal value As Long)
Const ProcName As String = "CapacityIncrement"
On Error GoTo Err

mCapacityIncrement = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let InitialCapacity(ByVal value As Long)
Const ProcName As String = "InitialCapacity"
On Error GoTo Err

ReDim Preserve mEntries(value + 1) As ListEntry ' allow 2 for first and last entries
If mCapacityIncrement = 0 Then mCapacityIncrement = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Get Size() As Long
Const ProcName As String = "Size"
On Error GoTo Err

Size = mSize

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Friend Function add(ByVal Entry As Variant, key As Variant, Optional ByVal hint As Long) As Long
Dim currEntry As Long

Const ProcName As String = "add"
On Error GoTo Err

If hint < 0 Or hint >= mNextFreeIndex Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid value for hint argument"
End If

If mNextFreeIndex > UBound(mEntries) Then
    ReDim Preserve mEntries(UBound(mEntries) + mCapacityIncrement) As ListEntry
End If

mEntries(mNextFreeIndex).item = Entry
mEntries(mNextFreeIndex).key = key

add = mNextFreeIndex

If mSize = 0 Then
    ' no entries yet
    mEntries(add).Next = mLast
    mEntries(add).prev = mFirst
    mEntries(mFirst).Next = add
    mEntries(mLast).prev = add
Else
    If hint = mFirst Then
        currEntry = mEntries(mFirst).Next
    ElseIf hint = mLast Then
        currEntry = mEntries(mLast).prev
    Else
        currEntry = hint
    End If
    
    If mEntries(currEntry).key > key Then
        Do While currEntry <> mFirst And _
            mEntries(currEntry).key > key
            currEntry = mEntries(currEntry).prev
        Loop
        mEntries(mEntries(currEntry).Next).prev = add
        mEntries(add).Next = mEntries(currEntry).Next
        mEntries(add).prev = currEntry
        mEntries(currEntry).Next = add
    Else
        Do While currEntry <> mLast And _
            mEntries(currEntry).key <= key
            currEntry = mEntries(currEntry).Next
        Loop
        mEntries(mEntries(currEntry).prev).Next = add
        mEntries(add).prev = mEntries(currEntry).prev
        mEntries(add).Next = currEntry
        mEntries(currEntry).prev = add
    End If
End If

mNextFreeIndex = mNextFreeIndex + 1
mSize = mSize + 1

'If mEntries(mFirst).Next = 1 Then Stop

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Function firstIndex() As Long
Const ProcName As String = "firstIndex"
On Error GoTo Err

If mSize = 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "List is empty"
firstIndex = mEntries(mFirst).Next

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Function itemAt(ByVal Index As Long) As Variant
Const ProcName As String = "itemAt"
On Error GoTo Err

If mSize = 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "List is empty"
itemAt = mEntries(Index).item

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Function lastIndex() As Long
Const ProcName As String = "lastIndex"
On Error GoTo Err

If mSize = 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "List is empty"
lastIndex = mEntries(mLast).prev

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Function nextIndex(ByVal Index As Long) As Long
Const ProcName As String = "nextIndex"
On Error GoTo Err

If mSize = 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "List is empty"
If mEntries(Index).Next <> mLast Then
    nextIndex = mEntries(Index).Next
Else
    nextIndex = -1
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Function prevIndex(ByVal Index As Long) As Long
Const ProcName As String = "prevIndex"
On Error GoTo Err

If mSize = 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "List is empty"
If mEntries(Index).prev <> mFirst Then
    prevIndex = mEntries(Index).prev
Else
    prevIndex = -1
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Function Remove(ByVal Index As Long) As Variant
Const ProcName As String = "Remove"
On Error GoTo Err

If Index = 0 Or Index = 1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "Can't Remove head or tail entries"
If mSize = 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "List is empty"
Remove = mEntries(Index).item
mEntries(mEntries(Index).prev).Next = mEntries(Index).Next
mEntries(mEntries(Index).Next).prev = mEntries(Index).prev

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

'================================================================================
' Helper Functions
'================================================================================

