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
    next                    As Long
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
mEntries(mFirst).next = 1
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

Friend Property Get capacity() As Long
Const ProcName As String = "capacity"
On Error GoTo Err

capacity = UBound(mEntries) - 1 ' allow 2 for first and last entries

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let capacityIncrement(ByVal value As Long)
Const ProcName As String = "capacityIncrement"
On Error GoTo Err

mCapacityIncrement = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let initialCapacity(ByVal value As Long)
Const ProcName As String = "initialCapacity"
On Error GoTo Err

ReDim Preserve mEntries(value + 1) As ListEntry ' allow 2 for first and last entries
If mCapacityIncrement = 0 Then mCapacityIncrement = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get size() As Long
Const ProcName As String = "size"
On Error GoTo Err

size = mSize

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
    mEntries(add).next = mLast
    mEntries(add).prev = mFirst
    mEntries(mFirst).next = add
    mEntries(mLast).prev = add
Else
    If hint = mFirst Then
        currEntry = mEntries(mFirst).next
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
        mEntries(mEntries(currEntry).next).prev = add
        mEntries(add).next = mEntries(currEntry).next
        mEntries(add).prev = currEntry
        mEntries(currEntry).next = add
    Else
        Do While currEntry <> mLast And _
            mEntries(currEntry).key <= key
            currEntry = mEntries(currEntry).next
        Loop
        mEntries(mEntries(currEntry).prev).next = add
        mEntries(add).prev = mEntries(currEntry).prev
        mEntries(add).next = currEntry
        mEntries(currEntry).prev = add
    End If
End If

mNextFreeIndex = mNextFreeIndex + 1
mSize = mSize + 1

'If mEntries(mFirst).next = 1 Then Stop

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Function firstIndex() As Long
Const ProcName As String = "firstIndex"
On Error GoTo Err

If mSize = 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "List is empty"
firstIndex = mEntries(mFirst).next

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Function itemAt(ByVal index As Long) As Variant
Const ProcName As String = "itemAt"
On Error GoTo Err

If mSize = 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "List is empty"
itemAt = mEntries(index).item

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Function nextIndex(ByVal index As Long) As Long
Const ProcName As String = "nextIndex"
On Error GoTo Err

If mSize = 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "List is empty"
If mEntries(index).next <> mLast Then
    nextIndex = mEntries(index).next
Else
    nextIndex = -1
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Function prevIndex(ByVal index As Long) As Long
Const ProcName As String = "prevIndex"
On Error GoTo Err

If mSize = 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "List is empty"
If mEntries(index).prev <> mFirst Then
    prevIndex = mEntries(index).prev
Else
    prevIndex = -1
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Function remove(ByVal index As Long) As Variant
Const ProcName As String = "remove"
On Error GoTo Err

If index = 0 Or index = 1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "Can't remove head or tail entries"
If mSize = 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "List is empty"
remove = mEntries(index).item
mEntries(mEntries(index).prev).next = mEntries(index).next
mEntries(mEntries(index).next).prev = mEntries(index).prev

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'================================================================================
' Helper Functions
'================================================================================

