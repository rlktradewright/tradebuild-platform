VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Queue"
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

Private Const ModuleName                As String = "Queue"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type QueueEntry
    item                As Variant
    itemKey             As String
    itemRemoved         As Boolean
End Type

'================================================================================
' Member variables
'================================================================================

Private mInitialised As Boolean

Private mEntries() As QueueEntry
Private mSize As Long
Private mFirstEntry As Long
Private mNextEntry As Long

Private mCapacityIncrement As Long

'================================================================================
' Class Event Handlers
'================================================================================

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

Capacity = UBound(mEntries) + 1

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let CapacityIncrement(ByVal value As Long)
Const ProcName As String = "CapacityIncrement"
On Error GoTo Err

If value < 1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Capaciy increment must be >0"
mCapacityIncrement = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Get FrontKey() As String
Const ProcName As String = "FrontKey"
On Error GoTo Err

If mSize = 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Queue empty"

Do While mEntries(mFirstEntry).itemRemoved
    mFirstEntry = (mFirstEntry + 1) Mod (UBound(mEntries) + 1)
Loop

FrontKey = mEntries(mFirstEntry).itemKey

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let InitialCapacity(ByVal value As Long)
Const ProcName As String = "InitialCapacity"
On Error GoTo Err

If value < 1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Initial capaciy must be >0"
ReDim mEntries(value - 1) As QueueEntry
If mCapacityIncrement = 0 Then mCapacityIncrement = value
mInitialised = True

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

Friend Function AddToEnd(ByVal Entry As Variant) As String
Const ProcName As String = "AddToEnd"
Dim failpoint As String
On Error GoTo Err

If Not mInitialised Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Queue not initialised"
If mNextEntry = mFirstEntry And mSize <> 0 Then Resize
AddToEnd = GenerateGUIDString
mEntries(mNextEntry).item = Entry
mEntries(mNextEntry).itemKey = AddToEnd
mNextEntry = (mNextEntry + 1) Mod (UBound(mEntries) + 1)
mSize = mSize + 1

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub Remove(ByVal itemRef As Variant)
Const ProcName As String = "Remove"
Dim i As Long
Dim maxIndex As Long

Dim failpoint As String
On Error GoTo Err

If mSize = 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Queue empty"

If mFirstEntry < mNextEntry Then
    maxIndex = mNextEntry - 1
Else
    maxIndex = UBound(mEntries) + mNextEntry
End If

If VarType(itemRef) = vbString Then
    For i = mFirstEntry To maxIndex
        If mEntries(i Mod (UBound(mEntries) - 1)).itemKey = itemRef Then
            mEntries(i Mod (UBound(mEntries) - 1)).itemRemoved = True
            mEntries(i Mod (UBound(mEntries) - 1)).itemKey = ""
            mEntries(i Mod (UBound(mEntries) - 1)).item = Empty
            Exit For
        End If
    Next
ElseIf VarType(itemRef) = vbObject Then
    For i = mFirstEntry To maxIndex
        If mEntries(i Mod (UBound(mEntries) - 1)).item Is itemRef Then
            mEntries(i Mod (UBound(mEntries) - 1)).itemRemoved = True
            mEntries(i Mod (UBound(mEntries) - 1)).itemKey = ""
            mEntries(i Mod (UBound(mEntries) - 1)).item = Empty
            Exit For
        End If
    Next
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Argument type must be string or object"
End If

If i > maxIndex Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Specified item not in queue"

mSize = mSize - 1

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Function RemoveFromFront() As Variant
Const ProcName As String = "RemoveFromFront"
Dim failpoint As String
On Error GoTo Err

If mSize = 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Queue empty"

Do While mEntries(mFirstEntry).itemRemoved
    mFirstEntry = (mFirstEntry + 1) Mod (UBound(mEntries) + 1)
Loop

failpoint = "100"
RemoveFromFront = mEntries(mFirstEntry).item

failpoint = "200"
mEntries(mFirstEntry).item = Empty
mEntries(mFirstEntry).itemKey = ""
mFirstEntry = (mFirstEntry + 1) Mod (UBound(mEntries) + 1)
mSize = mSize - 1

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub Resize()
Const ProcName As String = "Resize"
Dim i As Long
Dim currUbound As Long

currUbound = UBound(mEntries)

ReDim Preserve mEntries(currUbound + mCapacityIncrement) As QueueEntry

If mNextEntry <= mFirstEntry Then
    For i = 0 To mNextEntry - 1
        mEntries((currUbound + i + 1) Mod (UBound(mEntries) + 1)) = mEntries(i)
        mEntries(i).item = Empty
        mEntries(i).itemKey = ""
        mEntries(i).itemRemoved = False
    Next
    mNextEntry = (mNextEntry + currUbound + 1) Mod (UBound(mEntries) + 1)
End If
End Sub




