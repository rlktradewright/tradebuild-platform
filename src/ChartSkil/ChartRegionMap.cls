VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartRegionMap"
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

Private Type RegionTableEntry
    prev            As Long
    next            As Long
    Region          As ChartRegion
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ChartRegionMap"

'@================================================================================
' Member variables
'@================================================================================

Private mRegions()                                  As RegionTableEntry
Private mRegionsIndex                               As Long

Private mFirstIndex                                 As Long
Private mLastIndex                                  As Long

Private mUpdateSequence                             As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mRegions(1) As RegionTableEntry
mRegionsIndex = 1

mFirstIndex = 0
mLastIndex = 1

nextEntryIndex(mFirstIndex) = mLastIndex
previousEntryIndex(mLastIndex) = mFirstIndex

End Sub

'@================================================================================
' IEnumerable Interface Members
'@================================================================================

Private Property Get IEnumerable_Enumerator() As Enumerator
Set IEnumerable_Enumerator = getEnumerator
End Property

Private Function IEnumerable_NextItem(Item As Variant, cookie As Variant) As Boolean
If cookie = &H7FFFFFFF Then Exit Function
If IsEmpty(cookie) Then cookie = mFirstIndex

cookie = nextEntryIndex(cookie)
If cookie = mLastIndex Then
    cookie = &H7FFFFFFF
    Exit Function
End If

Set Item = Me.Item(CLng(cookie))
IEnumerable_NextItem = True

End Function

Private Sub IEnumerable_Remove(cookie As Variant)
Remove cookie
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
IEnumerable_UpdateSequence = mUpdateSequence
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Item(ByVal handle As Long) As ChartRegion
Const ProcName As String = "Item"

On Error GoTo Err

Set Item = mRegions(handle).Region

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Append(ByVal Region As ChartRegion) As Long
Const ProcName As String = "Append"

On Error GoTo Err

Append = allocateEntry
Set mRegions(Append).Region = Region

insertEntryBefore Append, mLastIndex

changed

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function IsFirst(ByVal handle As Long)
Const ProcName As String = "IsFirst"

On Error GoTo Err

IsFirst = (nextEntryIndex(mFirstIndex) = handle)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub MoveDown(ByVal handle As Long)
Dim newLocation As Long

Const ProcName As String = "MoveDown"

On Error GoTo Err

If nextEntryIndex(handle) = mLastIndex Then Exit Sub
newLocation = nextEntryIndex(nextEntryIndex(handle))

deleteEntry handle

insertEntryBefore handle, newLocation

changed

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub MoveUp(ByVal handle As Long)
Dim newLocation As Long

Const ProcName As String = "MoveUp"

On Error GoTo Err

If previousEntryIndex(previousEntryIndex(handle)) = mFirstIndex Then Exit Sub
newLocation = previousEntryIndex(handle)

deleteEntry handle

insertEntryBefore handle, newLocation

changed

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = getEnumerator
End Function

Friend Sub Remove(ByVal handle As Long)
Const ProcName As String = "Remove"

On Error GoTo Err

releaseEntry handle
deleteEntry handle
changed

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function allocateEntry() As Long
Const ProcName As String = "allocateEntry"

On Error GoTo Err

mRegionsIndex = mRegionsIndex + 1
If mRegionsIndex > UBound(mRegions) Then ReDim Preserve mRegions(2 * (UBound(mRegions) + 1) - 1) As RegionTableEntry
allocateEntry = mRegionsIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub changed()
mUpdateSequence = mUpdateSequence + 1
End Sub

Private Sub deleteEntry(ByVal index As Long)
Const ProcName As String = "deleteEntry"

On Error GoTo Err

nextEntryIndex(previousEntryIndex(index)) = nextEntryIndex(index)
previousEntryIndex(nextEntryIndex(index)) = previousEntryIndex(index)
nextEntryIndex(index) = 0
previousEntryIndex(index) = 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getEnumerator() As Enumerator
Const ProcName As String = "getEnumerator"

On Error GoTo Err

Set getEnumerator = CreateEnumerator(Me)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub insertEntryBefore( _
                ByVal newIndex As Long, _
                ByVal index As Long)
Const ProcName As String = "insertEntryBefore"

On Error GoTo Err

previousEntryIndex(newIndex) = previousEntryIndex(index)
nextEntryIndex(newIndex) = index
nextEntryIndex(previousEntryIndex(index)) = newIndex
previousEntryIndex(index) = newIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Property Let nextEntryIndex(ByVal index As Long, ByVal Value As Long)
Const ProcName As String = "nextEntryIndex"

On Error GoTo Err

mRegions(index).next = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get nextEntryIndex(ByVal index As Long) As Long
Const ProcName As String = "nextEntryIndex"

On Error GoTo Err

nextEntryIndex = mRegions(index).next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Let previousEntryIndex(ByVal index As Long, ByVal Value As Long)
Const ProcName As String = "previousEntryIndex"

On Error GoTo Err

mRegions(index).prev = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get previousEntryIndex(ByVal index As Long) As Long
Const ProcName As String = "previousEntryIndex"

On Error GoTo Err

previousEntryIndex = mRegions(index).prev

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub releaseEntry(ByVal handle As Long)
Const ProcName As String = "releaseEntry"

On Error GoTo Err

Set mRegions(handle).Region = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

