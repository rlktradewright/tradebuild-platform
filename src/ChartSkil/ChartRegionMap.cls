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

Implements Enumerable

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
' Enumerable Interface Members
'@================================================================================

Private Property Get Enumerable_Enumerator() As TWUtilities30.Enumerator
Set Enumerable_Enumerator = getEnumerator
End Property

Private Function Enumerable_NextItem(Item As Variant, cookie As Variant) As Boolean
If cookie = &H7FFFFFFF Then Exit Function
If IsEmpty(cookie) Then cookie = mFirstIndex

cookie = nextEntryIndex(cookie)
If cookie = mLastIndex Then
    cookie = &H7FFFFFFF
    Exit Function
End If

Set Item = Me.Item(CLng(cookie))
Enumerable_NextItem = True

End Function

Private Sub Enumerable_Remove(cookie As Variant)
Remove cookie
End Sub

Private Function Enumerable_Skip(ByVal numToSkip As Long, cookie As Variant) As Boolean
Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
        ProjectName & "." & ModuleName & ":" & "Enumerable_Skip", _
        "Skip not supported"
End Function

Private Property Get Enumerable_UpdateSequence() As Long
Enumerable_UpdateSequence = mUpdateSequence
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Item(ByVal handle As Long) As ChartRegion
Set Item = mRegions(handle).Region
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Append(ByVal Region As ChartRegion) As Long
Append = allocateEntry
Set mRegions(Append).Region = Region

insertEntryBefore Append, mLastIndex

changed
End Function

Friend Function IsFirst(ByVal handle As Long)
IsFirst = (nextEntryIndex(mFirstIndex) = handle)
End Function

Friend Sub MoveDown(ByVal handle As Long)
Dim newLocation As Long

If nextEntryIndex(handle) = mLastIndex Then Exit Sub
newLocation = nextEntryIndex(nextEntryIndex(handle))

deleteEntry handle

insertEntryBefore handle, newLocation

changed
End Sub

Friend Sub MoveUp(ByVal handle As Long)
Dim newLocation As Long

If previousEntryIndex(previousEntryIndex(handle)) = mFirstIndex Then Exit Sub
newLocation = previousEntryIndex(handle)

deleteEntry handle

insertEntryBefore handle, newLocation

changed
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = getEnumerator
End Function

Friend Sub Remove(ByVal handle As Long)
releaseEntry handle
deleteEntry handle
changed
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function allocateEntry() As Long
mRegionsIndex = mRegionsIndex + 1
If mRegionsIndex > UBound(mRegions) Then ReDim Preserve mRegions(2 * (UBound(mRegions) + 1) - 1) As RegionTableEntry
allocateEntry = mRegionsIndex
End Function

Private Sub changed()
mUpdateSequence = mUpdateSequence + 1
End Sub

Private Sub deleteEntry(ByVal index As Long)
nextEntryIndex(previousEntryIndex(index)) = nextEntryIndex(index)
previousEntryIndex(nextEntryIndex(index)) = previousEntryIndex(index)
nextEntryIndex(index) = 0
previousEntryIndex(index) = 0
End Sub

Private Function getEnumerator() As Enumerator
Set getEnumerator = CreateEnumerator(Me)
End Function

Private Sub insertEntryBefore( _
                ByVal newIndex As Long, _
                ByVal index As Long)
previousEntryIndex(newIndex) = previousEntryIndex(index)
nextEntryIndex(newIndex) = index
nextEntryIndex(previousEntryIndex(index)) = newIndex
previousEntryIndex(index) = newIndex

End Sub

Private Property Let nextEntryIndex(ByVal index As Long, ByVal value As Long)
mRegions(index).next = value
End Property

Private Property Get nextEntryIndex(ByVal index As Long) As Long
nextEntryIndex = mRegions(index).next
End Property

Private Property Let previousEntryIndex(ByVal index As Long, ByVal value As Long)
mRegions(index).prev = value
End Property

Private Property Get previousEntryIndex(ByVal index As Long) As Long
previousEntryIndex = mRegions(index).prev
End Property

Private Sub releaseEntry(ByVal handle As Long)
Set mRegions(handle).Region = Nothing
End Sub

