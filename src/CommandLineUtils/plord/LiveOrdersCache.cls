VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LiveOrdersCache"
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

Private Const ModuleName                            As String = "LiveOrdersCache"

'@================================================================================
' Member variables
'@================================================================================

Private mLiveOrders                                 As New EnumerableCollection

Private mLiveOrdersList()                           As LiveOrderEntry
Private mLiveOrdersListNextIndex                    As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mLiveOrdersList(15) As LiveOrderEntry
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Count() As Long
Count = mLiveOrders.Count
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add(ByVal pOrderEntry As LiveOrderEntry)
Const ProcName As String = "Add"
On Error GoTo Err

pOrderEntry.Timestamp = GetTimestamp
mLiveOrders.Add pOrderEntry, pOrderEntry.Key
If mLiveOrdersListNextIndex > UBound(mLiveOrdersList) Then
    ReDim Preserve mLiveOrdersList(2 * (UBound(mLiveOrdersList) + 1) - 1) As LiveOrderEntry
End If
Set mLiveOrdersList(mLiveOrdersListNextIndex) = pOrderEntry
mLiveOrdersListNextIndex = mLiveOrdersListNextIndex + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function Item(ByVal pKey As String) As LiveOrderEntry
Const ProcName As String = "Item"
On Error GoTo Err

Set Item = mLiveOrders.Item(pKey)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function ItemAtIndex(ByVal pIndex As Long) As LiveOrderEntry
Const ProcName As String = "ItemAtIndex"
On Error GoTo Err

AssertArgument mLiveOrdersListNextIndex > 0, "Invalid index: cache is empty"
AssertArgument checkIndex(pIndex), "Invalid index: must be >= -" & (mLiveOrdersListNextIndex - 1) & _
                                                        " and <= " & mLiveOrdersListNextIndex

If pIndex > 0 Then
    Set ItemAtIndex = mLiveOrdersList(pIndex - 1)
Else
    Set ItemAtIndex = mLiveOrdersList(pIndex + mLiveOrdersListNextIndex - 1)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
On Error GoTo Err

Set NewEnum = mLiveOrders.Enumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Remove(ByVal pKey As String)
Const ProcName As String = "Remove"
On Error GoTo Err

mLiveOrders.Remove pKey

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function TryItem(ByVal pKey As String, _
                ByRef pItem As LiveOrderEntry) As Boolean
Const ProcName As String = "TryItem"
On Error GoTo Err

TryItem = mLiveOrders.TryItem(pKey, pItem)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function TryItemAtIndex( _
                ByVal pIndex As Long, _
                ByRef pItem As LiveOrderEntry) As Boolean
Const ProcName As String = "TryItemAtIndex"
On Error GoTo Err

If mLiveOrdersListNextIndex = 0 Then Exit Function
If Not checkIndex(pIndex) Then Exit Function

Set pItem = ItemAtIndex(pIndex)
TryItemAtIndex = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function checkIndex(ByVal pIndex As Long) As Boolean
If pIndex > 0 And pIndex <= mLiveOrdersListNextIndex Then
    checkIndex = True
ElseIf pIndex <= 0 And (pIndex + mLiveOrdersListNextIndex - 1) >= 0 Then
    checkIndex = True
End If
End Function


