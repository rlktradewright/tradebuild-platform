VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InMemoryStore"
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

Private Const ModuleName                            As String = "InMemoryStore"

'@================================================================================
' Member variables
'@================================================================================

Private mOrderPersistenceDataGroups                 As New EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IEnumerable Interface Members
'@================================================================================

Private Property Get IEnumerable_Enumerator() As Enumerator
Set IEnumerable_Enumerator = Enumerator
End Property

Private Function IEnumerable_NextItem(Item As Variant, pCookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_NextItem"
On Error GoTo Err

Dim lCookie As EnumerationCookie
Dim lGroupsEn As Enumerator
Dim lItemsEn As Enumerator

Dim lGroup As OrderPersistenceDataGroup

If IsEmpty(pCookie) Then
    Set lGroupsEn = mOrderPersistenceDataGroups.Enumerator
    Set lCookie.Current = lGroupsEn
    
    If Not lGroupsEn.MoveNext Then
        IEnumerable_NextItem = False
        Exit Function
    End If
    
Else
    lCookie = pCookie
    Set lGroupsEn = lCookie.Current
    Set lItemsEn = lCookie.Next
    
    If lItemsEn.MoveNext Then
        Set Item = lItemsEn.Current
        IEnumerable_NextItem = True
        Exit Function
    End If
    
    If Not lGroupsEn.MoveNext Then
        IEnumerable_NextItem = False
        Exit Function
    End If
End If
    
Set lGroup = lGroupsEn.Current
Set lItemsEn = lGroup.Enumerator
Set lCookie.Next = lItemsEn
pCookie = lCookie

If lItemsEn.MoveNext Then
    Set Item = lItemsEn.Current
    IEnumerable_NextItem = True
Else
    IEnumerable_NextItem = False
End If
        
Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IEnumerable_Remove(pCookie As Variant)

End Sub

Private Function IEnumerable_Skip(ByVal numToSkip As Long, pCookie As Variant) As Boolean

End Function

Private Property Get IEnumerable_UpdateSequence() As Long

End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Enumerator() As Enumerator
Const ProcName As String = "Enumerator"
On Error GoTo Err

Set Enumerator = CreateEnumerator(Me)

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get StoreIsEmpty() As Boolean
Const ProcName As String = "StoreIsEmpty"
On Error GoTo Err

StoreIsEmpty = (mOrderPersistenceDataGroups.Count = 0)

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add(ByVal pOpItem As OPI)
Const ProcName As String = "Add"
On Error GoTo Err

Dim opg As OrderPersistenceDataGroup

If Not mOrderPersistenceDataGroups.Contains(pOpItem.RCN) Then
    Set opg = New OrderPersistenceDataGroup
    mOrderPersistenceDataGroups.Add opg, pOpItem.RCN
Else
    Set opg = mOrderPersistenceDataGroups.Item(pOpItem.RCN)
End If

opg.Add pOpItem

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Delete(ByVal pOpItem As OPI)
Const ProcName As String = "Delete"
On Error GoTo Err

If mOrderPersistenceDataGroups.Contains(pOpItem.RCN) Then mOrderPersistenceDataGroups.Remove pOpItem.RCN

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




