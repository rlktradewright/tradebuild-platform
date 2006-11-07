VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeakCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuildSP.IWeakCollection

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mCollection As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mCollection = New Collection
End Sub

'================================================================================
' IWeakCollection Interface Members
'================================================================================

Private Sub IWeakCollection_Add( _
                ByVal Item As Object, _
                Optional ByVal Key As Variant, _
                Optional ByVal before As Variant, _
                Optional ByVal after As Variant)
add Item, Key, before, after
End Sub

Private Function IWeakCollection_Count() As Long
IWeakCollection_Count = Count
End Function

Private Function IWeakCollection_Item(ByVal index As Variant) As Object
Set IWeakCollection_Item = Item(index)
End Function

Private Sub IWeakCollection_Remove(ByVal index As Variant)
remove index
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Public Sub add( _
                ByVal Item As Object, _
                Optional ByVal Key As Variant, _
                Optional ByVal before As Variant, _
                Optional ByVal after As Variant)
Dim weakRef As WeakReference

Set weakRef = gNewWeakReference(Item)

mCollection.add weakRef, Key, before, after
End Sub

Public Function Count() As Long
Count = mCollection.Count
End Function

Public Function Item( _
                ByVal index As Variant) As Object
Attribute Item.VB_UserMemId = 0
Dim weakRef As WeakReference

Set weakRef = mCollection.Item(index)
Set Item = weakRef.target
End Function

Public Sub remove( _
                ByVal index As Variant)
mCollection.remove index
End Sub

'================================================================================
' Helper Functions
'================================================================================



