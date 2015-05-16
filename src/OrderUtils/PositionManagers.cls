VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PositionManagers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements Enumerable

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged( _
                ev As CollectionChangeEventData)

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                        As String = "PositionManagers"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mPositionManagers            As EnumerableCollection
Attribute mPositionManagers.VB_VarHelpID = -1

Private mCollectionChangeListeners              As New Listeners

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mPositionManagers = New EnumerableCollection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get Enumerable_Enumerator() As Enumerator
Const ProcName As String = "Enumerable_Enumerator"
On Error GoTo Err

Set Enumerable_Enumerator = getEnumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function Enumerable_NextItem(Item As Variant, pCookie As Variant) As Boolean

End Function

Private Sub Enumerable_Remove(pCookie As Variant)

End Sub

Private Function Enumerable_Skip(ByVal numToSkip As Long, pCookie As Variant) As Boolean

End Function

Private Property Get Enumerable_UpdateSequence() As Long

End Property

'@================================================================================
' mPositionManagers Event Handlers
'@================================================================================

Private Sub mPositionManagers_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mPositionManagers_CollectionChanged"
On Error GoTo Err

Set ev.Source = Me
fireCollectionChanged ev

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Count() As Long
Const ProcName As String = "Count"
On Error GoTo Err

Count = mPositionManagers.Count

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Enumerator() As Enumerator
Set Enumerator = mPositionManagers.Enumerator
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Add(ByVal pPositionManager As PositionManager) As PositionManager
Const ProcName As String = "Add"
On Error GoTo Err

mPositionManagers.Add pPositionManager, pPositionManager.Name
Set Add = pPositionManager

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub AddCollectionChangeListener( _
                ByVal pListener As CollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
On Error GoTo Err

mCollectionChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

mPositionManagers.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Contains(ByVal pName As String) As Boolean
Const ProcName As String = "Item"
On Error GoTo Err

Contains = mPositionManagers.Contains(pName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Item(ByVal pName As String) As PositionManager
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
On Error GoTo Err

Set Item = mPositionManagers.Item(pName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
On Error GoTo Err

Set NewEnum = getEnumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub NotifyChange(pPositionManager As PositionManager)
Const ProcName As String = "NotifyChange"
On Error GoTo Err

mPositionManagers.NotifyChange pPositionManager.Name

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Remove(ByVal pPositionManager As PositionManager)
Const ProcName As String = "Remove"
On Error GoTo Err

mPositionManagers.Remove pPositionManager.Name

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveCollectionChangeListener(ByVal pListener As CollectionChangeListener)
Const ProcName As String = "RemoveCollectionChangeListener"
On Error GoTo Err

mCollectionChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireCollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "fireCollectionChanged"
On Error GoTo Err

mCollectionChangeListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mCollectionChangeListeners.Count
    Dim lListener As CollectionChangeListener
    Set lListener = mCollectionChangeListeners.GetListener(i)
    lListener.Change ev
Next

RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getEnumerator() As Enumerator
Const ProcName As String = "getEnumerator"
On Error GoTo Err

Set getEnumerator = mPositionManagers.Enumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function







