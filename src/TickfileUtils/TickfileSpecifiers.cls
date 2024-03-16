VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickFileSpecifiers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Event CollectionChanged( _
                ev As CollectionChangeEventData)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "TickFileSpecifiers"

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileSpecifiers                 As New EnumerableCollection

Private mChangeListeners                    As New Listeners

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Count() As Long
Const ProcName As String = "Count"
On Error GoTo Err

Count = mTickfileSpecifiers.Count

Exit Property

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Enumerator() As Enumerator
Set Enumerator = mTickfileSpecifiers.Enumerator
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add(ByVal tfs As TickfileSpecifier) As TickfileSpecifier
Const ProcName As String = "Add"
On Error GoTo Err

Set Add = tfs
mTickfileSpecifiers.Add tfs
fireCollectionChange CollItemAdded, tfs

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub AddCollectionChangeListener( _
                ByVal pListener As ICollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
On Error GoTo Err

mChangeListeners.Add pListener

Exit Sub

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

mTickfileSpecifiers.Clear

Exit Sub

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Clone() As TickFileSpecifiers
Const ProcName As String = "Clone"
On Error GoTo Err

Set Clone = New TickFileSpecifiers

Dim lSpec As TickfileSpecifier
For Each lSpec In mTickfileSpecifiers
    Clone.Add lSpec
Next

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function Item( _
                ByVal index As Variant) As TickfileSpecifier
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
On Error GoTo Err

Set Item = mTickfileSpecifiers(index)

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
On Error GoTo Err

Set NewEnum = mTickfileSpecifiers.Enumerator

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Remove( _
                ByVal index As Variant)
Const ProcName As String = "Remove"
On Error GoTo Err

Dim tfs As TickfileSpecifier

If IsObject(index) Then
    If index Is Nothing Then Exit Sub
    Set tfs = index
    mTickfileSpecifiers.Remove index
Else
    Set tfs = Item(index)
    mTickfileSpecifiers.Remove index
End If
fireCollectionChange CollItemRemoved, tfs

Exit Sub

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveCollectionChangeListener(ByVal pListener As ICollectionChangeListener)
Const ProcName As String = "RemoveCollectionChangeListener"
On Error GoTo Err

mChangeListeners.Remove pListener

Exit Sub

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireCollectionChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As TickfileSpecifier)
Const ProcName As String = "fireCollectionChange"
On Error GoTo Err

Dim ev As CollectionChangeEventData
Set ev.Source = Me
ev.changeType = changeType
Set ev.AffectedItem = affectedObject

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mChangeListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As ICollectionChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Change ev
    Next
End If

RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Sub


