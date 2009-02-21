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
                ev As CollectionChangeEvent)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "TradeBuild26"
Private Const ModuleName                    As String = "TickFileSpecifiers"

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileSpecifiers As Collection

Private mChangeListeners As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mTickfileSpecifiers = New Collection
Set mChangeListeners = New Collection
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

Public Property Get Count() As Long
Count = mTickfileSpecifiers.Count
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function add(ByVal tfs As TickfileSpecifier) As TickfileSpecifier
Set add = tfs
mTickfileSpecifiers.add tfs, CStr(ObjPtr(tfs))
fireChange CollItemAdded, tfs
End Function

Public Sub addCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.add value
End Sub

Public Sub Clear()
Set mTickfileSpecifiers = New Collection
End Sub

Public Function item( _
                ByVal index As Variant) As TickfileSpecifier
Attribute item.VB_UserMemId = 0
Set item = mTickfileSpecifiers(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mTickfileSpecifiers.[_NewEnum]
End Function

Public Sub remove( _
                ByVal index As Variant)
Dim tfs As TickfileSpecifier

If IsObject(index) Then
    If index Is Nothing Then Exit Sub
    Set tfs = index
    mTickfileSpecifiers.remove CStr(ObjPtr(tfs))
Else
    Set tfs = item(index)
    mTickfileSpecifiers.remove index
End If
fireChange CollItemRemoved, tfs
End Sub

Public Sub removeCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.item(i) Is value Then mChangeListeners.remove i
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As TickfileSpecifier)
Dim listener As CollectionChangeListener
Dim ev As CollectionChangeEvent
Set ev.source = Me
ev.changeType = changeType
Set ev.affectedItem = affectedObject
For Each listener In mChangeListeners
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)
End Sub


