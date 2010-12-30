VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WorkSpaces"
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
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged( _
                ev As CollectionChangeEventData)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "WorkSpaces"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mWorkSpaces As Collection
Private mChangeListeners As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mWorkSpaces = New Collection
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
Count = mWorkSpaces.Count
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add(ByVal name As String) As WorkSpace
Const ProcName As String = "Add"

On Error GoTo Err

Set Add = New WorkSpace
Add.name = name
mWorkSpaces.Add Add, name
fireChange CollItemAdded, Add

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub AddCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"

On Error GoTo Err

mChangeListeners.Add value

gNotifyExistingCollectionMembers mWorkSpaces, value, Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function Item(index As Variant) As WorkSpace
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"

On Error GoTo Err

Set Item = mWorkSpaces(index)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mWorkSpaces.[_NewEnum]
End Function

Public Sub Remove(ByVal pWorkspace As WorkSpace)
Const ProcName As String = "Remove"

On Error GoTo Err

mWorkSpaces.Remove pWorkspace.name
fireChange CollItemRemoved, pWorkspace

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
Const ProcName As String = "RemoveCollectionChangeListener"

On Error GoTo Err

For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As WorkSpace)
Dim listener As CollectionChangeListener
Dim ev As CollectionChangeEventData
Const ProcName As String = "fireChange"

On Error GoTo Err

Set ev.source = Me
ev.changeType = changeType
Set ev.affectedItem = affectedObject
For Each listener In mChangeListeners
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub




