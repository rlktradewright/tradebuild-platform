VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Orders"
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

Private Const ModuleName                As String = "Orders"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mOrders As Collection

Private mChangeListeners As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mOrders = New Collection
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
Const ProcName As String = "Count"

On Error GoTo Err

Count = mOrders.Count

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Add(ByVal pOrder As Order) As Order
Const ProcName As String = "Add"

On Error GoTo Err

Set Add = pOrder
mOrders.Add pOrder, pOrder.Id
fireChange CollItemAdded, pOrder

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub AddCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"

On Error GoTo Err

mChangeListeners.Add value

gNotifyExistingCollectionMembers mOrders, value, Me
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Clear()

Const ProcName As String = "Clear"
On Error GoTo Err

Set mOrders = New Collection
fireChange CollCollectionCleared, Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function Contains(index As Variant) As Boolean
Const ProcName As String = "Contains"
Dim lOrder As Order
On Error Resume Next
Set lOrder = mOrders(index)
If Not lOrder Is Nothing Then Contains = True
End Function

Public Function Item(index As Variant) As Order
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"

On Error GoTo Err

Set Item = mOrders(index)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mOrders.[_NewEnum]
End Function

Public Sub Remove( _
                ByVal pOrder As Order)
Const ProcName As String = "Remove"

On Error GoTo Err

mOrders.Remove pOrder.Id
fireChange CollItemRemoved, pOrder

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
                ByVal affectedObject As Order)
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




