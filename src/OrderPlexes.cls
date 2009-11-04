VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPlexes"
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
                ev As CollectionChangeEvent)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "OrderPlexes"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mOrderPlexes As Collection
Private mChangeListeners As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mOrderPlexes = New Collection
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
Dim failpoint As Long
On Error GoTo Err

Count = mOrderPlexes.Count

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Add( _
                ByVal pOrderPlex As orderPlex) As orderPlex
Const ProcName As String = "Add"
Dim failpoint As Long
On Error GoTo Err

Set Add = pOrderPlex
mOrderPlexes.Add pOrderPlex, pOrderPlex.Key
fireChange CollItemAdded, pOrderPlex

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Sub AddCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
Dim failpoint As Long
On Error GoTo Err

mChangeListeners.Add value

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Function Contains(ByVal index As Variant) As Boolean
Dim op As orderPlex
On Error Resume Next
Set op = Item(index)
If Not op Is Nothing Then Contains = True
End Function

Public Function Item( _
                index As Variant) As orderPlex
Const ProcName As String = "Item"
Dim failpoint As Long
On Error GoTo Err

Set Item = mOrderPlexes(index)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mOrderPlexes.[_NewEnum]
End Function

Public Sub Remove( _
                ByVal pOrderPlex As orderPlex)
Const ProcName As String = "Remove"
Dim failpoint As Long
On Error GoTo Err

mOrderPlexes.Remove pOrderPlex.Key
fireChange CollItemRemoved, pOrderPlex

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemoveCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
Const ProcName As String = "RemoveCollectionChangeListener"
Dim failpoint As Long
On Error GoTo Err

For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As orderPlex)
Dim listener As CollectionChangeListener
Dim ev As CollectionChangeEvent
Const ProcName As String = "fireChange"
Dim failpoint As Long
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub





