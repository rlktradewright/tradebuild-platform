VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Executions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements IEnumerable

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged( _
                ev As CollectionChangeEventData)

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                        As String = "Executions"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mExecutions                  As EnumerableCollection
Attribute mExecutions.VB_VarHelpID = -1

Private mCollectionChangeListeners              As New Listeners

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mExecutions = New EnumerableCollection

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IEnumerable Interface Members
'@================================================================================

Private Property Get IEnumerable_Enumerator() As Enumerator
Const ProcName As String = "IEnumerable_Enumerator"
On Error GoTo Err

Set IEnumerable_Enumerator = getEnumerator

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Private Function IEnumerable_NextItem(Item As Variant, pCookie As Variant) As Boolean

End Function

Private Sub IEnumerable_Remove(pCookie As Variant)

End Sub

Private Function IEnumerable_Skip(ByVal numToSkip As Long, pCookie As Variant) As Boolean

End Function

Private Property Get IEnumerable_UpdateSequence() As Long

End Property

'@================================================================================
' mExecutions Event Handlers
'@================================================================================

Private Sub mExecutions_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mExecutions_CollectionChanged"
On Error GoTo Err

Set ev.Source = Me
If ev.changeType = CollCollectionCleared Then Set ev.AffectedItem = Nothing
fireCollectionChanged ev

Exit Sub

Err:
GOrders.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Count() As Long
Const ProcName As String = "Count"
On Error GoTo Err

Count = mExecutions.Count

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add(ByVal pExecution As IExecutionReport) As IExecutionReport
Const ProcName As String = "Add"
On Error GoTo Err

mExecutions.Add pExecution, pExecution.Id

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub AddCollectionChangeListener( _
                ByVal pListener As ICollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
On Error GoTo Err

mCollectionChangeListeners.Add pListener

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

mExecutions.Clear

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Contains(ByVal pIndex As Variant) As Boolean
Const ProcName As String = "Contains"
On Error GoTo Err

Contains = mExecutions.Contains(pIndex)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function Item(ByVal pExecId As String) As IExecutionReport
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
On Error GoTo Err

Set Item = mExecutions.Item(pExecId)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
On Error GoTo Err

Set NewEnum = getEnumerator

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub NotifyChange(pExecution As IExecutionReport)
Const ProcName As String = "NotifyChange"
On Error GoTo Err

mExecutions.NotifyChange pExecution.Id

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveCollectionChangeListener(ByVal pListener As ICollectionChangeListener)
Const ProcName As String = "RemoveCollectionChangeListener"
On Error GoTo Err

mCollectionChangeListeners.Remove pListener

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireCollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "fireCollectionChanged"
On Error GoTo Err

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mCollectionChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mCollectionChangeListeners.GetCurrentListeners(sCurrentListeners)
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getEnumerator() As Enumerator
Const ProcName As String = "getEnumerator"
On Error GoTo Err

Set getEnumerator = mExecutions.Enumerator

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function





