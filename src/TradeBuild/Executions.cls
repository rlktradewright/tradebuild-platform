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
Attribute VB_Creatable = False
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

'================================================================================
' Events
'================================================================================

Event ExecutionAdded(ByVal pExecution As Execution)

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

Private mExecutions As Collection
Private mChangeListeners As WeakCollection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mExecutions = New Collection
Set mChangeListeners = New WeakCollection
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get Count() As Long
Count = mExecutions.Count
End Property

Public Property Get Item(ByVal key As String) As Execution
Set Item = mExecutions.Item(key)
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub add(pExecution As Execution)
mExecutions.add pExecution, pExecution.ExecId
fireChange CollItemAdded, pExecution
RaiseEvent ExecutionAdded(pExecution)
End Sub

Public Sub addCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.add value
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mExecutions.[_NewEnum]
End Function

Public Sub removeCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.remove i
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Execution)
Dim listener As CollectionChangeListener
Dim i As Long
Dim Change As CollectionChangeEvent
Set Change.source = Me
Change.changeType = changeType
Set Change.affectedObject = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change Change
Next
End Sub




