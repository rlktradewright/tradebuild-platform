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

'@================================================================================
' Description
'@================================================================================
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


Private Const ModuleName                As String = "Executions"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mExecutions As Collection
Private mChangeListeners As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mExecutions = New Collection
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
Count = mExecutions.Count
End Property

Public Property Get Item(ByVal Key As String) As Execution
Set Item = mExecutions.Item(Key)
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add(pExecution As Execution)
mExecutions.Add pExecution, pExecution.ExecId
fireChange CollItemAdded, pExecution
End Sub

Public Sub AddCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.Add value
End Sub

Friend Sub Change(pExecution As Execution)
Dim exec As Execution
Set exec = mExecutions.Item(pExecution.ExecId)
If Not exec Is pExecution Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "change", _
            "Changed Item is not a member of the collection"
End If
fireChange CollItemChanged, pExecution
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mExecutions.[_NewEnum]
End Function

Public Sub RemoveCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Execution)
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




