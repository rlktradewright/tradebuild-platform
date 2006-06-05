VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskManager"
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
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

Public Enum ErrorCodes
    CantCreateGUID = vbObjectError + 512
End Enum
'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mQuantumMillisecs As Long

Private mTasks As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTasks = New Collection
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

Public Property Let quantumMillisecs( _
                ByVal value As Long)
mQuantumMillisecs = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub finish( _
                ByVal tc As TaskContext)
mTasks.Remove tc.name
End Sub

Public Function startTask( _
                ByVal target As Runnable, _
                Optional ByVal name As String) As TaskCompletion
Dim tc As TaskContext

If name = "" Then name = gGenerateGUIDString

Set tc = New TaskContext
tc.taskManager = Me
tc.name = name
tc.target = target
tc.quantumMillisecs = mQuantumMillisecs
tc.run

mTasks.Add tc, name

Set startTask = tc
End Function

'================================================================================
' Helper Functions
'================================================================================

