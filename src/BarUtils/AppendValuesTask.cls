VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppendValuesTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements ITask

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "AppendValuesTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                        As TaskContext
Private mDestValues                         As ValueCache
Private mSourceValues                       As ValueCache
Private mIndex                              As Long
Private mEt                                 As New ElapsedTimer

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_cancel()

End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

mIndex = mIndex + 1

Dim val As SValue
val = mSourceValues.GetSValue(mIndex)
mDestValues.AddValue val.Value, val.BarNumber, val.Timestamp, val.BarStartTime

If mEt.ElapsedTimeMicroseconds > 50000 Then
    mTaskContext.Progress 100 * mIndex / mSourceValues.Count, Empty
    mEt.StartTiming
End If

mIndex = mIndex + 1

If mIndex > mSourceValues.Count Then
    mTaskContext.Progress 100, Empty
    mTaskContext.Finish Empty, False
    Exit Sub
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
    
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_taskName() As String
ITask_taskName = mTaskContext.Name
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal destValues As ValueCache, _
                ByVal sourceValues As ValueCache)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mSourceValues = sourceValues
Set mDestValues = destValues
mEt.StartTiming

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


