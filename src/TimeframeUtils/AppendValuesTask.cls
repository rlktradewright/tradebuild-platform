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

Implements Task

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

Private Const ProjectName                   As String = "TimeframeUtils26"
Private Const ModuleName                    As String = "AppendValuesTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                        As TaskContext
Private mDestValues                         As ValueCache
Private mSourceValues                       As ValueCache
Private mIndex                              As Long
Private mProgressModulus                    As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_cancel()

End Sub

Private Sub Task_run()
Dim val As SValue
Dim i As Long

Const ProcName As String = "Task_run"
Dim failpoint As String
On Error GoTo Err

mIndex = mIndex + 1

Do While mIndex <= mSourceValues.Count And i < 100
    val = mSourceValues.GetSValue(mIndex)
    mDestValues.AddValue val.value, val.BarNumber, val.Timestamp, val.BarStartTime
    
    If mProgressModulus <> 0 Then
        If mIndex Mod mProgressModulus = 0 Then mTaskContext.Progress 100 * mIndex / mSourceValues.Count, Empty
    End If
    
    mIndex = mIndex + 1
    i = i + 1
Loop

If mIndex > mSourceValues.Count Then
    mTaskContext.Progress 100, Empty
    mTaskContext.Finish Empty, False
    Exit Sub
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
    
End Sub

Private Property Let Task_TaskContext(ByVal value As TWUtilities30.TaskContext)
Set mTaskContext = value
End Property

Private Property Get Task_taskName() As String
Task_taskName = mTaskContext.Name
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
Dim failpoint As String
On Error GoTo Err

Set mSourceValues = sourceValues
Set mDestValues = destValues
If mSourceValues.Count / 20 >= 200 Then mProgressModulus = mSourceValues.Count / 20

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


