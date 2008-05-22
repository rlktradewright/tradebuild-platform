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

mIndex = mIndex + 1

Do While mIndex <= mSourceValues.count And i < 100
    val = mSourceValues.getSValue(mIndex)
    mDestValues.addValue val.value, val.barNumber, val.timestamp, val.BarStartTime
    
    If mProgressModulus <> 0 Then
        If mIndex Mod mProgressModulus = 0 Then mTaskContext.progress 100 * mIndex / mSourceValues.count
    End If
    
    mIndex = mIndex + 1
    i = i + 1
Loop

If mIndex > mSourceValues.count Then
    mTaskContext.progress 100
    mTaskContext.finish Empty, False
    Exit Sub
End If
    
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

Friend Sub initialise( _
                ByVal destValues As ValueCache, _
                ByVal sourceValues As ValueCache)
Set mSourceValues = sourceValues
Set mDestValues = destValues
If mSourceValues.count / 20 >= 200 Then mProgressModulus = mSourceValues.count / 20
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


