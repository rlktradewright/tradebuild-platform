VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppendBarsTask"
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
Private Const ModuleName                    As String = "AppendBarsTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                        As TaskContext
Private mDestBars                           As Bars
Private mSourceBars                         As Bars
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
Dim i As Long

Do While mIndex <= mSourceBars.count And i < 100
    mDestBars.add mSourceBars.Bar(mIndex), _
                    mDestBars.currentBarNumber + 1, _
                    mSourceBars.Bar(mIndex).timestamp
    
    If mProgressModulus <> 0 Then
        If mIndex Mod mProgressModulus = 0 Then mTaskContext.progress 100 * mIndex / mSourceBars.count
    End If
    
    mIndex = mIndex + 1
    i = i + 1
Loop

If mIndex > mSourceBars.count Then
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
                ByVal destBars As Bars, _
                ByVal sourceBars As Bars)
Set mSourceBars = sourceBars
Set mDestBars = destBars
If mSourceBars.count / 20 >= 200 Then mProgressModulus = mSourceBars.count / 20

mIndex = 1

End Sub

'@================================================================================
' Helper Functions
'@================================================================================


