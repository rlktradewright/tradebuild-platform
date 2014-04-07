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

Private Const ModuleName                    As String = "AppendBarsTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                        As TaskContext
Private mDestBars                           As Bars
Private mSourceBars                         As Bars
Private mIndex                              As Long
Private mNumberToAppend                     As Long
Private mBarsBuilder                        As BarsBuilder
Private mEt                                 As New ElapsedTimer

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_cancel()

End Sub

Private Sub Task_run()
Const ProcName As String = "Task_run"
On Error GoTo Err

mDestBars.Add mSourceBars.Bar(mIndex), _
                mDestBars.CurrentBarNumber + 1, _
                mSourceBars.Bar(mIndex).Timestamp
mBarsBuilder.FireBarAdded mSourceBars.Bar(mIndex)

If mEt.ElapsedTimeMicroseconds > 50000 Then
    mTaskContext.Progress 100 * mIndex / mNumberToAppend, Empty
    mEt.StartTiming
End If

mIndex = mIndex + 1

If mIndex > mNumberToAppend Then
    mTaskContext.Progress 100, Empty
    mTaskContext.Finish Empty, False
    Exit Sub
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let Task_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
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
                ByVal destBars As Bars, _
                ByVal sourceBars As Bars, _
                ByVal numberToAppend As Long, _
                ByVal pBarsBuilder As BarsBuilder)
Set mSourceBars = sourceBars
Set mDestBars = destBars
mNumberToAppend = numberToAppend
Set mBarsBuilder = pBarsBuilder

mIndex = 1
mEt.StartTiming

End Sub

'@================================================================================
' Helper Functions
'@================================================================================


