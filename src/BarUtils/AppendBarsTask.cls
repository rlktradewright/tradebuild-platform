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
' ITask Interface Members
'@================================================================================

Private Sub ITask_cancel()

End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

' do this here to allow for the case mNumberToAppend=0
If mIndex > mNumberToAppend Then
    mTaskContext.Progress 100, Empty
    mTaskContext.Finish Empty, False
    Exit Sub
End If

mDestBars.Add mSourceBars.Bar(mIndex), _
                mDestBars.CurrentBarNumber + 1, _
                mSourceBars.Bar(mIndex).Timestamp
mBarsBuilder.FireBarAdded mSourceBars.Bar(mIndex)

If mEt.ElapsedTimeMicroseconds > 250000 Then
    mTaskContext.Progress 100 * mIndex / mNumberToAppend, Empty
    mEt.StartTiming
End If

mIndex = mIndex + 1

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
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


