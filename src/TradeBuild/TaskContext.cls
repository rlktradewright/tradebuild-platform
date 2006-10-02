VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskContext"
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
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuildSP.ITaskContext
Implements TradeBuildSP.ITaskCompletion
Implements TaskCompletion

'================================================================================
' Events
'================================================================================

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

Private mTaskManager As taskManager

Private mName As String

Private mTarget As Task
Private mData As Variant

Private mQuantumMillisecs As Long

Private WithEvents mTimer As TimerUtils.IntervalTimer
Attribute mTimer.VB_VarHelpID = -1
Private mElapsedTimer As TimerUtils.ElapsedTimer

Private mFinished As Boolean

Private mTaskCompletionListeners As WeakCollection

Private mYieldCount As Long

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' ITaskCompletion Interface Members
'================================================================================

Private Sub ITaskCompletion_addTaskCompletionListener( _
                ByVal listener As TradeBuildSP.ITaskCompletionListener)
addTaskCompletionListener listener
End Sub

Private Sub ITaskCompletion_removeTaskCompletionListener( _
                ByVal listener As TradeBuildSP.ITaskCompletionListener)
removeTaskCompletionListener listener
End Sub

'================================================================================
' ITaskContext Interface Members
'================================================================================

Private Sub ITaskContext_finish(Optional ByVal message As String)
finish message
End Sub

Private Property Get ITaskContext_isQuantumComplete() As Boolean
ITaskContext_isQuantumComplete = isQuantumComplete
End Property

Private Property Get ITaskContext_name() As String
ITaskContext_name = name
End Property

Private Sub ITaskContext_yield()
yield
End Sub

Private Property Get ITaskContext_yieldCount() As Long
ITaskContext_yieldCount = yieldCount
End Property

'================================================================================
' TaskCompletion Interface Members
'================================================================================

Private Sub TaskCompletion_addTaskCompletionListener( _
                ByVal listener As TaskCompletionListener)
addTaskCompletionListener listener
End Sub

Private Sub TaskCompletion_removeTaskCompletionListener( _
                ByVal listener As TaskCompletionListener)
removeTaskCompletionListener listener
End Sub

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
Static running As Boolean

If Not running Then
    Debug.Print "Running " & mName
    running = True
End If
mElapsedTimer.StartTiming

On Error GoTo err

Do
    mTarget.run
Loop Until mFinished Or _
            mElapsedTimer.ElapsedTimeMicroseconds / 1000 > mQuantumMillisecs

If Not mFinished Then yield

Exit Sub

err:
Debug.Print "Finished " & mName & ": error " & err.number & " " & err.Description
fireCompletion err.number, err.Description
Set mTaskCompletionListeners = Nothing
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let data(ByVal value As Variant)
mData = value
End Property

Public Property Get isQuantumComplete() As Boolean
If mElapsedTimer.ElapsedTimeMicroseconds / 1000 > mQuantumMillisecs Then isQuantumComplete = True
End Property

Friend Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

Friend Property Let quantumMillisecs(ByVal value As Long)
mQuantumMillisecs = value
End Property

Friend Property Let target(ByVal value As Task)
Set mTarget = value
End Property

Friend Property Let taskManager(ByVal value As taskManager)
Set mTaskManager = value
End Property

Public Property Get yieldCount() As Long
yieldCount = mYieldCount
End Property

'================================================================================
' Methods
'================================================================================

Public Sub finish(Optional ByVal message As String)

Debug.Print "Finish " & mName

mTaskManager.finish Me
Set mTaskManager = Nothing
mFinished = True

fireCompletion 0, message
Set mTarget = Nothing
Set mTaskCompletionListeners = Nothing
End Sub

Friend Sub run()
Debug.Print "Starting " & mName

mTarget.TaskContext = Me

Set mElapsedTimer = New ElapsedTimer

Set mTimer = New IntervalTimer
mTimer.TimerIntervalMillisecs = 1
mTimer.RepeatNotifications = False
mTimer.RandomIntervals = True
mTimer.StartTimer
End Sub

Public Sub yield()
mYieldCount = mYieldCount + 1
DoEvents
mTimer.TimerIntervalMillisecs = 1
mTimer.StartTimer
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub addTaskCompletionListener(ByVal listener As Object)
If mTaskCompletionListeners Is Nothing Then
    Set mTaskCompletionListeners = New WeakCollection
End If

mTaskCompletionListeners.add listener
End Sub

Private Sub fireCompletion( _
                ByVal resultCode As Long, _
                ByVal message As String)
Dim ev As TaskCompletionEvent
Dim spEv As TradeBuildSP.TaskCompletionEvent
Dim obj As Object
Dim listener As TaskCompletionListener
Dim i As Long
Dim spListener As TradeBuildSP.ITaskCompletionListener

If mTaskCompletionListeners Is Nothing Then Exit Sub

For i = 1 To mTaskCompletionListeners.Count
    Set obj = mTaskCompletionListeners(i)
    If TypeOf obj Is TaskCompletionListener Then
        Set listener = obj
        ev.resultCode = resultCode
        ev.message = message
        Set ev.source = mTarget
        ev.data = mData
        listener.taskCompleted ev
    Else
        Set spListener = obj
        spEv.resultCode = resultCode
        spEv.message = message
        Set spEv.source = mTarget
        spEv.data = mData
        spListener.taskCompleted spEv
    End If
Next
End Sub

Private Sub removeTaskCompletionListener(ByVal listener As TaskCompletionListener)
Dim i As Long

If mTaskCompletionListeners Is Nothing Then Exit Sub

For i = mTaskCompletionListeners.Count To 1 Step -1
    If mTaskCompletionListeners(i) Is listener Then mTaskCompletionListeners.remove i
Next
End Sub


