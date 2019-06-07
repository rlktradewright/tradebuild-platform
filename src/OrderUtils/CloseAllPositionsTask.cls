VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CloseAllPositionsTask"
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

Implements IChangeListener
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

Private Const ModuleName                            As String = "CloseAllPositionsTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private mNumberRemaining                            As Long

Private mClosePositionsResult                       As New ClosePositionsResult

Private mFutureBuilder                              As New FutureBuilder

Private mEnLive                                     As Enumerator
Private mEnSimulated                                As Enumerator

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
On Error GoTo Err

Dim pm As PositionManager
Set pm = ev.Source

Dim changeType As PositionManagerChangeTypes
changeType = ev.changeType

Select Case changeType
Case PositionSizeChanged
    Exit Sub
Case PositionManagerChangeTypes.PositionClosed
    If pm.PositionSize = 0 And pm.PendingPositionSize = 0 Then
        mNumberRemaining = mNumberRemaining - 1
        mClosePositionsResult.NumberOfPositionsClosed = mClosePositionsResult.NumberOfPositionsClosed + 1
        pm.RemoveChangeListener Me
    End If
Case PositionManagerChangeTypes.ProviderReadinessChanged
    If Not pm.ProviderIsReady Then
        mNumberRemaining = mNumberRemaining - 1
        mClosePositionsResult.NumberOfPositionsNotClosed = mClosePositionsResult.NumberOfPositionsNotClosed + 1
        pm.RemoveChangeListener Me
    ElseIf pm.ProviderIsReady Then
        mNumberRemaining = mNumberRemaining + 1
        pm.ClosePositions
        pm.AddChangeListener Me
    End If
End Select

If mTaskContext.State = TaskStateSuspended Then mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()
    
End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

If mEnLive Is Nothing Then
ElseIf mEnLive.MoveNext Then
    processPM mEnLive.Current
    Exit Sub
Else
    Set mEnLive = Nothing
End If

If mEnSimulated Is Nothing Then
ElseIf mEnSimulated.MoveNext Then
    processPM mEnSimulated.Current
    Exit Sub
Else
    Set mEnSimulated = Nothing
End If

If mNumberRemaining = 0 Then
    mFutureBuilder.Value = mClosePositionsResult
    mFutureBuilder.Complete
    mTaskContext.Finish mClosePositionsResult, False
Else
    mTaskContext.Suspend -1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Future() As IFuture
Set Future = mFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pOrderManager As OrderManager, _
                ByVal pTypes As PositionTypes)
Const ProcName As String = "Initialise"
On Error GoTo Err

If pTypes And PositionTypeLive Then Set mEnLive = pOrderManager.PositionManagersLive.Enumerator
If pTypes And PositionTypeSimulated Then Set mEnSimulated = pOrderManager.PositionManagersSimulated.Enumerator

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub processPM( _
                ByVal pm As PositionManager)
Const ProcName As String = "processPM"
On Error GoTo Err

If pm.PositionSize = 0 And pm.PendingPositionSize = 0 Then Exit Sub

If pm.ProviderIsReady Then
    mNumberRemaining = mNumberRemaining + 1
    pm.ClosePositions
    pm.AddChangeListener Me
Else
    mClosePositionsResult.NumberOfPositionsNotClosed = mClosePositionsResult.NumberOfPositionsNotClosed + 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub





