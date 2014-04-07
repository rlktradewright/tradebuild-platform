VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClosePositionsTask"
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

Implements ChangeListener
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

Private Const ModuleName                    As String = "ClosePositionsTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                        As TaskContext
Private mNumberRemaining                    As Long
Private mTypes                              As PositionTypes
Private mFlags                              As ClosePositionFlags

Private mClosePositionsResult               As ClosePositionsResult

Private mWorkspaces                         As WorkSpaces

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change(ev As ChangeEventData)
Dim pm As PositionManager
Dim changeType As PositionManagerChangeTypes

Const ProcName As String = "ChangeListener_Change"
On Error GoTo Err

Set pm = ev.source
changeType = ev.changeType
Select Case changeType
Case PositionManagerChangeTypes.PositionClosed
    If pm.PositionSize = 0 And pm.PendingPositionSize = 0 Then
        mNumberRemaining = mNumberRemaining - 1
        mClosePositionsResult.NumberOfPositionsClosed = mClosePositionsResult.NumberOfPositionsClosed + 1
        pm.RemoveChangeListener Me
        If mNumberRemaining = 0 Then mTaskContext.Finish mClosePositionsResult, False
    End If
Case PositionManagerChangeTypes.ProviderReadinessChanged
    If Not pm.ProviderIsReady Then
        mNumberRemaining = mNumberRemaining - 1
        mClosePositionsResult.NumberOfPositionsNotClosed = mClosePositionsResult.NumberOfPositionsNotClosed + 1
        pm.RemoveChangeListener Me
        If mNumberRemaining = 0 Then mTaskContext.Finish mClosePositionsResult, False
    End If
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()
    
End Sub

Private Sub Task_run()
Const ProcName As String = "Task_run"
On Error GoTo Err

initiateCloses
If mNumberRemaining = 0 Then
    mTaskContext.Finish 0, False
Else
    mTaskContext.suspend -1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let Task_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.Name
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
                ByVal pWorkspaces As WorkSpaces, _
                ByVal pTypes As PositionTypes, _
                ByVal pFlags As ClosePositionFlags)
Set mWorkspaces = pWorkspaces
mTypes = pTypes
mFlags = pFlags
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


Private Sub check( _
                ByVal pm As PositionManager)
Const ProcName As String = "check"
On Error GoTo Err

If pm.PositionSize <> 0 Or _
    pm.PendingPositionSize <> 0 _
Then
    If pm.ProviderIsReady Then
        mNumberRemaining = mNumberRemaining + 1
        pm.ClosePositions mFlags
        pm.AddChangeListener Me
    Else
        mClosePositionsResult.NumberOfPositionsNotClosed = mClosePositionsResult.NumberOfPositionsNotClosed + 1
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initiateCloses()
Const ProcName As String = "initiateCloses"
On Error GoTo Err

Dim ws As WorkSpace
Dim tckr As Ticker

For Each ws In mWorkspaces
    For Each tckr In ws.Tickers
        If mTypes And PositionTypeLive Then check tckr.PositionManager
        If mTypes And PositionTypeSimulated Then check tckr.PositionManagerSimulated
    Next
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

