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

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change(ev As TWUtilities30.ChangeEvent)
Dim pm As PositionManager
Dim changeType As PositionManagerChangeTypes

Set pm = ev.source
changeType = ev.changeType
Select Case changeType
Case PositionManagerChangeTypes.PositionClosed
    If pm.positionSize = 0 And pm.pendingPositionSize = 0 Then
        mNumberRemaining = mNumberRemaining - 1
        mClosePositionsResult.numberOfPositionsClosed = mClosePositionsResult.numberOfPositionsClosed + 1
        pm.removeChangeListener Me
        If mNumberRemaining = 0 Then mTaskContext.finish mClosePositionsResult, False
    End If
Case PositionManagerChangeTypes.ProviderReadinessChanged
    If Not pm.providerIsReady Then
        mNumberRemaining = mNumberRemaining - 1
        mClosePositionsResult.numberOfPositionsNotClosed = mClosePositionsResult.numberOfPositionsNotClosed + 1
        pm.removeChangeListener Me
        If mNumberRemaining = 0 Then mTaskContext.finish mClosePositionsResult, False
    End If
End Select
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()
    
End Sub

Private Sub Task_run()

initiateCloses
If mNumberRemaining = 0 Then
    mTaskContext.finish 0, False
Else
    mTaskContext.suspend -1
End If

End Sub

Private Property Let Task_TaskContext(ByVal RHS As TWUtilities30.TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.name
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
                ByVal types As PositionTypes, _
                ByVal flags As ClosePositionFlags)
mTypes = types
mFlags = flags
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


Private Sub check( _
                ByVal pm As PositionManager)
If pm.positionSize <> 0 Or _
    pm.pendingPositionSize <> 0 _
Then
    If pm.providerIsReady Then
        mNumberRemaining = mNumberRemaining + 1
        pm.closePositions mFlags
        pm.addChangeListener Me
    Else
        mClosePositionsResult.numberOfPositionsNotClosed = mClosePositionsResult.numberOfPositionsNotClosed + 1
    End If
End If
End Sub

Private Sub initiateCloses()
Dim ws As WorkSpace
Dim tckr As Ticker

For Each ws In gTB.WorkSpaces
    For Each tckr In ws.Tickers
        If mTypes And PositionTypeLive Then check tckr.PositionManager
        If mTypes And PositionTypeSimulated Then check tckr.PositionManagerSimulated
    Next
Next
End Sub

