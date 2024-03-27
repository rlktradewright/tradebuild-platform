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

Private Const ModuleName                            As String = "ClosePositionsTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private mNumberRemaining                            As Long

Private mClosePositionsResult                       As New ClosePositionsResult

Private mFutureBuilder                              As New FutureBuilder

Private mEnLive                                     As Enumerator
Private mEnSimulated                                As Enumerator

Private mGroupName                                  As String

Private mCloseoutMode                               As CloseoutModes
Private mPriceSpec                                  As PriceSpecifier

Private mBracketOrderListener                       As BracketOrderListener

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

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
        If pm.IsSimulated Then
            mClosePositionsResult.NumberOfSimulatedPositionsClosed = mClosePositionsResult.NumberOfSimulatedPositionsClosed + 1
        Else
            mClosePositionsResult.NumberOfLivePositionsClosed = mClosePositionsResult.NumberOfLivePositionsClosed + 1
        End If
        pm.RemoveChangeListener Me
    End If
Case PositionManagerChangeTypes.PositionCloseFailed
    mNumberRemaining = mNumberRemaining - 1
    If pm.IsSimulated Then
        mClosePositionsResult.AddSimulatedErrorMessage pm.ContractName & ": " & pm.PositionCloseErrorMessage
        mClosePositionsResult.NumberOfSimulatedPositionsNotClosed = mClosePositionsResult.NumberOfSimulatedPositionsNotClosed + 1
    Else
        mClosePositionsResult.AddLiveErrorMessage pm.ContractName & ": " & pm.PositionCloseErrorMessage
        mClosePositionsResult.NumberOfLivePositionsNotClosed = mClosePositionsResult.NumberOfLivePositionsNotClosed + 1
    End If
    pm.RemoveChangeListener Me
Case PositionManagerChangeTypes.ProviderReadinessChanged
    If Not pm.ProviderIsReady Then
        mNumberRemaining = mNumberRemaining - 1
        If pm.IsSimulated Then
            mClosePositionsResult.NumberOfSimulatedPositionsNotClosed = mClosePositionsResult.NumberOfSimulatedPositionsNotClosed + 1
        Else
            mClosePositionsResult.NumberOfLivePositionsNotClosed = mClosePositionsResult.NumberOfLivePositionsNotClosed + 1
        End If
        pm.RemoveChangeListener Me
    ElseIf pm.ProviderIsReady Then
        mNumberRemaining = mNumberRemaining + 1
        pm.ClosePositions mCloseoutMode, mPriceSpec
        pm.AddChangeListener Me
    End If
End Select

mTaskContext.Continue

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.Value Is Nothing Then Exit Sub

If Not mBracketOrderListener Is Nothing Then mBracketOrderListener.Add ev.Future.Value

Exit Sub

Err:
GOrders.NotifyUnhandledError ProcName, ModuleName
End Sub

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
                ByVal pCloseoutMode As CloseoutModes, _
                Optional ByVal pPriceSpec As PriceSpecifier, _
                Optional ByVal pGroupName As String, _
                Optional ByVal pTypes As PositionTypes = PositionTypeLive, _
                Optional ByVal pBracketOrderListener As BracketOrderListener)
Const ProcName As String = "Initialise"
On Error GoTo Err

mGroupName = UCase$(pGroupName)
mCloseoutMode = pCloseoutMode
Set mPriceSpec = pPriceSpec
If (pTypes And PositionTypeLive) = PositionTypeLive Then Set mEnLive = pOrderManager.PositionManagersLive.Enumerator
If (pTypes And PositionTypeSimulated) = PositionTypeSimulated Then Set mEnSimulated = pOrderManager.PositionManagersSimulated.Enumerator
Set mBracketOrderListener = pBracketOrderListener

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub processPM( _
                ByVal pm As PositionManager)
Const ProcName As String = "processPM"
On Error GoTo Err

If mGroupName <> "" And (UCase$(pm.GroupName) <> mGroupName) Then Exit Sub

If pm.PositionSize <> 0 Then
ElseIf pm.PendingPositionSize <> 0 Then
ElseIf pm.HasBracketOrdersWithNonZeroSizes Then
ElseIf pm.HasBracketOrdersWithNonZeroPendingSizes Then
Else
    Exit Sub
End If

If pm.ProviderIsReady Then
    mNumberRemaining = mNumberRemaining + 1
    pm.AddChangeListener Me
    mFutureWaiter.Add pm.ClosePositions(mCloseoutMode, mPriceSpec)
Else
    If pm.IsSimulated Then
        mClosePositionsResult.NumberOfSimulatedPositionsNotClosed = mClosePositionsResult.NumberOfSimulatedPositionsNotClosed + 1
    Else
        mClosePositionsResult.NumberOfLivePositionsNotClosed = mClosePositionsResult.NumberOfLivePositionsNotClosed + 1
    End If
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub





