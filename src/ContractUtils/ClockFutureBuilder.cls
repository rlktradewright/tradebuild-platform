VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClockFutureBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

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

Private Const ModuleName                            As String = "ClockFutureBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mFutureBuilder                   As FutureBuilder
Attribute mFutureBuilder.VB_VarHelpID = -1

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mIsSimulated                                As Boolean
Private mClockRate                                  As Single

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mSelfRef = Me
Set mFutureBuilder = New FutureBuilder
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureBuilder Event Handlers
'@================================================================================

Private Sub mFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mFutureBuilder_Cancelled"
On Error GoTo Err

mFutureWaiter.Cancel

Exit Sub

Err:
ContractsGlobals.gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then
    If Future.IsPending Then mFutureBuilder.Cancel
ElseIf ev.Future.IsFaulted Then
    If Future.IsPending Then mFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
ElseIf ev.Future.IsAvailable Then
    Dim lContract As IContract: Set lContract = ev.Future.Value
    gLogger.Log "Contract available for clock: " & lContract.Specifier.ToString, ProcName, ModuleName, LogLevelDetail
    setupClock ev.Future.Value
End If

Exit Sub

Err:
ContractsGlobals.gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Future() As IFuture
Set Future = mFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractFuture As IFuture, _
                ByVal pIsSimulated As Boolean, _
                ByVal pClockRate As Single)
Const ProcName As String = "Initialise"
On Error GoTo Err

mIsSimulated = pIsSimulated
mClockRate = pClockRate

If pContractFuture.IsAvailable Then
    setupClock pContractFuture.Value
Else
    mFutureWaiter.Add pContractFuture
End If

Exit Sub

Err:
ContractsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub setupClock(ByVal pContract As IContract)
Const ProcName As String = "setupClock"
On Error GoTo Err

If mIsSimulated Then
    mFutureBuilder.Value = CreateSimulatedClock(mClockRate, pContract.TimezoneName)
Else
    mFutureBuilder.Value = GetClock(pContract.TimezoneName)
End If
mFutureBuilder.Complete

Exit Sub

Err:
ContractsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Sub



