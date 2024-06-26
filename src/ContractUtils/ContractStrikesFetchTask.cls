VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractStrikesFetchTsk"
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

Private Const ModuleName                            As String = "ContractStrikesFetchTsk"

'@================================================================================
' Member variables
'@================================================================================

Private mUnderlyingContractSpecifier                As IContractSpecifier
Private mExchange                                   As String
Private mPrimaryContractStore                       As IContractStore
Private mExpiry                                     As String
Private mSecondaryContractStore                     As IContractStore
Private mCallerCookie                               As Variant

Private mTaskContext                                As TaskContext

Private mUsedSecondaryContractStore                 As Boolean

Private WithEvents mStrikesFutureBuilder            As FutureBuilder
Attribute mStrikesFutureBuilder.VB_VarHelpID = -1

Private mFetchFuture                                As IFuture

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStrikesFutureBuilder = New FutureBuilder
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

Dim lContractStore As IContractStore
Set lContractStore = mPrimaryContractStore
If lContractStore Is Nothing Then
    Set lContractStore = mSecondaryContractStore
    mUsedSecondaryContractStore = True
End If

Set mFetchFuture = lContractStore.FetchOptionStrikes( _
                                    mUnderlyingContractSpecifier, _
                                    mExchange, _
                                    mExpiry)
mFutureWaiter.Add mFetchFuture

mTaskContext.suspend -1

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Const ProcName As String = "ITask_TaskContext"
On Error GoTo Err

Set mTaskContext = Value

Exit Property

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get ITask_TaskName() As String
Const ProcName As String = "ITask_TaskName"
On Error GoTo Err

ITask_TaskName = mTaskContext.Name

Exit Property

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' mStrikesFutureBuilder Event Handlers
'@================================================================================

Private Sub mStrikesFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mStrikesFutureBuilder_Cancelled"
On Error GoTo Err

If Not mFetchFuture Is Nothing Then
    If Not mFetchFuture.IsCancelled Then mFetchFuture.Cancel
End If
mTaskContext.Finish Empty, True

Exit Sub

Err:
GContracts.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then
    mStrikesFutureBuilder.Cancel
ElseIf ev.Future.IsFaulted Then
    mStrikesFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
Else
    mStrikesFutureBuilder.Value = ev.Future.Value
    mStrikesFutureBuilder.Complete
    mTaskContext.Finish Empty, False
End If

Exit Sub

Err:
GContracts.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get StrikesFuture() As IFuture
Const ProcName As String = "StrikesFuture"
On Error GoTo Err

Set StrikesFuture = mStrikesFutureBuilder.Future

Exit Property

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pUnderlyingContractSpecifier, _
                ByVal pExchange As String, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pExpiry As String, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pCookie As Variant)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mUnderlyingContractSpecifier = pUnderlyingContractSpecifier
mExchange = pExchange
Set mPrimaryContractStore = pPrimaryContractStore
mExpiry = pExpiry
Set mSecondaryContractStore = pSecondaryContractStore
GContracts.SetVariant mCallerCookie, pCookie

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




