VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsContractDtlsFutBldr"
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

Private Const ModuleName                            As String = "TwsContractDtlsFutBldr"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mFutureBuilder                   As FutureBuilder
Attribute mFutureBuilder.VB_VarHelpID = -1

Private WithEvents mTaskController                  As TaskController
Attribute mTaskController.VB_VarHelpID = -1

Private mContractRequester                          As ContractDetailsRequester

Private mContract                                   As IContract
Private mTwsContract                                As TwsContractDetails

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mSelfRef                                    As Object

Private mContractCache                              As TwsContractCache

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureBuilder = New FutureBuilder
Set mFutureWaiter = New FutureWaiter
Set mSelfRef = Me
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

If Not mTaskController Is Nothing Then
    mTaskController.CancelTask
    Set mTaskController = Nothing
End If

mFutureBuilder.Cancel
Set mSelfRef = Nothing
    
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsPending Then Exit Sub

If ev.Future.IsCancelled Then
    mFutureBuilder.Cancel
ElseIf ev.Future.IsFaulted Then
    mFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
ElseIf TypeOf ev.Future.Value Is IContract Then
    Set mContract = ev.Future.Value
    processContract mContract
ElseIf TypeOf ev.Future.Value Is EnumerableCollection Then
    Dim lTwsContracts As EnumerableCollection
    Set lTwsContracts = ev.Future.Value
    If lTwsContracts.Count <> 1 Then
        mFutureBuilder.Fail ErrorCodes.ErrIllegalArgumentException, "Contract's specifier is not uniquely specified", ProjectName & "." & ModuleName & ":" & ProcName
    Else
        gLogger.Log "TwsContractDetails retrieved from IB for IContract: " & mContract.Specifier.ToString, ProcName, ModuleName, LogLevelDetail
        mFutureBuilder.Value = lTwsContracts.Item(1)
        mFutureBuilder.Complete
        Set mSelfRef = Nothing
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTaskController Event Handlers
'@================================================================================

Private Sub mTaskController_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mTaskController_Completed"
On Error GoTo Err

If ev.ErrorNumber <> 0 Then
    mFutureBuilder.Fail ev.ErrorNumber, ev.ErrorMessage, ev.ErrorSource
ElseIf ev.Cancelled Then
    mFutureBuilder.Cancel
Else
    gLogger.Log "IContract converted to TwsContractDetails: " & mContract.Specifier.ToString, ProcName, ModuleName, LogLevelDetail
    Set mTwsContract = ev.Result
    mFutureBuilder.Value = mTwsContract
    mFutureBuilder.Complete
    Set mSelfRef = Nothing
End If

Set mSelfRef = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Future() As IFuture
Const ProcName As String = "Future"
On Error GoTo Err

Set Future = mFutureBuilder.Future

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractFuture As IFuture, _
                ByVal pContractCache As TwsContractCache)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractRequester = pContractRequester
mFutureWaiter.Add pContractFuture
Set mContractCache = pContractCache

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub processContract(ByVal pContract As IContract)
Const ProcName As String = "processContract"
On Error GoTo Err

If pContract.Specifier.Sectype = SecTypeCombo Then
    Dim lTask As New ComboContractConversionTask
    lTask.Initialise mContractRequester, pContract.Specifier
    Set mTaskController = StartTask(lTask, PriorityNormal)
Else
    mFutureWaiter.Add gFetchContracts(mContractRequester, mContractCache, pContract.Specifier, Nothing, Empty, True)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


