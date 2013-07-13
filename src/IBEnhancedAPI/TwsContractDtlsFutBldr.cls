VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsContractDetailsFuture"
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

Implements IFuture

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

Private Const ModuleName                            As String = "TwsContractDetailsFuture"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mFutureBuilder                   As FutureBuilder
Attribute mFutureBuilder.VB_VarHelpID = -1

Private WithEvents mTaskController                  As TaskController
Attribute mTaskController.VB_VarHelpID = -1

Private mContractRequester                          As ContractDetailsRequester

Private mTwsContract                                As TwsContractDetails

Private mIsCancelled                                As Boolean
Private mIsAvailable                                As Boolean
Private mIsFaulted                                  As Boolean
Private mErrorNumber                                As Long
Private mErrorMessage                               As String
Private mErrorSource                                As String

Private mListeners                                  As New EnumerableCollection

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureBuilder = New FutureBuilder
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' IFuture Interface Members
'@================================================================================

Private Sub IFuture_AddStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "IFuture_AddStateChangeListener"
On Error GoTo Err

mListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IFuture_Cancel()
Const ProcName As String = "IFuture_Cancel"
On Error GoTo Err

If Not IFuture_IsPending Then Exit Sub

mTaskController.CancelTask

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IFuture_ErrorMessage() As String
IFuture_ErrorMessage = mErrorMessage
End Property

Private Property Get IFuture_ErrorNumber() As Long
IFuture_ErrorNumber = mErrorNumber
End Property

Private Property Get IFuture_ErrorSource() As String
IFuture_ErrorSource = mErrorSource
End Property

Private Property Get IFuture_IsAvailable() As Boolean
IFuture_IsAvailable = mIsAvailable
End Property

Private Property Get IFuture_IsCancelled() As Boolean
IFuture_IsCancelled = mIsCancelled
End Property

Private Property Get IFuture_IsFaulted() As Boolean
IFuture_IsFaulted = mIsFaulted
End Property

Private Property Get IFuture_IsPending() As Boolean
IFuture_IsPending = Not (mIsAvailable Or mIsCancelled Or mIsFaulted)
End Property

Private Sub IFuture_RemoveStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "IFuture_RemoveStateChangeListener"
On Error GoTo Err

mListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IFuture_Value() As Object
Const ProcName As String = "IFuture_Value"
On Error GoTo Err

Assert Not mIsCancelled, "Future has been cancelled"
Assert Not mIsFaulted, "Future has failed: error=" & mErrorNumber & "; description=" & mErrorMessage & "; source=" & mErrorSource
Assert mIsAvailable, "Not yet completed"

Set IFuture_Value = mTwsContract

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then
    mIsCancelled = True
    fireStateChange FutureStateCancelled
ElseIf ev.Future.IsFaulted Then
    mIsFaulted = False
    mErrorNumber = ev.Future.ErrorNumber
    mErrorMessage = ev.Future.ErrorMessage
    mErrorSource = ev.Future.ErrorSource
    fireStateChange FutureStateErrored
ElseIf ev.Future.IsAvailable Then
    processContract ev.Future.Value
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mTaskController Event Handlers
'@================================================================================

Private Sub mTaskController_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mTaskController_Completed"
On Error GoTo Err

If ev.ErrorNumber <> 0 Then
    mIsFaulted = True
    mErrorNumber = ev.ErrorNumber
    mErrorMessage = ev.ErrorMessage
    mErrorSource = ev.ErrorSource
    fireStateChange FutureStateErrored
ElseIf ev.Cancelled Then
    mIsCancelled = True
    fireStateChange FutureStateCancelled
Else
    Set mTwsContract = ev.Result
    
    mIsAvailable = True
    fireStateChange FutureStateCompleted
End If

mListeners.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractFuture As IFuture)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractRequester = pContractRequester
mFutureWaiter.Add pContractFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireStateChange(ByVal pState As FutureStates)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim ev As StateChangeEventData

Set ev.Source = Me
ev.State = pState

Dim lListener As StateChangeListener
For Each lListener In mListeners
    lListener.Change ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processContract(ByVal pContract As IContract)
Const ProcName As String = "processContract"
On Error GoTo Err

If pContract.Specifier.Sectype = SecTypeCombo Then
    Dim lTask As New ComboContractConversionTask
    lTask.Initialise mContractRequester, pContract.Specifier
    Set mTaskController = StartTask(lTask, PriorityNormal)
Else
    Set mTwsContract = gContractToTwsContractDetails(pContract)
    mIsAvailable = True
    fireStateChange FutureStateCompleted
End If


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


