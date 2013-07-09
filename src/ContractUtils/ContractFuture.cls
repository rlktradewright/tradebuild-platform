VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractFuture"
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

Private Const ModuleName                            As String = "ContractFuture"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTaskController                  As TaskController
Attribute mTaskController.VB_VarHelpID = -1
Private mReturnIContractsObject                     As Boolean

Private mContracts                                  As IContracts
Private mContract                                   As IContract
Private mIsCancelled                                As Boolean
Private mIsAvailable                                As Boolean
Private mIsFaulted                                  As Boolean
Private mErrorNumber                                As Long
Private mErrorMessage                               As String
Private mErrorSource                                As String

Private mListeners                                  As New EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

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

Assert Not mIsCancelled, "Contract fetch task was cancelled"
Assert Not mIsFaulted, "Contract fetch task failed: error=" & mErrorNumber & "; description=" & mErrorMessage & "; source=" & mErrorSource
Assert mIsAvailable, "Contract fetch task has not yet completed"

If mReturnIContractsObject Then
    Set IFuture_Value = mContracts
Else
    Set IFuture_Value = mContract
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

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
    Set mContracts = ev.Result
    
    If mReturnIContractsObject Then
        mIsAvailable = True
        fireStateChange FutureStateCompleted
    ElseIf mContracts.Count = 0 Then
        mIsFaulted = True
        mErrorNumber = ErrorCodes.ErrIllegalArgumentException
        mErrorMessage = "No such contract"
        fireStateChange FutureStateErrored
    ElseIf mContracts.Count > 1 Then
        mIsFaulted = True
        mErrorNumber = ErrorCodes.ErrIllegalArgumentException
        mErrorMessage = "Contract not uniquely specified"
        fireStateChange FutureStateErrored
    Else
        mIsAvailable = True
        Set mContract = mContracts.ItemAtIndex(1)
        fireStateChange FutureStateCompleted
    End If
End If

mListeners.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pTaskController As TaskController, _
                ByVal pReturnIContractsObject As Boolean)
Set mTaskController = pTaskController
mReturnIContractsObject = pReturnIContractsObject
End Sub

Friend Sub InitialiseFromContract(ByVal pContract As IContract)
Set mContract = pContract
mIsAvailable = True
End Sub

Friend Sub InitialiseFromContracts(ByVal pContracts As IContracts)
Set mContracts = pContracts
mReturnIContractsObject = True
mIsAvailable = True
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


