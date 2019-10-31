VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractFetchTask"
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

Implements IContractFetchListener
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

Private Const ModuleName                            As String = "ContractFetchTask"

'@================================================================================
' Member variables
'@================================================================================

Private mPrimaryContractStore                       As IContractStore
Private mSecondaryContractStore                     As IContractStore

Private mContractSpec                               As IContractSpecifier

Private mTaskContext                                As TaskContext

Private mUsedSecondaryContractStore                 As Boolean

Private WithEvents mContractsFutureBuilder          As FutureBuilder
Attribute mContractsFutureBuilder.VB_VarHelpID = -1
Private WithEvents mContractFutureBuilder           As FutureBuilder
Attribute mContractFutureBuilder.VB_VarHelpID = -1

Private mSingleContractOnly                         As Boolean

Private mListener                                   As IContractFetchListener

Private mCookie                                     As Variant

Private mFetchFuture                                As IFuture

Private mContractsBuilder                           As New ContractsBuilder

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mExpiryOffset                               As Long
Private mSortKeys()                                 As ContractSortKeyIds

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter

mExpiryOffset = -1

ReDim mSortKeys(0) As ContractSortKeyIds
mSortKeys(0) = ContractSortKeyExpiry
End Sub

'@================================================================================
' IContractFetchListener Interface Members
'@================================================================================

Private Sub IContractFetchListener_FetchCancelled(ByVal pCookie As Variant)
Const ProcName As String = "IContractFetchListener_FetchCancelled"
On Error GoTo Err

If Not futureIsPending Then Exit Sub

If mSingleContractOnly Then
    mContractFutureBuilder.Cancel
Else
    mContractsFutureBuilder.Cancel
End If

If Not mListener Is Nothing Then mListener.FetchCancelled pCookie
mTaskContext.Finish Empty, False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_FetchCompleted(ByVal pCookie As Variant)
Const ProcName As String = "IContractFetchListener_FetchCompleted"
On Error GoTo Err

If Not futureIsPending Then Exit Sub

If gIsOffsetExpiry(mContractSpec.Expiry) And _
    mContractsBuilder.Contracts.Count <> 0 _
Then
    mExpiryOffset = CLng(mContractSpec.Expiry)
    
    ' we've now got at least one contract that meets the specified
    ' contract spec. Now we need to find all the contracts that do
    ' so (apart from the expiry date)
    Dim lContractSpec As IContractSpecifier
    Set lContractSpec = mContractsBuilder.Contracts.ItemAtIndex(1).Specifier
    With lContractSpec
        Set mContractSpec = gCreateContractSpecifier("", _
                                                    .Symbol, _
                                                    .Exchange, _
                                                    .SecType, _
                                                    .CurrencyCode, _
                                                    "", _
                                                    .Multiplier, _
                                                    .Strike, _
                                                    .Right)
    End With
    mTaskContext.Continue
    Exit Sub
End If

If mExpiryOffset <> -1 Then
    mContractsBuilder.Contracts.SortKeys = mSortKeys
    Dim lAllContracts As IContracts: Set lAllContracts = mContractsBuilder.Contracts
    
    Set mContractsBuilder = New ContractsBuilder
    Dim lContract As IContract
    Dim lIndex As Long: lIndex = -1
    For Each lContract In lAllContracts
        If gIsContractExpired(lContract) Then
            lIndex = -1
        Else
            lIndex = lIndex + 1
        End If
        If lIndex = mExpiryOffset Then
            ' we've found the contract we're after
            mContractsBuilder.Add lContract
            If Not mListener Is Nothing Then mListener.NotifyContract pCookie, lContract
            Exit For
        End If
    Next
End If


Dim lContracts As IContracts
Set lContracts = mContractsBuilder.Contracts
If lContracts.Count = 0 And Not mUsedSecondaryContractStore Then If trySecondaryContractSP Then Exit Sub

If mSingleContractOnly Then
    mFutureWaiter.Add mContractFutureBuilder.Future, pCookie
    If lContracts.Count = 0 Then
        mContractFutureBuilder.Fail ErrorCodes.ErrIllegalArgumentException, "No such contract", ""
    Else
        mContractFutureBuilder.Value = lContracts.ItemAtIndex(1)
        mContractFutureBuilder.Complete
    End If
Else
    mFutureWaiter.Add mContractsFutureBuilder.Future, pCookie
    mContractsFutureBuilder.Value = lContracts
    mContractsFutureBuilder.Complete
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_FetchFailed(ByVal pCookie As Variant, ByVal pErrorCode As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
Const ProcName As String = "IContractFetchListener_FetchFailed"
On Error GoTo Err

If Not futureIsPending Then Exit Sub

If mSingleContractOnly Then
    mContractFutureBuilder.Fail pErrorCode, pErrorMessage, pErrorSource
Else
    mContractsFutureBuilder.Fail pErrorCode, pErrorMessage, pErrorSource
End If

If Not mListener Is Nothing Then mListener.FetchFailed pCookie, pErrorCode, pErrorMessage, pErrorSource
mTaskContext.Finish Empty, False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_NotifyContract(ByVal pCookie As Variant, ByVal pContract As IContract)
Const ProcName As String = "IContractFetchListener_NotifyContract"
On Error GoTo Err

If Not futureIsPending Then Exit Sub

If gIsOffsetExpiry(mContractSpec.Expiry) Or mExpiryOffset <> -1 Then
    mContractsBuilder.Add pContract
Else
    If mContractsBuilder.Contracts.Count = 1 And mSingleContractOnly Then
        mContractFutureBuilder.Fail ErrorCodes.ErrIllegalArgumentException, "Contract is not uniquely specified", ""
        mFetchFuture.Cancel
        mTaskContext.Finish Empty, False
        Exit Sub
    End If
    
    mContractsBuilder.Add pContract
    
    If Not mListener Is Nothing Then mListener.NotifyContract pCookie, pContract
End If

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
Const ProcName As String = "ITask_Run"
On Error GoTo Err

Dim lContractStore As IContractStore
Set lContractStore = mPrimaryContractStore
If lContractStore Is Nothing Then
    Set lContractStore = mSecondaryContractStore
    mUsedSecondaryContractStore = True
End If

Set mFetchFuture = lContractStore.FetchContracts(mContractSpec, Me, mCookie)

mTaskContext.suspend -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Const ProcName As String = "ITask_TaskContext"
On Error GoTo Err

Set mTaskContext = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get ITask_TaskName() As String
Const ProcName As String = "ITask_TaskName"
On Error GoTo Err

ITask_TaskName = mTaskContext.Name

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' mContractFutureBuilder Handlers
'@================================================================================

Private Sub mContractFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mContractFutureBuilder_Cancelled"
On Error GoTo Err

If Not mFetchFuture Is Nothing Then mFetchFuture.Cancel
If Not mListener Is Nothing Then mListener.FetchCancelled mContractFutureBuilder.Cookie
mTaskContext.Finish Empty, True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mContractsFutureBuilder Handlers
'@================================================================================

Private Sub mContractsFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mContractsFutureBuilder_Cancelled"
On Error GoTo Err

If Not mFetchFuture Is Nothing Then mFetchFuture.Cancel
If Not mListener Is Nothing Then mListener.FetchCancelled mContractsFutureBuilder.Cookie
mTaskContext.Finish Empty, True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not mListener Is Nothing Then mListener.FetchCompleted ev.ContinuationData
mTaskContext.Finish Empty, False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub
'@================================================================================
' Properties
'@================================================================================

Friend Property Get ContractFuture() As IFuture
Const ProcName As String = "ContractFuture"
On Error GoTo Err

Set ContractFuture = mContractFutureBuilder.Future

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ContractsFuture() As IFuture
Const ProcName As String = "ContractsFuture"
On Error GoTo Err

Set ContractsFuture = mContractsFutureBuilder.Future

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore, _
                ByVal pCookie As Variant, _
                ByVal pListener As IContractFetchListener, _
                ByVal pSingleContractOnly As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractSpec = pContractSpec
Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore

gSetVariant mCookie, pCookie

Set mListener = pListener

mSingleContractOnly = pSingleContractOnly
If mSingleContractOnly Then
    Set mContractFutureBuilder = New FutureBuilder
    mContractFutureBuilder.Cookie = mCookie
Else
    Set mContractsFutureBuilder = New FutureBuilder
    mContractsFutureBuilder.Cookie = mCookie
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function futureIsPending() As Boolean
Const ProcName As String = "futureIsPending"
On Error GoTo Err

If mSingleContractOnly Then
    futureIsPending = mContractFutureBuilder.Future.IsPending
Else
    futureIsPending = mContractsFutureBuilder.Future.IsPending
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function trySecondaryContractSP() As Boolean
Const ProcName As String = "trySecondaryContractSP"
On Error GoTo Err

mUsedSecondaryContractStore = True
If mSecondaryContractStore Is Nothing Then
    trySecondaryContractSP = False
Else
    Set mFetchFuture = mSecondaryContractStore.FetchContracts(mContractSpec, Me, mCookie)
    trySecondaryContractSP = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




