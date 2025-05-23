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
Private mCurrentContractStore                       As IContractStore

Private mInitialContractSpec                        As IContractSpecifier
Private mContractSpec                               As IContractSpecifier

Private mTaskContext                                As TaskContext

Private mUsedSecondaryContractStore                 As Boolean

Private WithEvents mContractsFutureBuilder          As FutureBuilder
Attribute mContractsFutureBuilder.VB_VarHelpID = -1
Private WithEvents mContractFutureBuilder           As FutureBuilder
Attribute mContractFutureBuilder.VB_VarHelpID = -1

Private mSingleContractOnly                         As Boolean

Private mGotAllSampleContracts                      As Boolean
Private mOffsetContractsOutstanding                 As Long

Private mListener                                   As IContractFetchListener

Private mCallerCookie                               As Variant
Private mMyCookie                                   As Variant

Private mFetchFuture                                As IFuture

Private mContractsBuilder                           As New ContractsBuilder

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1
Private WithEvents mOffsetContractFutureWaiter      As FutureWaiter
Attribute mOffsetContractFutureWaiter.VB_VarHelpID = -1
Private WithEvents mSortedContractsFutureWaiter     As FutureWaiter
Attribute mSortedContractsFutureWaiter.VB_VarHelpID = -1

Private mIsOffsetExpiry                             As Boolean
Private mExpiryOffset                               As Long
Private mDaysBeforeExpiryToSwitch                   As Long

Private mSorted                                     As Boolean

Private mSortKeys()                                 As ContractSortKeyIds
Private mSortDescending                             As Boolean

Private mCancelled                                  As Boolean

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
mExpiryOffset = -1
End Sub

Private Sub Class_Terminate()
Const ProcName As String = "Class_Terminate"
GContracts.Logger.Log "ContractFetchTask terminated for: " & mInitialContractSpec.ToString, ProcName, ModuleName, LogLevelDetail
End Sub

'@================================================================================
' IContractFetchListener Interface Members
'@================================================================================

Private Sub IContractFetchListener_FetchCancelled(ByVal pCookie As Variant)
Const ProcName As String = "IContractFetchListener_FetchCancelled"
On Error GoTo Err

If mSingleContractOnly Then
    GContracts.Logger.Log "Contract fetch cancelled for: " & mInitialContractSpec.ToString, ProcName, ModuleName, LogLevelDetail
    mContractFutureBuilder.Cancel
Else
    GContracts.Logger.Log "Contracts fetch cancelled for: " & mInitialContractSpec.ToString, ProcName, ModuleName, LogLevelDetail
    mContractsFutureBuilder.Cancel
End If

If Not mListener Is Nothing Then mListener.FetchCancelled mCallerCookie
mTaskContext.Finish Empty, False
Set mSelfRef = Nothing

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_FetchCompleted(ByVal pCookie As Variant)
Const ProcName As String = "IContractFetchListener_FetchCompleted"
On Error GoTo Err

GContracts.Logger.Log "mOffsetContractsOutstanding=" & mOffsetContractsOutstanding & "; mIsOffsetExpiry=" & mIsOffsetExpiry, ProcName, ModuleName, LogLevelMediumDetail

If Not mIsOffsetExpiry Then
    completeContractProcessing
ElseIf mOffsetContractsOutstanding = 0 Then
    completeContractProcessing
Else
    GContracts.Logger.Log "Got all sample contracts for: " & mContractSpec.ToString, ProcName, ModuleName, LogLevelMediumDetail
    mGotAllSampleContracts = True
End If

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_FetchFailed(ByVal pCookie As Variant, ByVal pErrorCode As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
Const ProcName As String = "IContractFetchListener_FetchFailed"
On Error GoTo Err

If mSingleContractOnly Then
    GContracts.Logger.Log "Contract fetch failed for: " & mInitialContractSpec.ToString & _
                "; (" & pErrorCode & ") " & pErrorMessage, ProcName, ModuleName, LogLevelDetail
    mContractFutureBuilder.Fail pErrorCode, pErrorMessage, pErrorSource
Else
    GContracts.Logger.Log "ContractS fetch failed for: " & mInitialContractSpec.ToString & _
                "; (" & pErrorCode & ") " & pErrorMessage, ProcName, ModuleName, LogLevelDetail
    mContractsFutureBuilder.Fail pErrorCode, pErrorMessage, pErrorSource
End If

If Not mListener Is Nothing Then mListener.FetchFailed mCallerCookie, pErrorCode, pErrorMessage, pErrorSource
mTaskContext.Finish Empty, False
Set mSelfRef = Nothing

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_NotifyContract(ByVal pCookie As Variant, ByVal pContract As IContract)
Const ProcName As String = "IContractFetchListener_NotifyContract"
On Error GoTo Err

If mCancelled Then Exit Sub

If mIsOffsetExpiry Then
    GContracts.Logger.Log "Sample contract received: " & pContract.Specifier.ToString, ProcName, ModuleName, LogLevelMediumDetail
    Dim lResolver As New OffsetContractResolver
    mOffsetContractFutureWaiter.Add lResolver.Resolve(pContract, _
                                        mExpiryOffset, _
                                        mDaysBeforeExpiryToSwitch, _
                                        mCurrentContractStore)
    mOffsetContractsOutstanding = mOffsetContractsOutstanding + 1
Else
    GContracts.Logger.Log "Contract received: " & pContract.Specifier.ToString, ProcName, ModuleName, LogLevelMediumDetail
    processContract pContract
End If

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

GContracts.Logger.Log "Running FetchContractTask for " & mContractSpec.ToString, ProcName, ModuleName, LogLevelMediumDetail

Set mCurrentContractStore = mPrimaryContractStore
If mCurrentContractStore Is Nothing Then
    Set mCurrentContractStore = mSecondaryContractStore
    mUsedSecondaryContractStore = True
End If

Static l As Long
l = l + 1
If l Mod 10 = 0 Then GContracts.Logger.Log "Task counter = " & l, ProcName, ModuleName, LogLevelDetail

If mSorted And Not mIsOffsetExpiry Then
    Set mFetchFuture = mCurrentContractStore.FetchContractsSorted( _
                                                    mContractSpec, _
                                                    mSortKeys, _
                                                    mSortDescending, _
                                                    mMyCookie)
    mSortedContractsFutureWaiter.Add mFetchFuture
Else
    Set mFetchFuture = mCurrentContractStore.FetchContracts( _
                                                    mContractSpec, _
                                                    Me, _
                                                    mMyCookie)
End If

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
' mContractFutureBuilder Event Handlers
'@================================================================================

Private Sub mContractFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mContractFutureBuilder_Cancelled"
On Error GoTo Err

GContracts.Logger.Log "Future builder cancelled: " & mInitialContractSpec.ToString, ProcName, ModuleName, LogLevelDetail

If Not mFetchFuture Is Nothing Then
    If Not mFetchFuture.IsCancelled Then mFetchFuture.Cancel
End If
If Not mListener Is Nothing Then mListener.FetchCancelled mContractFutureBuilder.Cookie
mTaskContext.Finish Empty, True

Set mSelfRef = Nothing

Exit Sub

Err:
GContracts.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mContractsFutureBuilder Event Handlers
'@================================================================================

Private Sub mContractsFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mContractsFutureBuilder_Cancelled"
On Error GoTo Err

GContracts.Logger.Log "Future builder cancelled: " & mInitialContractSpec.ToString, ProcName, ModuleName, LogLevelDetail

If Not mFetchFuture Is Nothing Then
    If Not mFetchFuture.IsCancelled Then mFetchFuture.Cancel
End If
If Not mListener Is Nothing Then mListener.FetchCancelled mContractsFutureBuilder.Cookie
mTaskContext.Finish Empty, True

Set mSelfRef = Nothing

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
    Assert False, "Unexpected cancellation"
ElseIf ev.Future.IsFaulted Then
    GContracts.Logger.Log "Fetch failed: " & mInitialContractSpec.ToString & _
                ": (" & ev.Future.ErrorNumber & "); " & ev.Future.ErrorMessage, _
                ProcName, ModuleName, LogLevelDetail
    If Not mListener Is Nothing Then mListener.FetchFailed _
                                            ev.ContinuationData, _
                                            ev.Future.ErrorNumber, _
                                            ev.Future.ErrorMessage, _
                                            ev.Future.ErrorSource
    mTaskContext.Error ev.Future.ErrorNumber, ev.Future.ErrorSource, ev.Future.ErrorMessage
    Set mSelfRef = Nothing
    Exit Sub
End If

If mIsOffsetExpiry Then
    GContracts.Logger.Log "Sample contract fetch completed: " & mContractSpec.ToString, ProcName, ModuleName, LogLevelMediumDetail
Else
    GContracts.Logger.Log "Fetch completed: " & mInitialContractSpec.ToString, ProcName, ModuleName, LogLevelDetail
    If Not mListener Is Nothing Then mListener.FetchCompleted ev.ContinuationData
    mTaskContext.Finish Empty, False
    Set mSelfRef = Nothing
End If

Exit Sub

Err:
GContracts.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mOffsetContractFutureWaiter Event Handlers
'@================================================================================

Private Sub mOffsetContractFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mOffsetContractFutureWaiter_WaitCompleted"
On Error GoTo Err

Dim lContract As IContract

If ev.Future.IsCancelled Then Exit Sub

If ev.Future.IsFaulted Then
    mOffsetContractFutureWaiter.Clear
    If Not mListener Is Nothing Then mListener.FetchFailed _
                                            mContractsFutureBuilder.Cookie, _
                                            ev.Future.ErrorNumber, _
                                            ev.Future.ErrorMessage, _
                                            ev.Future.ErrorSource
    mTaskContext.Error ev.Future.ErrorNumber, ev.Future.ErrorSource, ev.Future.ErrorMessage
    Set mSelfRef = Nothing
    Exit Sub
End If

Set lContract = ev.Future.Value
If lContract Is Nothing Then
    GContracts.Logger.Log "Required offset contract not found: " & mInitialContractSpec.ToString, ProcName, ModuleName, LogLevelDetail
Else
    GContracts.Logger.Log "Got required offset contract: " & lContract.Specifier.ToString, ProcName, ModuleName, LogLevelDetail
End If

mOffsetContractsOutstanding = mOffsetContractsOutstanding - 1
GContracts.Logger.Log "mOffsetContractsOutstanding=" & mOffsetContractsOutstanding & "; mGotAllSampleContracts=" & mGotAllSampleContracts, ProcName, ModuleName, LogLevelDetail

If Not lContract Is Nothing Then processContract lContract

If mOffsetContractsOutstanding <> 0 Then
ElseIf mGotAllSampleContracts Then
    completeContractProcessing
End If

Exit Sub

Err:
GContracts.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mSortedContractsFutureWaiter Event Handlers
'@================================================================================

Private Sub mSortedContractsFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mSortedContractsFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then

ElseIf ev.Future.IsFaulted Then

Else
    Dim lContracts As IContracts
    Set lContracts = ev.Future.Value

    mFutureWaiter.Add mContractsFutureBuilder.Future, mCallerCookie
    If mContractsFutureBuilder.Future.IsPending Then
        mContractsFutureBuilder.Value = lContracts
        mContractsFutureBuilder.Complete
    End If
    
    GContracts.Logger.Log "Sorted contract fetch completed for: " & mInitialContractSpec.ToString, ProcName, ModuleName, LogLevelDetail
End If

Exit Sub

Err:
GContracts.NotifyUnhandledError ProcName, ModuleName
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
GContracts.HandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ContractsFuture() As IFuture
Const ProcName As String = "ContractsFuture"
On Error GoTo Err

Set ContractsFuture = mContractsFutureBuilder.Future

Exit Property

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
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

GContracts.Logger.Log "Fetching contract for " & pContractSpec.ToString, ProcName, ModuleName, LogLevelDetail

Set mSelfRef = Me

Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore

GContracts.SetVariant mCallerCookie, pCookie

Set mListener = pListener

mSingleContractOnly = pSingleContractOnly
createFutureBuilder

Set mInitialContractSpec = pContractSpec
mContractsBuilder.ContractSpecifier = mInitialContractSpec

If Not GContractUtils.IsOffsetExpiry(pContractSpec.Expiry) Then
    Set mContractSpec = mInitialContractSpec
Else
    GContracts.Logger.Log "Contract has offset expiry: " & pContractSpec.Expiry, ProcName, ModuleName, LogLevelDetail
    Set mOffsetContractFutureWaiter = New FutureWaiter
    mIsOffsetExpiry = True
    Dim lErrorMessage As String
    Assert GContractUtils.TryParseOffsetExpiry( _
                    mInitialContractSpec.Expiry, _
                    mExpiryOffset, _
                    mDaysBeforeExpiryToSwitch, _
                    lErrorMessage), _
            lErrorMessage
    With mInitialContractSpec
        Set mContractSpec = GContractUtils.CreateContractSpecifier(.LocalSymbol, _
                                                    .Symbol, _
                                                    .TradingClass, _
                                                    .Exchange, _
                                                    .SecType, _
                                                    .CurrencyCode, _
                                                    GContractUtils.GetCurrentQuarterExpiry, _
                                                    .Multiplier, _
                                                    .Strike, _
                                                    .Right)
        GContracts.Logger.Log "Fetch sample contract(s) for: " & mContractSpec.ToString, ProcName, ModuleName, LogLevelDetail
    End With
End If

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub
                
Friend Sub InitialiseSorted( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore, _
                ByVal pCookie As Variant, _
                ByRef pSortkeys() As ContractSortKeyIds, _
                ByVal pSortDescending As Boolean)
Const ProcName As String = "InitialiseSorted"
On Error GoTo Err

mSorted = True

Dim n As Long: n = -1
On Error Resume Next
    n = UBound(pSortkeys)
On Error GoTo Err
AssertArgument n >= 0, "Empty sort keys"

mSortKeys = pSortkeys
mSortDescending = pSortDescending

Set mSortedContractsFutureWaiter = New FutureWaiter

Initialise pContractSpec, _
    pPrimaryContractStore, _
    pSecondaryContractStore, _
    pCookie, _
    Nothing, _
    False

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub
                
'@================================================================================
' Helper Functions
'@================================================================================

Private Sub completeContractProcessing()
Const ProcName As String = "completeContractProcessing"
On Error GoTo Err

Dim lContracts As IContracts: Set lContracts = mContractsBuilder.Contracts

If lContracts.Count = 0 And Not mUsedSecondaryContractStore Then
    If trySecondaryContractStore Then Exit Sub
End If

If mSingleContractOnly Then
    mFutureWaiter.Add mContractFutureBuilder.Future, mCallerCookie
    If mContractFutureBuilder.Future.IsPending Then
        If lContracts.Count = 0 Then
            mContractFutureBuilder.Value = Nothing
        ElseIf lContracts.Count = 1 Then
            mContractFutureBuilder.Value = lContracts.ItemAtIndex(1)
        End If
        mContractFutureBuilder.Complete
    End If
Else
    mFutureWaiter.Add mContractsFutureBuilder.Future, mCallerCookie
    If mContractsFutureBuilder.Future.IsPending Then
        mContractsFutureBuilder.Value = lContracts
        mContractsFutureBuilder.Complete
    End If
End If

GContracts.Logger.Log "Contract fetch completed for: " & mInitialContractSpec.ToString, ProcName, ModuleName, LogLevelMediumDetail

Set mSelfRef = Nothing

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createFutureBuilder()
Const ProcName As String = "createFutureBuilder"
On Error GoTo Err

If mSingleContractOnly Then
    Set mContractFutureBuilder = New FutureBuilder
    mContractFutureBuilder.Cookie = mCallerCookie
Else
    Set mContractsFutureBuilder = New FutureBuilder
    mContractsFutureBuilder.Cookie = mCallerCookie
End If

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

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
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processContract(ByVal pContract As IContract)
Const ProcName As String = "processContract"
On Error GoTo Err

If mCancelled Then Exit Sub

If mContractsBuilder.Contracts.Count = 1 And mSingleContractOnly Then
    mCancelled = True
    GContracts.Logger.Log "Contract fetch error for: " & mInitialContractSpec.ToString & _
                "; contract is not uniquely specified", ProcName, ModuleName, LogLevelDetail
    mContractFutureBuilder.Fail ErrorCodes.ErrIllegalArgumentException, "Contract is not uniquely specified", ""
    mFetchFuture.Cancel
    mTaskContext.Finish Empty, False
    Set mSelfRef = Nothing
    Exit Sub
End If

mContractsBuilder.Add pContract

If Not mListener Is Nothing Then mListener.NotifyContract mCallerCookie, pContract

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function trySecondaryContractStore() As Boolean
Const ProcName As String = "trySecondaryContractStore"
On Error GoTo Err

mUsedSecondaryContractStore = True
If mSecondaryContractStore Is Nothing Then
    trySecondaryContractStore = False
Else
    Set mCurrentContractStore = mSecondaryContractStore
    Set mFetchFuture = mCurrentContractStore.FetchContracts(mContractSpec, Me, mMyCookie)
    trySecondaryContractStore = True
End If

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function




