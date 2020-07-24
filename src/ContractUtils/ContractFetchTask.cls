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

Private Const CookieSampleContractFetch             As String = "Sample"
Private Const CookieTargetContractFetch             As String = "Target"

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

Private mGotSampleContract                          As Boolean

Private mListener                                   As IContractFetchListener

Private mCallerCookie                               As Variant
Private mMyCookie                                   As Variant

Private mFetchFuture                                As IFuture

Private mContractsBuilder                           As New ContractsBuilder

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mIsOffsetExpiry                             As Boolean
Private mExpiryOffset                               As Long
Private mDaysBeforeExpiryToSwitch                   As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter

mExpiryOffset = -1

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

If Not mListener Is Nothing Then mListener.FetchCancelled mCallerCookie
mTaskContext.Finish Empty, False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_FetchCompleted(ByVal pCookie As Variant)
Const ProcName As String = "IContractFetchListener_FetchCompleted"
On Error GoTo Err

If Not futureIsPending Then Exit Sub
If pCookie = CookieSampleContractFetch Then
    If mGotSampleContract Then Exit Sub
ElseIf mIsOffsetExpiry Then
    processOffsetContracts
End If

completeContractProcessing

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

If Not mListener Is Nothing Then mListener.FetchFailed mCallerCookie, pErrorCode, pErrorMessage, pErrorSource
mTaskContext.Finish Empty, False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_NotifyContract(ByVal pCookie As Variant, ByVal pContract As IContract)
Const ProcName As String = "IContractFetchListener_NotifyContract"
On Error GoTo Err

If Not futureIsPending Then Exit Sub

If pCookie = CookieSampleContractFetch Then
    If Not mGotSampleContract Then
        mGotSampleContract = True
        processSampleContract pContract
    End If
ElseIf mIsOffsetExpiry Then
    mContractsBuilder.Add pContract
Else
    If mContractsBuilder.Contracts.Count = 1 And mSingleContractOnly Then
        mContractFutureBuilder.Fail ErrorCodes.ErrIllegalArgumentException, "Contract is not uniquely specified", ""
        mFetchFuture.Cancel
        mTaskContext.Finish Empty, False
        Exit Sub
    End If
    
    mContractsBuilder.Add pContract
    
    If Not mListener Is Nothing Then mListener.NotifyContract mCallerCookie, pContract
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

Set mFetchFuture = lContractStore.FetchContracts(mContractSpec, Me, mMyCookie)

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
' mContractFutureBuilder Event Handlers
'@================================================================================

Private Sub mContractFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mContractFutureBuilder_Cancelled"
On Error GoTo Err

If Not mFetchFuture Is Nothing Then
    If Not mFetchFuture.IsCancelled Then mFetchFuture.Cancel
End If
If Not mListener Is Nothing Then mListener.FetchCancelled mContractFutureBuilder.Cookie
mTaskContext.Finish Empty, True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mContractsFutureBuilder Event Handlers
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
' mFutureWaiter Event Handlers
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

Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore

gSetVariant mCallerCookie, pCookie

Set mListener = pListener

mSingleContractOnly = pSingleContractOnly
createFutureBuilder

If Not gIsOffsetExpiry(pContractSpec.Expiry) Then
    mMyCookie = CookieTargetContractFetch
    Set mContractSpec = pContractSpec
Else
    mIsOffsetExpiry = True
    mMyCookie = CookieSampleContractFetch
    gParseOffsetExpiry pContractSpec.Expiry, _
                    mExpiryOffset, _
                    mDaysBeforeExpiryToSwitch
    With pContractSpec
        Set mContractSpec = gCreateContractSpecifier(.LocalSymbol, _
                                                    .Symbol, _
                                                    .Exchange, _
                                                    .SecType, _
                                                    .CurrencyCode, _
                                                    "", _
                                                    .Multiplier, _
                                                    .Strike, _
                                                    .Right)
    End With
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
Friend Sub InitialiseSorted( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore, _
                ByVal pCookie As Variant, _
                ByVal pListener As IContractFetchListener, _
                ByRef pSortkeys() As ContractSortKeyIds, _
                ByVal pSortDescending As Boolean)
Const ProcName As String = "InitialiseSorted"
On Error GoTo Err

Dim n As Long: n = -1
On Error Resume Next
    n = UBound(pSortkeys)
On Error GoTo Err
AssertArgument n >= 0, "Empty sort keys"

Dim ar() As ContractSortKeyIds: ar = pSortkeys

mContractsBuilder.SortKeys = ar
mContractsBuilder.SortAscending = Not pSortDescending

Initialise pContractSpec, _
    pPrimaryContractStore, _
    pSecondaryContractStore, _
    pCookie, _
    pListener, _
    False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
'@================================================================================
' Helper Functions
'@================================================================================

Private Sub completeContractProcessing()
Const ProcName As String = "completeContractProcessing"
On Error GoTo Err

Dim lContracts As IContracts
Set lContracts = mContractsBuilder.Contracts
If lContracts.Count = 0 And Not mUsedSecondaryContractStore Then If trySecondaryContractSP Then Exit Sub

If mSingleContractOnly Then
    mFutureWaiter.Add mContractFutureBuilder.Future, mCallerCookie
    If lContracts.Count = 0 Then
        mContractFutureBuilder.Fail ErrorCodes.ErrIllegalArgumentException, "No such contract", ""
    Else
        mContractFutureBuilder.Value = lContracts.ItemAtIndex(1)
        mContractFutureBuilder.Complete
    End If
Else
    mFutureWaiter.Add mContractsFutureBuilder.Future, mCallerCookie
    mContractsFutureBuilder.Value = lContracts
    mContractsFutureBuilder.Complete
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processSampleContract(ByVal pFirstContract As IContract)

Const ProcName As String = "processSampleContract"
On Error GoTo Err

' we've now got a contract that meets the specified
' contract spec. Now we need to find all the contracts that do
' so (apart from the expiry date)
With pFirstContract.Specifier
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

mMyCookie = CookieTargetContractFetch
Set mContractsBuilder = New ContractsBuilder
ReDim lSortkeys(0) As ContractSortKeyIds
lSortkeys(0) = ContractSortKeyExpiry
mContractsBuilder.SortKeys = lSortkeys

mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processOffsetContracts()
Const ProcName As String = "processOffsetContracts"
On Error GoTo Err

If mExpiryOffset <> -1 Then
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
        If lIndex < mExpiryOffset Then
        ElseIf GetOffsetSessionTimes( _
                            lContract.ExpiryDate, _
                            -mDaysBeforeExpiryToSwitch).StartTime >= Int(CDbl(Now)) _
        Then
            ' we've found the contract we're after
            mContractsBuilder.Add lContract
            If Not mListener Is Nothing Then mListener.NotifyContract mCallerCookie, lContract
            Exit For
        End If
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function trySecondaryContractSP() As Boolean
Const ProcName As String = "trySecondaryContractSP"
On Error GoTo Err

mUsedSecondaryContractStore = True
If mSecondaryContractStore Is Nothing Then
    trySecondaryContractSP = False
Else
    Set mFetchFuture = mSecondaryContractStore.FetchContracts(mContractSpec, Me, mMyCookie)
    trySecondaryContractSP = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




