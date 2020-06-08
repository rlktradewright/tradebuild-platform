VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractDetailsRequester"
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

Implements IContractDetailsConsumer
Implements ITwsConnectionStateListener
Implements ITask

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum States
    StateFree
    StateAllocated
    StateRequestQueued
    StateRequestInProgress
    StateRequestCancelling
    StateRequestCompleted
End Enum

'@================================================================================
' Types
'@================================================================================

Private Type ContractDetailsRequestTableEntry
    Index                   As Long ' of this entry
    State                   As States
    Contract                As TwsContract
    Requester               As IContractDetailsConsumer
    NumberReturned          As Long
    PendingContracts        As EnumerableCollection
    Next                    As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ContractDetailsRequester"

Private Const MaxConcurrentContractDetailsRequests  As Long = 20

'@================================================================================
' Member variables
'@================================================================================

Private mTwsAPI                                     As TwsAPI
Attribute mTwsAPI.VB_VarHelpID = -1

Private mRequests()                                 As ContractDetailsRequestTableEntry
Private mRequestsIndex                              As Long
Private mFirstFreeRequestsEntry                     As Long
Private mFirstPendingRequestsEntry                  As Long
Private mLastPendingRequestsEntry                   As Long

Private mFirstPendingContractsEntry                 As Long
Private mLastPendingContractsEntry                  As Long

Private mTaskContext                                As TaskContext

Private mConcurrentRequests                         As Long
Private mMaxConcurrentRequests                      As Long

Private mContractCache                              As TwsContractCache

Private mIsTwsConnectedToIBServers                  As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mRequests(7) As ContractDetailsRequestTableEntry
mFirstFreeRequestsEntry = NullIndex
mFirstPendingRequestsEntry = NullIndex
mLastPendingRequestsEntry = NullIndex
mFirstPendingContractsEntry = NullIndex
mLastPendingContractsEntry = NullIndex
End Sub

'@================================================================================
' IContractDetailsConsumer Interface Members
'@================================================================================

Private Sub IContractDetailsConsumer_EndContractDetails(ByVal pRequestId As Long)
Const ProcName As String = "IContractDetailsConsumer_EndContractDetails"
On Error GoTo Err

If mRequests(pRequestId).State = StateRequestCancelling Then
    releaseEntry pRequestId
ElseIf mRequests(pRequestId).NumberReturned = 1 Then
    mRequests(pRequestId).Requester.EndContractDetails pRequestId
    releaseEntry pRequestId
Else
    mRequests(pRequestId).State = StateRequestCompleted
End If
CompleteRequest

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyContract(ByVal pRequestId As Long, pContractDetails As TwsContractDetails)
Const ProcName As String = "IContractDetailsConsumer_NotifyContract"
On Error GoTo Err

If mRequests(pRequestId).NumberReturned = 0 Then
    mRequests(pRequestId).NumberReturned = 1
    mContractCache.Add pContractDetails
    If mRequests(pRequestId).State <> StateRequestCancelling Then mRequests(pRequestId).Requester.NotifyContract pRequestId, pContractDetails
Else
    mRequests(pRequestId).NumberReturned = mRequests(pRequestId).NumberReturned + 1
    If mRequests(pRequestId).PendingContracts Is Nothing Then
        Set mRequests(pRequestId).PendingContracts = New EnumerableCollection
        If mFirstPendingContractsEntry = NullIndex Then mFirstPendingContractsEntry = pRequestId
        If mLastPendingContractsEntry <> NullIndex Then mRequests(mLastPendingContractsEntry).Next = pRequestId
        mLastPendingContractsEntry = pRequestId
        mRequests(mLastPendingContractsEntry).Next = NullIndex
    End If
    mRequests(pRequestId).PendingContracts.Add pContractDetails
    mTaskContext.Continue
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IContractDetailsConsumer_NotifyError"
On Error GoTo Err

If mRequests(pRequestId).State <> StateRequestCancelling Then mRequests(pRequestId).Requester.NotifyError pRequestId, pErrorCode, pErrorMsg
CompleteRequest

If mRequests(pRequestId).NumberReturned > 1 Then
    removeEntryFromPendingContractsQueue pRequestId
End If

releaseEntry pRequestId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITwsConnectionStateListener Interface Members
'@================================================================================

Private Sub ITwsConnectionStateListener_NotifyAPIConnectionStateChange(ByVal pSource As Object, ByVal pState As ApiConnectionStates, ByVal pMessage As String)
Const ProcName As String = "ITwsConnectionStateListener_NotifyAPIConnectionStateChange"
On Error GoTo Err

Select Case pState
Case ApiConnNotConnected

Case ApiConnConnecting

Case ApiConnConnected
    ' we'll assume that TWS has a server connection: if not,
    ' it will soon tell us
    mIsTwsConnectedToIBServers = True
    
    doSchedule
Case ApiConnFailed

End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionClosed(ByVal pSource As Object)
mIsTwsConnectedToIBServers = False
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionRecovered(ByVal pSource As Object, ByVal pDataLost As Boolean)
Const ProcName As String = "ITwsConnectionStateListener_NotifyIBServerConnectionRecovered"
On Error GoTo Err

mIsTwsConnectedToIBServers = True
doSchedule

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

doSchedule

If mFirstPendingContractsEntry <> NullIndex Then
    Dim lPendingContractsEntryIndex As Long
    lPendingContractsEntryIndex = mFirstPendingContractsEntry
    
    With mRequests(lPendingContractsEntryIndex)
        Dim en As Enumerator: Set en = .PendingContracts.Enumerator
        If en.MoveNext Then
            Dim lContractDetails As TwsContractDetails
            Set lContractDetails = en.Current
            mContractCache.Add lContractDetails
            If .State <> StateRequestCancelling Then
                .Requester.NotifyContract lPendingContractsEntryIndex, lContractDetails
            End If
            en.Remove
        End If
        If .PendingContracts.Count = 0 And _
            .State = StateRequestCompleted _
        Then
            If .State <> StateRequestCancelling Then
                .Requester.EndContractDetails lPendingContractsEntryIndex
            End If
            removeEntryFromPendingContractsQueue lPendingContractsEntryIndex
            releaseEntry lPendingContractsEntryIndex
        End If
    End With
End If
 
If mFirstPendingRequestsEntry = NullIndex And _
    mFirstPendingContractsEntry = NullIndex _
Then mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub cancelRequest(ByVal pRequestId As Long)
Const ProcName As String = "CancelRequest"
On Error GoTo Err

Select Case mRequests(pRequestId).State
Case StateFree

Case StateAllocated
    releaseEntry pRequestId
Case StateRequestQueued
    unqueueEntry pRequestId
    releaseEntry pRequestId
Case StateRequestInProgress
    mRequests(pRequestId).State = StateRequestCancelling
Case StateRequestCancelling

End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTwsAPI As TwsAPI, _
                ByVal pContractCache As TwsContractCache)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTwsAPI = pTwsAPI
Set mContractCache = pContractCache

mMaxConcurrentRequests = MaxConcurrentContractDetailsRequests

gLog "Max concurrent contract details requests", ModuleName, ProcName, CStr(mMaxConcurrentRequests), LogLevelDetail

StartTask Me, PriorityNormal

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function RequestContractDetails( _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pRequester As IContractDetailsConsumer) As Long
Const ProcName As String = "RequestContractDetails"
On Error GoTo Err

Dim lIndex As Long
lIndex = allocateEntry
Set mRequests(lIndex).Contract = gContractSpecToTwsContract(pContractSpecifier)
Set mRequests(lIndex).Requester = pRequester

queueEntry lIndex

Schedule

RequestContractDetails = lIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Schedule()
Const ProcName As String = "Schedule"
On Error GoTo Err

mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function allocateEntry() As Long
Const ProcName As String = "allocateEntry"
On Error GoTo Err

If mRequestsIndex < UBound(mRequests) Then
    mRequestsIndex = mRequestsIndex + 1
    allocateEntry = mRequestsIndex
ElseIf mFirstFreeRequestsEntry <> NullIndex Then
    allocateEntry = mFirstFreeRequestsEntry
    mFirstFreeRequestsEntry = mRequests(mFirstFreeRequestsEntry).Next
Else
    ReDim Preserve mRequests(2 * (UBound(mRequests) + 1) - 1) As ContractDetailsRequestTableEntry
    mRequestsIndex = mRequestsIndex + 1
    allocateEntry = mRequestsIndex
End If

mRequests(allocateEntry).State = StateAllocated
mRequests(allocateEntry).Index = allocateEntry

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub CompleteRequest()
Const ProcName As String = "completeRequest"
On Error GoTo Err

If mConcurrentRequests > 0 Then mConcurrentRequests = mConcurrentRequests - 1
Schedule

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doSchedule()
Const ProcName As String = "doSchedule"
On Error GoTo Err

If mFirstPendingRequestsEntry = NullIndex Then Exit Sub

gLog "Scheduling contract requests", ModuleName, ProcName, , LogLevelDetail

If mTwsAPI.ConnectionState <> TwsConnConnected Then
    gLog "Can't schedule contract requests: not connected to TWS", ModuleName, ProcName, , LogLevelDetail
    Exit Sub
End If

If Not mIsTwsConnectedToIBServers Then
    gLog "Can't schedule contract requests: TWS not connected to IB servers", ModuleName, ProcName, , LogLevelDetail
    Exit Sub
End If

Dim lIndex As Long: lIndex = mFirstPendingRequestsEntry
Dim lCurrIndex As Long

Do While lIndex <> NullIndex And mConcurrentRequests < mMaxConcurrentRequests
    lCurrIndex = lIndex
    lIndex = mRequests(lIndex).Next
    If gLogger.IsLoggable(LogLevelDetail) Then gLog "Requesting contract details for ", ModuleName, ProcName, mRequests(lCurrIndex).Contract.ToString, LogLevelDetail
    mTwsAPI.RequestContractDetails lCurrIndex, mRequests(lCurrIndex).Contract
        
    mConcurrentRequests = mConcurrentRequests + 1
        
    unqueueEntry lCurrIndex
    mRequests(lCurrIndex).State = StateRequestInProgress
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub queueEntry(ByVal pIndex As Long)
Const ProcName As String = "queueEntry"
On Error GoTo Err

mRequests(pIndex).State = StateRequestQueued
If mFirstPendingRequestsEntry = NullIndex Then mFirstPendingRequestsEntry = pIndex
If mLastPendingRequestsEntry <> NullIndex Then mRequests(mLastPendingRequestsEntry).Next = pIndex
mLastPendingRequestsEntry = pIndex
mRequests(mLastPendingRequestsEntry).Next = NullIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function releaseEntry(ByVal pIndex As Long)
Const ProcName As String = "releaseEntry"
On Error GoTo Err

mRequests(pIndex).Next = mFirstFreeRequestsEntry
mFirstFreeRequestsEntry = pIndex
mRequests(pIndex).State = StateFree
Set mRequests(pIndex).Contract = Nothing
Set mRequests(pIndex).Requester = Nothing
mRequests(pIndex).NumberReturned = 0
Set mRequests(pIndex).PendingContracts = Nothing

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub removeEntryFromPendingContractsQueue(ByVal pIndexToRelease As Long)
Dim lIndex As Long: lIndex = mFirstPendingContractsEntry
Dim lPrevIndex As Long: lPrevIndex = NullIndex
Do While lIndex <> NullIndex
    If lIndex = pIndexToRelease Then
        If lPrevIndex <> NullIndex Then
            mRequests(lPrevIndex).Next = mRequests(pIndexToRelease).Next
        Else
            mFirstPendingContractsEntry = mRequests(pIndexToRelease).Next
        End If
        Exit Do
    End If
    lPrevIndex = lIndex
    lIndex = mRequests(lIndex).Next
Loop
End Sub

Private Sub unqueueEntry(ByVal pIndex As Long)
Const ProcName As String = "unqueueEntry"
On Error GoTo Err

Dim lIndex As Long
Dim lPrevIndex As Long

Assert mFirstPendingRequestsEntry <> NullIndex And mLastPendingRequestsEntry <> NullIndex, "No pending entries"

lIndex = mFirstPendingRequestsEntry
lPrevIndex = NullIndex
Do While lIndex <> pIndex
    lPrevIndex = lIndex
    lIndex = mRequests(lIndex).Next
    Assert lIndex <> NullIndex, "Can't find entry to be removed"
Loop

If lIndex = mFirstPendingRequestsEntry Then mFirstPendingRequestsEntry = mRequests(lIndex).Next
If lIndex = mLastPendingRequestsEntry Then mLastPendingRequestsEntry = lPrevIndex

If lPrevIndex <> NullIndex Then mRequests(lPrevIndex).Next = mRequests(lIndex).Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




