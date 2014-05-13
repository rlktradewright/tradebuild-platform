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
Implements Task

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
End Enum

'@================================================================================
' Types
'@================================================================================

Private Type ContractDetailsRequestTableEntry
    State                   As States
    Contract                As TwsContract
    Requester               As IContractDetailsConsumer
    Next                    As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ContractDetailsRequester"

Private Const MaxConcurrentContractDetailsRequests  As Long = 5

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

Private mTaskContext                                As TaskContext

Private mConcurrentRequests                         As Long
Private mMaxConcurrentRequests                      As Long

Private mContractCache                              As TwsContractCache

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mRequests(7) As ContractDetailsRequestTableEntry
mFirstFreeRequestsEntry = NullIndex
mFirstPendingRequestsEntry = NullIndex
mLastPendingRequestsEntry = NullIndex
End Sub

'@================================================================================
' IContractDetailsConsumer Interface Members
'@================================================================================

Private Sub IContractDetailsConsumer_EndContractDetails(ByVal pRequestId As Long)
Const ProcName As String = "IContractDetailsConsumer_EndContractDetails"
On Error GoTo Err

If mRequests(pRequestId).State <> StateRequestCancelling Then mRequests(pRequestId).Requester.EndContractDetails pRequestId
CompleteRequest pRequestId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyContract(ByVal pRequestId As Long, pContractDetails As TwsContractDetails)
Const ProcName As String = "IContractDetailsConsumer_NotifyContract"
On Error GoTo Err

mContractCache.Add pContractDetails
If mRequests(pRequestId).State <> StateRequestCancelling Then mRequests(pRequestId).Requester.NotifyContract pRequestId, pContractDetails

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IContractDetailsConsumer_NotifyError"
On Error GoTo Err

If mRequests(pRequestId).State <> StateRequestCancelling Then mRequests(pRequestId).Requester.NotifyError pRequestId, pErrorCode, pErrorMsg
CompleteRequest pRequestId

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
    doSchedule
Case ApiConnFailed

End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionClosed(ByVal pSource As Object)

End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionRecovered(ByVal pSource As Object, ByVal pDataLost As Boolean)

End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()

End Sub

Private Sub Task_Run()
Const ProcName As String = "Task_Run"
On Error GoTo Err

doSchedule

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let Task_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.Name
End Property

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub CancelRequest(ByVal pRequestId As Long)
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

RequestContractDetails = lIndex

Schedule

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

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub CompleteRequest(ByVal pRequestId As Long)
Const ProcName As String = "completeRequest"
On Error GoTo Err

If mConcurrentRequests > 0 Then mConcurrentRequests = mConcurrentRequests - 1
releaseEntry pRequestId
Schedule

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doSchedule()
Const ProcName As String = "doSchedule"
On Error GoTo Err

Dim lIndex As Long
Dim lCurrIndex As Long

gLog "Scheduling contract requests", ModuleName, ProcName, , LogLevelDetail

If mTwsAPI.ConnectionState <> TwsConnConnected Then
    gLog "Scheduling contract requests", ModuleName, ProcName, , LogLevelDetail
    mTaskContext.Suspend -1
    Exit Sub
End If

lIndex = mFirstPendingRequestsEntry

Do While lIndex <> NullIndex And mConcurrentRequests < mMaxConcurrentRequests
    lCurrIndex = lIndex
    lIndex = mRequests(lIndex).Next
    If gLogger.IsLoggable(LogLevelDetail) Then gLog "Requesting contract details for", ModuleName, ProcName, mRequests(lCurrIndex).Contract.ToString, LogLevelDetail
    mTwsAPI.RequestContractDetails lCurrIndex, mRequests(lCurrIndex).Contract
        
    mConcurrentRequests = mConcurrentRequests + 1
        
    unqueueEntry lCurrIndex
    mRequests(lCurrIndex).State = StateRequestInProgress
Loop

mTaskContext.Suspend -1

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

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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




