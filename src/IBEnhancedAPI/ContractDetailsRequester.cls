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
Implements Task

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

Private Const ModuleName                            As String = "ContractDetailsRequester"

Private Const MaxConcurrentContractDetailsRequests  As Long = 5

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTwsAPI                          As TwsAPI
Attribute mTwsAPI.VB_VarHelpID = -1

Private mRequests()                                 As ContractDetailsRequestTableEntry
Private mRequestsIndex                              As Long
Private mFirstFreeRequestsEntry                     As Long
Private mFirstPendingRequestsEntry                  As Long
Private mLastPendingRequestsEntry                   As Long

Private mTaskContext                                As TaskContext

Private mConcurrentRequests                         As Long
Private mMaxConcurrentRequests                      As Long

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

mRequests(pRequestId).Requester.EndContractDetails pRequestId

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyContract(ByVal pRequestId As Long, pContractDetails As TwsContractDetails)
Const ProcName As String = "IContractDetailsConsumer_NotifyContract"
On Error GoTo Err

mRequests(pRequestId).Requester.NotifyContract pRequestId, pContractDetails

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IContractDetailsConsumer_NotifyError"
On Error GoTo Err

mRequests(pRequestId).Requester.NotifyError pRequestId, pErrorCode, pErrorMsg

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Let Task_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.Name
End Property

'@================================================================================
' mTwsAPI Event Handlers
'@================================================================================

Private Sub mTwsAPI_Connected()
Const ProcName As String = "mTwsAPI_Connected"
On Error GoTo Err

If mTwsAPI.SerialiseContractDetailsRequests Then
    mMaxConcurrentRequests = 1
Else
    mMaxConcurrentRequests = MaxConcurrentContractDetailsRequests
End If

doSchedule

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub CompleteRequest(ByVal pRequestId As Long)
Const ProcName As String = "CompleteRequest"
On Error GoTo Err

If mConcurrentRequests > 0 Then mConcurrentRequests = mConcurrentRequests - 1
releaseEntry pRequestId
Schedule

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise(ByVal pTwsAPI As TwsAPI)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTwsAPI = pTwsAPI

StartTask Me, PriorityNormal

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Function RequestContractDetails( _
                ByVal pContractSpecifier As ContractSpecifier, _
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub Schedule()
Const ProcName As String = "Schedule"
On Error GoTo Err

mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub doSchedule()
Const ProcName As String = "doSchedule"
On Error GoTo Err

Dim lIndex As Long
Dim lCurrIndex As Long

If mTwsAPI.ConnectionState <> ConnConnected Then
    mTaskContext.Suspend -1
    Exit Sub
End If

lIndex = mFirstPendingRequestsEntry

Do While lIndex <> NullIndex And mConcurrentRequests < mMaxConcurrentRequests
    lCurrIndex = lIndex
    lIndex = mRequests(lIndex).Next
    mTwsAPI.RequestContractDetails lCurrIndex, mRequests(lCurrIndex).Contract
        
    mConcurrentRequests = mConcurrentRequests + 1
        
    unqueueEntry lCurrIndex
Loop

mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub queueEntry(ByVal pIndex As Long)
Const ProcName As String = "queueEntry"
On Error GoTo Err

If mFirstPendingRequestsEntry = NullIndex Then mFirstPendingRequestsEntry = pIndex
If mLastPendingRequestsEntry <> NullIndex Then mRequests(mLastPendingRequestsEntry).Next = pIndex
mLastPendingRequestsEntry = pIndex
mRequests(mLastPendingRequestsEntry).Next = NullIndex

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function releaseEntry(ByVal pIndex As Long)
Const ProcName As String = "releaseEntry"
On Error GoTo Err

mRequests(pIndex).Next = mFirstFreeRequestsEntry
mFirstFreeRequestsEntry = pIndex
Set mRequests(pIndex).Contract = Nothing
Set mRequests(pIndex).Requester = Nothing

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub unqueueEntry(ByVal pIndex As Long)
Const ProcName As String = "unqueueEntry"
On Error GoTo Err

Dim lIndex As Long
Dim lPrevIndex As Long

If mFirstPendingRequestsEntry = NullIndex Or _
    mLastPendingRequestsEntry = NullIndex Then Err.Raise ErrorCodes.ErrIllegalStateException, , "No pending entries"

lIndex = mFirstPendingRequestsEntry
lPrevIndex = NullIndex
Do While lIndex <> pIndex
    lPrevIndex = lIndex
    lIndex = mRequests(lIndex).Next
    If lIndex = NullIndex Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Can't find entry to be removed"
Loop

If lIndex = mFirstPendingRequestsEntry Then mFirstPendingRequestsEntry = mRequests(lIndex).Next
If lIndex = mLastPendingRequestsEntry Then mLastPendingRequestsEntry = lPrevIndex

If lPrevIndex <> NullIndex Then mRequests(lPrevIndex).Next = mRequests(lIndex).Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub




