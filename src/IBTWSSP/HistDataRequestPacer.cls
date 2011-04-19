VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataRequestPacer"
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

Implements StateChangeListener
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

Private Const ModuleName                            As String = "HistDataRequestPacer"

'Private Const MaxConcurrentHistDataRequests        As Long = 50 ' limit imposed by IB
Private Const MaxConcurrentHistDataRequests         As Long = 5 ' try serialising them

'@================================================================================
' Member variables
'@================================================================================

Private mTWSAPI                                     As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1

Private mRequests()                                 As HistoricalDataRequestTableEntry
Private mRequestsIndex                              As Long
Private mFirstFreeRequestsEntry                     As Long
Private mFirstPendingRequestsEntry                  As Long
Private mLastPendingRequestsEntry                   As Long
Private mLastPriorityPendingRequestsEntry           As Long

Private mTaskContext                                As TaskContext

Private mConcurrentHistDataRequests                 As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mRequests(7) As HistoricalDataRequestTableEntry
mFirstFreeRequestsEntry = NullIndex
mFirstPendingRequestsEntry = NullIndex
mLastPendingRequestsEntry = NullIndex
mLastPriorityPendingRequestsEntry = NullIndex

StartTask Me, PriorityNormal
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Dim tli As TimerListItem

Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

Set tli = ev.source
releaseEntry CLng(tli.data)
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

Private Property Let Task_TaskContext(ByVal RHS As TWUtilities30.TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.Name
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub CancelHistoricalDataRequest( _
                ByVal pRequestId As Long)
Const ProcName As String = "CancelHistoricalDataRequest"
On Error GoTo Err

If mRequests(pRequestId).Cancelling Then Exit Sub
mRequests(pRequestId).Cancelling = True

If mRequests(pRequestId).ActualSubmissionTime = 0 Then
    unqueueEntry pRequestId
    releaseEntry pRequestId
Else
    mTWSAPI.CancelHistoricalData mRequests(pRequestId).Request
    
    ' it's possible that TWS has already sent the data but we haven't received it yet,
    ' so we keep this entry for 'long enough' to ensure that any pending data
    ' has been received - ie 60 seconds
    GetGlobalTimerList.Add(pRequestId, 60).AddStateChangeListener Me
End If

Schedule

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub CompleteRequest(ByVal pRequestId As Long)
Const ProcName As String = "CompleteRequest"
On Error GoTo Err

If mConcurrentHistDataRequests > 0 Then mConcurrentHistDataRequests = mConcurrentHistDataRequests - 1
releaseEntry pRequestId
Schedule

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Finish()
Set mTWSAPI = Nothing
If Not mTaskContext Is Nothing Then mTaskContext.Finish Empty, False
End Sub

Friend Function GetRequester(ByVal pIndex As Long) As IHistDataRequester
Const ProcName As String = "GetRequester"
On Error GoTo Err

Set GetRequester = mRequests(pIndex).Request.Requester

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub Initialise(ByVal pTWSAPI As TWSAPI)
Set mTWSAPI = pTWSAPI
End Sub

Friend Function RequestHistoricalData( _
                ByRef pRequest As HistoricalDataRequest, _
                ByVal pPriority As Boolean) As Long
Const ProcName As String = "RequestHistoricalData"
On Error GoTo Err

Dim lIndex As Long

lIndex = allocateEntry
mRequests(lIndex).Request = pRequest
mRequests(lIndex).Request.Id = lIndex

queueEntry lIndex, pPriority

RequestHistoricalData = lIndex

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
    ReDim Preserve mRequests(2 * (UBound(mRequests) + 1) - 1) As HistoricalDataRequestTableEntry
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
Dim lDelay As Long
Dim lTimestamp As Date
Dim lEarliest As Date

If mTWSAPI.ConnectionState <> ConnConnected Then
    mTaskContext.Suspend -1
    Exit Sub
End If

lTimestamp = GetTimestamp
lEarliest = MaxDate
lIndex = mFirstPendingRequestsEntry
lDelay = gDelayTillNextSubmission

Do While lDelay = 0 And lIndex <> NullIndex And mConcurrentHistDataRequests < MaxConcurrentHistDataRequests
    lCurrIndex = lIndex
    lIndex = mRequests(lIndex).Next
    If mRequests(lCurrIndex).EarliestSubmissionTime = 0 Then mRequests(lCurrIndex).EarliestSubmissionTime = gGetEarliestSubmissionTime(mRequests(lCurrIndex).Request)
    If mRequests(lCurrIndex).EarliestSubmissionTime < lTimestamp Then
        
        mTWSAPI.RequestHistoricalData mRequests(lCurrIndex).Request
        mRequests(lCurrIndex).ActualSubmissionTime = gRecordSubmission(mRequests(lCurrIndex).Request)
        
        mConcurrentHistDataRequests = mConcurrentHistDataRequests + 1
        
        unqueueEntry lCurrIndex
        lDelay = gDelayTillNextSubmission
    ElseIf mRequests(lCurrIndex).EarliestSubmissionTime < lEarliest Then
        lEarliest = mRequests(lCurrIndex).EarliestSubmissionTime
    End If
Loop

If lDelay <> 0 Then
    mTaskContext.Suspend lDelay
ElseIf lEarliest <> MaxDate Then
    mTaskContext.Suspend (lEarliest - lTimestamp) * 86400# * 1000#
Else
    mTaskContext.Suspend -1
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub queueEntry(ByVal pIndex As Long, ByVal pPriority As Boolean)
Const ProcName As String = "queueEntry"
On Error GoTo Err

If pPriority Then
    If mFirstPendingRequestsEntry = NullIndex Then
        mFirstPendingRequestsEntry = pIndex
        mLastPendingRequestsEntry = pIndex
        mRequests(pIndex).Next = NullIndex
    ElseIf mLastPriorityPendingRequestsEntry = NullIndex Then
        mRequests(pIndex).Next = mFirstPendingRequestsEntry
        mFirstPendingRequestsEntry = pIndex
    Else
        mRequests(pIndex).Next = mRequests(mLastPriorityPendingRequestsEntry).Next
        mRequests(mLastPriorityPendingRequestsEntry).Next = pIndex
    End If
    mLastPriorityPendingRequestsEntry = pIndex
Else
    If mFirstPendingRequestsEntry = NullIndex Then mFirstPendingRequestsEntry = pIndex
    If mLastPendingRequestsEntry <> NullIndex Then mRequests(mLastPendingRequestsEntry).Next = pIndex
    mLastPendingRequestsEntry = pIndex
    mRequests(mLastPendingRequestsEntry).Next = NullIndex
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function releaseEntry(ByVal pIndex As Long)
Dim lBlankRequest As HistoricalDataRequest
mRequests(pIndex).Next = mFirstFreeRequestsEntry
mFirstFreeRequestsEntry = pIndex
mRequests(pIndex).Cancelling = False
mRequests(pIndex).ActualSubmissionTime = 0
mRequests(pIndex).EarliestSubmissionTime = 0
mRequests(pIndex).Request = lBlankRequest
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
If lIndex = mLastPriorityPendingRequestsEntry Then mLastPriorityPendingRequestsEntry = lPrevIndex

If lPrevIndex <> NullIndex Then mRequests(lPrevIndex).Next = mRequests(lIndex).Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


