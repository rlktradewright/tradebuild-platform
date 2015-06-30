VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataRequester"
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

Implements IHistDataConsumer
Implements ITwsConnectionStateListener
Implements IStateChangeListener
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
End Enum

'@================================================================================
' Types
'@================================================================================

Private Type HistoricalDataRequestTableEntry
    State                   As States
    EarliestSubmissionTime  As Date
    ActualSubmissionTime    As Date
    DelayReason             As String
    Request                 As TwsHistoricalDataRequest
    Requester               As HistDataRequestTask
    Next                    As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "HistDataRequester"

'Private Const MaxConcurrentHistDataRequests        As Long = 50 ' limit imposed by IB
Private Const MaxConcurrentHistDataRequests         As Long = 5

'@================================================================================
' Member variables
'@================================================================================

Private mTwsAPI                                     As TwsAPI
Attribute mTwsAPI.VB_VarHelpID = -1

Private mRequests()                                 As HistoricalDataRequestTableEntry
Private mRequestsIndex                              As Long

Private mFirstFreeRequestsEntry                     As Long

Private mFirstPendingRequestsEntry                  As Long
Private mLastPendingRequestsEntry                   As Long

Private mFirstPendingPriorityRequestsEntry          As Long
Private mLastPendingPriorityRequestsEntry           As Long

Private mTaskContext                                As TaskContext

Private mConcurrentHistDataRequests                 As Long

Private mIsRequestPacingDisabled                    As Boolean

Private mPacer                                      As HistDataRequestPacer

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mRequests(7) As HistoricalDataRequestTableEntry
mFirstFreeRequestsEntry = NullIndex
mFirstPendingRequestsEntry = NullIndex
mFirstPendingPriorityRequestsEntry = NullIndex
mLastPendingRequestsEntry = NullIndex
mLastPendingPriorityRequestsEntry = NullIndex
End Sub

'@================================================================================
' IHistDataConsumer Interface Members
'@================================================================================

Private Sub IHistDataConsumer_EndHistData(ByVal pRequestId As Long)
Const ProcName As String = "IHistDataConsumer_EndHistData"
On Error GoTo Err

CompleteRequest pRequestId
If mRequests(pRequestId).State = StateRequestCancelling Then Exit Sub
mRequests(pRequestId).Requester.EndHistData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IHistDataConsumer_NotifyBar(ByVal pRequestId As Long, pBar As TwsBar)
Const ProcName As String = "IHistDataConsumer_NotifyBar"
On Error GoTo Err

If mRequests(pRequestId).State = StateRequestCancelling Then Exit Sub
mRequests(pRequestId).Requester.NotifyBar pBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IHistDataConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IHistDataConsumer_NotifyError"
On Error GoTo Err

CompleteRequest pRequestId
If mRequests(pRequestId).State = StateRequestCancelling Then Exit Sub
mRequests(pRequestId).Requester.NotifyError pErrorCode, pErrorMsg

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IHistDataConsumer_StartHistData(ByVal pRequestId As Long, ByVal pStartDate As String, ByVal pEndDate As String, ByVal pBarCount As Long)
Const ProcName As String = "IHistDataConsumer_StartHistData"
On Error GoTo Err



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

If pState = TwsConnConnected Then doSchedule

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionClosed(ByVal pSource As Object)

End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionRecovered(ByVal pSource As Object, ByVal pDataLost As Boolean)

End Sub

'@================================================================================
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

Dim tli As TimerListItem
Set tli = ev.Source
releaseEntry CLng(tli.data)

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

Friend Property Get IsRequestPacingDisabled() As Boolean
IsRequestPacingDisabled = mIsRequestPacingDisabled
End Property

Friend Property Get TwsAPI() As TwsAPI
Set TwsAPI = mTwsAPI
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub CancelHistoricalDataRequest( _
                ByVal pRequesterId As Long)
Const ProcName As String = "CancelHistoricalDataRequest"
On Error GoTo Err

Select Case mRequests(pRequesterId).State
Case StateFree

Case StateAllocated

Case StateRequestQueued
    unqueueEntry pRequesterId
    Schedule
Case StateRequestInProgress
    mRequests(pRequesterId).State = StateRequestCancelling
    mTwsAPI.CancelHistoricalData pRequesterId
    
    ' it's possible that Tws has already sent the data but we haven't received it yet,
    ' so we keep this entry for 'long enough' to ensure that any pending data
    ' has been received - ie 60 seconds
    GetGlobalTimerList.Add(pRequesterId, 60).AddStateChangeListener Me
Case StateRequestCancelling
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub DisableHistoricalDataRequestPacing()
Const ProcName As String = "DisableHistoricalDataRequestPacing"
On Error GoTo Err

mIsRequestPacingDisabled = True
Set mPacer = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Set mTwsAPI = Nothing
If Not mTaskContext Is Nothing Then mTaskContext.Finish Empty, False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function GetTwsMessage() As Byte()
Const ProcName As String = "GetTwsMessage"
On Error GoTo Err

GetTwsMessage = mTwsAPI.GetRawTwsMessage

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise(ByVal pTwsAPI As TwsAPI)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTwsAPI = pTwsAPI
Set mPacer = New HistDataRequestPacer

StartTask Me, PriorityNormal

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ProcessTwsMessage( _
                ByRef pMessage() As Byte)
Const ProcName As String = "ProcessTwsMessage"
On Error GoTo Err

mTwsAPI.ProcessRawTwsMessage pMessage, False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function RegisterHistoricalDataRequestTask( _
                ByVal pRequester As HistDataRequestTask) As Long
Const ProcName As String = "RegisterHistoricalDataRequestTask"
On Error GoTo Err

Dim lIndex As Long

lIndex = allocateEntry
Set mRequests(lIndex).Requester = pRequester

RegisterHistoricalDataRequestTask = lIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub RequestHistoricalData( _
                ByVal pRequesterId As Long, _
                ByRef pRequest As TwsHistoricalDataRequest, _
                ByVal pPriority As Boolean)
Const ProcName As String = "RequestHistoricalData"
On Error GoTo Err

mRequests(pRequesterId).Request = pRequest

queueEntry pRequesterId, pPriority

Schedule

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Schedule()
Const ProcName As String = "Schedule"
On Error GoTo Err

mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function UnregisterHistoricalDataRequestTask( _
                ByVal pRequesterId As Long) As Long
Const ProcName As String = "UnregisterHistoricalDataRequestTask"
On Error GoTo Err

If mRequests(pRequesterId).State <> StateRequestCancelling Then releaseEntry pRequesterId

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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

mRequests(allocateEntry).EarliestSubmissionTime = MaxDate
mRequests(allocateEntry).State = StateAllocated

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub CompleteRequest(ByVal pRequestId As Long)
Const ProcName As String = "CompleteRequest"
On Error GoTo Err

mRequests(pRequestId).ActualSubmissionTime = 0
mRequests(pRequestId).EarliestSubmissionTime = MaxDate
mRequests(pRequestId).DelayReason = ""
If mConcurrentHistDataRequests > 0 Then mConcurrentHistDataRequests = mConcurrentHistDataRequests - 1
Schedule

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doSchedule()
Const ProcName As String = "doSchedule"
On Error GoTo Err

If mTwsAPI.ConnectionState <> TwsConnConnected Then
    mTaskContext.Suspend -1
    Exit Sub
End If

If mIsRequestPacingDisabled Then
    scheduleUnpaced
Else
    schedulePaced
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function removeEntry( _
                ByVal pIndex As Long, _
                ByRef pFirstQueueIndex As Long, _
                ByRef pLastQueueIndex As Long) As Boolean
Dim lIndex As Long
lIndex = pFirstQueueIndex

Dim lPrevIndex As Long
lPrevIndex = NullIndex

Do While lIndex <> pIndex And lIndex <> NullIndex
    lPrevIndex = lIndex
    lIndex = mRequests(lIndex).Next
Loop

If lIndex <> NullIndex Then
    If lIndex = pFirstQueueIndex Then pFirstQueueIndex = mRequests(lIndex).Next
    If lIndex = pLastQueueIndex Then pLastQueueIndex = lPrevIndex
    
    If lPrevIndex <> NullIndex Then mRequests(lPrevIndex).Next = mRequests(lIndex).Next
End If
removeEntry = (lIndex <> NullIndex)
End Function

Private Function getDelayTillEarliestSubmission(ByRef pReason As String) As Long
Const ProcName As String = "getDelayTillEarliestSubmission"
On Error GoTo Err

getDelayTillEarliestSubmission = getDelayTillEarliestSubmissionInQueue(pReason, mFirstPendingPriorityRequestsEntry)

If getDelayTillEarliestSubmission = 0 Then getDelayTillEarliestSubmission = getDelayTillEarliestSubmissionInQueue(pReason, mFirstPendingRequestsEntry)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getDelayTillEarliestSubmissionInQueue(ByRef pReason As String, ByVal pFirstPendingEntry As Long) As Long
Const ProcName As String = "getDelayTillEarliestSubmissionInQueue"
On Error GoTo Err

Dim lTimestamp As Date
lTimestamp = GetTimestamp

Dim lEarliest As Date
lEarliest = MaxDate

Dim lCurrIndex As Long
lCurrIndex = pFirstPendingEntry

Dim lEarliestReason As String

Do While lCurrIndex <> NullIndex
    setEarliestSubmissionTime lCurrIndex
    If mRequests(lCurrIndex).EarliestSubmissionTime < lEarliest Then
        lEarliest = mRequests(lCurrIndex).EarliestSubmissionTime
        lEarliestReason = mRequests(lCurrIndex).DelayReason
    End If
    lCurrIndex = mRequests(lCurrIndex).Next
Loop

If lEarliest <> MaxDate And lEarliest <> 0# Then
    pReason = lEarliestReason
    getDelayTillEarliestSubmissionInQueue = (lEarliest - lTimestamp) * 86400# * 1000#
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getNextReadyRequest() As Long
Const ProcName As String = "getNextReadyRequest"
On Error GoTo Err

getNextReadyRequest = getNextReadyRequestFromQueue(mFirstPendingPriorityRequestsEntry)
If getNextReadyRequest = NullIndex Then getNextReadyRequest = getNextReadyRequestFromQueue(mFirstPendingRequestsEntry)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getNextReadyRequestFromQueue(ByVal pFirstPendingEntry As Long) As Long
Const ProcName As String = "getNextReadyRequestFromQueue"
On Error GoTo Err

getNextReadyRequestFromQueue = NullIndex

Dim lTimestamp As Date
lTimestamp = GetTimestamp

Dim lCurrIndex As Long
lCurrIndex = pFirstPendingEntry

Do While lCurrIndex <> NullIndex
    setEarliestSubmissionTime lCurrIndex
    If mRequests(lCurrIndex).EarliestSubmissionTime <= lTimestamp Then
        getNextReadyRequestFromQueue = lCurrIndex
        Exit Do
    End If
    lCurrIndex = mRequests(lCurrIndex).Next
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub queueEntry(ByVal pIndex As Long, ByVal pPriority As Boolean)
Const ProcName As String = "queueEntry"
On Error GoTo Err

If pPriority Then
    If mFirstPendingPriorityRequestsEntry = NullIndex Then mFirstPendingPriorityRequestsEntry = pIndex
    If mLastPendingPriorityRequestsEntry <> NullIndex Then mRequests(mLastPendingPriorityRequestsEntry).Next = pIndex
    mLastPendingPriorityRequestsEntry = pIndex
    mRequests(mLastPendingPriorityRequestsEntry).Next = NullIndex
Else
    If mFirstPendingRequestsEntry = NullIndex Then mFirstPendingRequestsEntry = pIndex
    If mLastPendingRequestsEntry <> NullIndex Then mRequests(mLastPendingRequestsEntry).Next = pIndex
    mLastPendingRequestsEntry = pIndex
    mRequests(mLastPendingRequestsEntry).Next = NullIndex
End If

mRequests(pIndex).State = StateRequestQueued

If pPriority Then
    gLog "Queue priority entry " & pIndex & _
                "; FirstPriorityEntry=" & mFirstPendingPriorityRequestsEntry & _
                "; LastPriorityEntry=" & mLastPendingPriorityRequestsEntry, _
                ModuleName, ProcName, , LogLevelDetail
Else
    gLog "Queue entry " & pIndex & _
                "; FirstEntry=" & mFirstPendingRequestsEntry & _
                "; LastEntry=" & mLastPendingRequestsEntry, _
                ModuleName, ProcName, , LogLevelDetail
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function releaseEntry(ByVal pIndex As Long)
Const ProcName As String = "releaseEntry"
On Error GoTo Err

Dim lBlankRequest As TwsHistoricalDataRequest

mRequests(pIndex).State = StateFree
mRequests(pIndex).Next = mFirstFreeRequestsEntry
mFirstFreeRequestsEntry = pIndex
mRequests(pIndex).State = StateFree
mRequests(pIndex).ActualSubmissionTime = 0
mRequests(pIndex).EarliestSubmissionTime = 0
mRequests(pIndex).Request = lBlankRequest
mRequests(pIndex).DelayReason = ""
Set mRequests(pIndex).Requester = Nothing

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub schedulePaced()
Const ProcName As String = "schedulePaced"
On Error GoTo Err

Dim lReason As String
Dim lDelay As Long
lDelay = mPacer.GetDelayTillNextSubmission(lReason)

If lDelay = 0 Then
    Dim lIndex As Long
    lIndex = getNextReadyRequest
    If lIndex <> NullIndex Then
        submitEntry lIndex
        Exit Sub
    End If

    lDelay = mPacer.GetDelayTillNextSubmission(lReason)
    If lDelay = 0 Then lDelay = getDelayTillEarliestSubmission(lReason)
End If
    
If lDelay <> 0 Then
    gLog "Delay till next request (millisecs): " & lDelay & ": " & lReason, ModuleName, ProcName
    mTaskContext.Suspend lDelay
Else
    mTaskContext.Suspend -1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub scheduleUnpaced()
Const ProcName As String = "scheduleUnpaced"
On Error GoTo Err

Dim lIndex As Long
Dim lCurrIndex As Long

lIndex = mFirstPendingRequestsEntry

Do While lIndex <> NullIndex And mConcurrentHistDataRequests < MaxConcurrentHistDataRequests
    lCurrIndex = lIndex
    lIndex = mRequests(lIndex).Next
    mTwsAPI.RequestHistoricalData lCurrIndex, mRequests(lCurrIndex).Request
    mRequests(lCurrIndex).ActualSubmissionTime = GetTimestamp
        
    mConcurrentHistDataRequests = mConcurrentHistDataRequests + 1
        
    unqueueEntry lCurrIndex
Loop

mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setEarliestSubmissionTime(ByVal pIndex As Long)
Dim lReason As String
Const ProcName As String = "setEarliestSubmissionTime"
On Error GoTo Err

If mRequests(pIndex).EarliestSubmissionTime = MaxDate Then
    mRequests(pIndex).EarliestSubmissionTime = mPacer.GetEarliestSubmissionTime(mRequests(pIndex).Request, lReason)
    mRequests(pIndex).DelayReason = lReason
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub submitEntry(ByVal pIndex As Long)
Const ProcName As String = "submitEntry"
On Error GoTo Err

mTwsAPI.RequestHistoricalData pIndex, mRequests(pIndex).Request
mRequests(pIndex).ActualSubmissionTime = mPacer.RecordSubmission(mRequests(pIndex).Request)
mRequests(pIndex).State = StateRequestInProgress
unqueueEntry pIndex

mConcurrentHistDataRequests = mConcurrentHistDataRequests + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub unqueueEntry(ByVal pIndex As Long)
Const ProcName As String = "unqueueEntry"
On Error GoTo Err

Assert (mFirstPendingRequestsEntry <> NullIndex And mLastPendingRequestsEntry <> NullIndex) Or _
    (mFirstPendingPriorityRequestsEntry <> NullIndex And mLastPendingPriorityRequestsEntry <> NullIndex), "No pending entries"

If removeEntry(pIndex, mFirstPendingPriorityRequestsEntry, mLastPendingPriorityRequestsEntry) Then
    gLogger.Log "Unqueue priority entry " & pIndex & _
                "; FirstPriorityEntry=" & mFirstPendingPriorityRequestsEntry & _
                "; LastPriorityEntry=" & mLastPendingPriorityRequestsEntry, _
                ProcName, ModuleName, LogLevelMediumDetail
ElseIf removeEntry(pIndex, mFirstPendingRequestsEntry, mLastPendingRequestsEntry) Then
    gLogger.Log "Unqueue entry " & pIndex & _
                "; FirstEntry=" & mFirstPendingRequestsEntry & _
                "; LastEntry=" & mLastPendingRequestsEntry, _
                ProcName, ModuleName, LogLevelMediumDetail
Else
    Assert False, "Entry not found"
End If
            
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


