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

Private Enum HistoricalDataRequestTypes
    HistoricalDataRequestTypeHistoricalBars = 1
    HistoricalDataRequestTypeRealtimeBars
End Enum

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
    RequestType             As HistoricalDataRequestTypes
    State                   As States
    RetryCount              As Long
    EarliestSubmissionTime  As Date
    DelayReason             As String
    Request                 As TwsHistoricalDataRequest
    Requester               As Object
    RequesterID             As Long ' only applies to RealtimeBars requests
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

Private mHistDataRequests()                         As HistoricalDataRequestTableEntry
Private mHistDataRequestsIndex                      As Long

Private mFirstFreeHistDataRequestsEntry             As Long

Private mFirstPendingHistDataRequestsEntry          As Long
Private mLastPendingHistDataRequestsEntry           As Long

Private mFirstPendingPriorityHistDataRequestsEntry  As Long
Private mLastPendingPriorityHistDataRequestsEntry   As Long

Private mTaskContext                                As TaskContext

Private mConcurrentHistDataRequests                 As Long

Private mIsRequestPacingDisabled                    As Boolean

Private mPacer                                      As HistDataRequestPacer

Private mIsTwsConnectedToIBServers                  As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mHistDataRequests(7) As HistoricalDataRequestTableEntry
mFirstFreeHistDataRequestsEntry = NullIndex
mFirstPendingHistDataRequestsEntry = NullIndex
mFirstPendingPriorityHistDataRequestsEntry = NullIndex
mLastPendingHistDataRequestsEntry = NullIndex
mLastPendingPriorityHistDataRequestsEntry = NullIndex
End Sub

'@================================================================================
' IHistDataConsumer Interface Members
'@================================================================================

Private Sub IHistDataConsumer_EndHistData(ByVal pRequestId As Long)
Const ProcName As String = "IHistDataConsumer_EndHistData"
On Error GoTo Err

Dim isCancelling As Boolean
isCancelling = (mHistDataRequests(pRequestId).State = StateRequestCancelling)

CompleteRequest pRequestId
If isCancelling Then Exit Sub

Dim lHistDataRequestTask As HistDataRequestTask
Set lHistDataRequestTask = mHistDataRequests(pRequestId).Requester
lHistDataRequestTask.EndHistData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IHistDataConsumer_NotifyBar(ByVal pRequestId As Long, pBar As TwsBar)
Const ProcName As String = "IHistDataConsumer_NotifyBar"
On Error GoTo Err

If mHistDataRequests(pRequestId).State = StateRequestCancelling Then Exit Sub

Dim lHistDataRequestTask As HistDataRequestTask
Set lHistDataRequestTask = mHistDataRequests(pRequestId).Requester
lHistDataRequestTask.NotifyBar pBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IHistDataConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IHistDataConsumer_NotifyError"
On Error GoTo Err

Const PacingRetryLimit As Long = 10

If pErrorCode <> 162 Then
    ' drop through if not a Historical Market Data Service error message
ElseIf InStr(1, pErrorMsg, "pacing") = 0 Then
    ' drop through if not a pacing error
ElseIf mIsRequestPacingDisabled Then
    ' if request pacing is disabled, let the application handle it
ElseIf mHistDataRequests(pRequestId).RetryCount >= PacingRetryLimit Then
    ' we're retried several times already, now let the application handle it
    gLog "Pacing retry limit exceeded", ModuleName, ProcName
Else
    ' pacing violation: this should only happen if another API program
    ' using the same TWS instance has done requests that, combined with ours,
    ' infringe the pacing rules. As we are doing automatic request pacing,
    ' just requeue it in the hope that it will all sort itself out eventually,
    ' but don't retry more than PacingRetryLimit times
    mHistDataRequests(pRequestId).RetryCount = mHistDataRequests(pRequestId).RetryCount + 1
    gLog "Requeue request according to pacing rules: attempt " & mHistDataRequests(pRequestId).RetryCount, ModuleName, ProcName
    queueEntry pRequestId, False
    If mConcurrentHistDataRequests > 0 Then mConcurrentHistDataRequests = mConcurrentHistDataRequests - 1
    doSchedule
    Exit Sub
End If

CompleteRequest pRequestId
If mHistDataRequests(pRequestId).RequestType = HistoricalDataRequestTypeHistoricalBars Then
    Dim lHistDataRequestTask As HistDataRequestTask
    Set lHistDataRequestTask = mHistDataRequests(pRequestId).Requester
    lHistDataRequestTask.NotifyError pErrorCode, pErrorMsg
Else
    Dim lRealtimeBarRequester As RealtimeBarRequester
    Set lRealtimeBarRequester = mHistDataRequests(pRequestId).Requester
    If mHistDataRequests(pRequestId).State = StateRequestCancelling Then Exit Sub
    lRealtimeBarRequester.NotifyError mHistDataRequests(pRequestId).RequesterID, pErrorCode, pErrorMsg
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IHistDataConsumer_NotifyRealtimeBar(ByVal pRequestId As Long, pBar As TwsRealtimeBar)
Const ProcName As String = "IHistDataConsumer_NotifyRealtimeBar"
On Error GoTo Err

If mHistDataRequests(pRequestId).State = StateRequestCancelling Then Exit Sub

Dim lRealtimeBarRequester As RealtimeBarRequester
Set lRealtimeBarRequester = mHistDataRequests(pRequestId).Requester
lRealtimeBarRequester.NotifyRealtimeBar mHistDataRequests(pRequestId).RequesterID, pBar

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

If pState = TwsConnConnected Then
    ' we'll assume that TWS has a server connection: if not,
    ' it will soon tell us
    mIsTwsConnectedToIBServers = True
    
    doSchedule
End If

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

Select Case mHistDataRequests(pRequesterId).State
Case StateFree

Case StateAllocated

Case StateRequestQueued
    unqueueEntry pRequesterId
    Schedule
Case StateRequestInProgress
    mHistDataRequests(pRequesterId).State = StateRequestCancelling
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

Friend Sub CancelRealtimeBars( _
                ByVal pRequesterId As Long)
Const ProcName As String = "CancelRealtimeBars"
On Error GoTo Err

Dim i As Long
For i = 0 To mHistDataRequestsIndex
    With mHistDataRequests(i)
        If .Requester Is Nothing Then
        ElseIf .RequesterID = pRequesterId Then
            .State = StateRequestCancelling
            mTwsAPI.CancelRealtimeBars i
            
            ' it's possible that Tws has already sent the next realtime bar
            ' but we haven't received it yet, so we keep this entry for 'long
            ' enough' to ensure that any pending data has been received -
            ' ie 10 seconds
            GetGlobalTimerList.Add(pRequesterId, 10).AddStateChangeListener Me
            
            Exit For
        End If
    End With
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub DisableHistoricalDataRequestPacing()
Const ProcName As String = "DisableHistoricalDataRequestPacing"
On Error GoTo Err

AssertArgument Not mIsRequestPacingDisabled, "Hisotrical data request pacing is already disabled"

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

Friend Function RegisterHistoricalDataRequest( _
                ByVal pRequester As Object) As Long
Const ProcName As String = "RegisterHistoricalDataRequest"
On Error GoTo Err

Dim lIndex As Long

lIndex = allocateEntry
Set mHistDataRequests(lIndex).Requester = pRequester

RegisterHistoricalDataRequest = lIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub RequestHistoricalData( _
                ByVal pRegistrationId As Long, _
                ByRef pRequest As TwsHistoricalDataRequest, _
                ByVal pPriority As Boolean)
Const ProcName As String = "RequestHistoricalData"
On Error GoTo Err

Assert mHistDataRequests(pRegistrationId).State = StateAllocated, "A request is already outstanding "

mHistDataRequests(pRegistrationId).RequestType = HistoricalDataRequestTypeHistoricalBars
mHistDataRequests(pRegistrationId).Request = pRequest

queueEntry pRegistrationId, pPriority

Schedule

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RequestRealtimeBars( _
                ByVal pRequesterId As Long, _
                ByVal pContract As TwsContract, _
                ByVal pRequester As RealtimeBarRequester)
Const ProcName As String = "RequestRealtimeBars"
On Error GoTo Err

Dim lRegistrationId As Long
lRegistrationId = RegisterHistoricalDataRequest(pRequester)

mHistDataRequests(lRegistrationId).RequesterID = pRequesterId
mHistDataRequests(lRegistrationId).RequestType = HistoricalDataRequestTypes.HistoricalDataRequestTypeRealtimeBars
With mHistDataRequests(lRegistrationId).Request
    Set .Contract = pContract
    .WhatToShow = "TRADES"
End With

queueEntry lRegistrationId, False

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

Friend Function UnregisterHistoricalDataRequest( _
                ByVal pRequesterId As Long) As Long
Const ProcName As String = "UnregisterHistoricalDataRequest"
On Error GoTo Err

If mHistDataRequests(pRequesterId).State <> StateRequestCancelling Then releaseEntry pRequesterId

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

If mFirstFreeHistDataRequestsEntry <> NullIndex Then
    allocateEntry = mFirstFreeHistDataRequestsEntry
    mFirstFreeHistDataRequestsEntry = mHistDataRequests(mFirstFreeHistDataRequestsEntry).Next
ElseIf mHistDataRequestsIndex < UBound(mHistDataRequests) Then
    mHistDataRequestsIndex = mHistDataRequestsIndex + 1
    allocateEntry = mHistDataRequestsIndex
Else
    ReDim Preserve mHistDataRequests(2 * (UBound(mHistDataRequests) + 1) - 1) As HistoricalDataRequestTableEntry
    mHistDataRequestsIndex = mHistDataRequestsIndex + 1
    allocateEntry = mHistDataRequestsIndex
End If

Assert allocateEntry <= mTwsAPI.MaxHistoricalDataRequestId, "Maximum number of outstanding historical data requests exceeded"

mHistDataRequests(allocateEntry).EarliestSubmissionTime = MaxDate
mHistDataRequests(allocateEntry).State = StateAllocated

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub CompleteRequest(ByVal pRequestId As Long)
Const ProcName As String = "CompleteRequest"
On Error GoTo Err

mHistDataRequests(pRequestId).State = StateAllocated
mHistDataRequests(pRequestId).EarliestSubmissionTime = MaxDate
mHistDataRequests(pRequestId).DelayReason = ""
If mConcurrentHistDataRequests > 0 Then mConcurrentHistDataRequests = mConcurrentHistDataRequests - 1
Schedule

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doSchedule()
Const ProcName As String = "doSchedule"
On Error GoTo Err

If Not (mTwsAPI.ConnectionState = TwsConnConnected And mIsTwsConnectedToIBServers) Then
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
    lIndex = mHistDataRequests(lIndex).Next
Loop

If lIndex <> NullIndex Then
    If lIndex = pFirstQueueIndex Then pFirstQueueIndex = mHistDataRequests(lIndex).Next
    If lIndex = pLastQueueIndex Then pLastQueueIndex = lPrevIndex
    
    If lPrevIndex <> NullIndex Then mHistDataRequests(lPrevIndex).Next = mHistDataRequests(lIndex).Next
End If
removeEntry = (lIndex <> NullIndex)
End Function

Private Function getDelayTillEarliestSubmission(ByRef pReason As String) As Long
Const ProcName As String = "getDelayTillEarliestSubmission"
On Error GoTo Err

getDelayTillEarliestSubmission = getDelayTillEarliestSubmissionInQueue(pReason, mFirstPendingPriorityHistDataRequestsEntry)

If getDelayTillEarliestSubmission = 0 Then getDelayTillEarliestSubmission = getDelayTillEarliestSubmissionInQueue(pReason, mFirstPendingHistDataRequestsEntry)

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
    If mHistDataRequests(lCurrIndex).EarliestSubmissionTime < lEarliest Then
        lEarliest = mHistDataRequests(lCurrIndex).EarliestSubmissionTime
        lEarliestReason = mHistDataRequests(lCurrIndex).DelayReason
    End If
    lCurrIndex = mHistDataRequests(lCurrIndex).Next
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

getNextReadyRequest = getNextReadyRequestFromQueue(mFirstPendingPriorityHistDataRequestsEntry)
If getNextReadyRequest = NullIndex Then getNextReadyRequest = getNextReadyRequestFromQueue(mFirstPendingHistDataRequestsEntry)

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
    If mHistDataRequests(lCurrIndex).EarliestSubmissionTime <= lTimestamp Then
        getNextReadyRequestFromQueue = lCurrIndex
        Exit Do
    End If
    lCurrIndex = mHistDataRequests(lCurrIndex).Next
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub queueEntry(ByVal pIndex As Long, ByVal pPriority As Boolean)
Const ProcName As String = "queueEntry"
On Error GoTo Err

If pPriority Then
    If mFirstPendingPriorityHistDataRequestsEntry = NullIndex Then mFirstPendingPriorityHistDataRequestsEntry = pIndex
    If mLastPendingPriorityHistDataRequestsEntry <> NullIndex Then mHistDataRequests(mLastPendingPriorityHistDataRequestsEntry).Next = pIndex
    mLastPendingPriorityHistDataRequestsEntry = pIndex
    mHistDataRequests(mLastPendingPriorityHistDataRequestsEntry).Next = NullIndex
Else
    If mFirstPendingHistDataRequestsEntry = NullIndex Then mFirstPendingHistDataRequestsEntry = pIndex
    If mLastPendingHistDataRequestsEntry <> NullIndex Then mHistDataRequests(mLastPendingHistDataRequestsEntry).Next = pIndex
    mLastPendingHistDataRequestsEntry = pIndex
    mHistDataRequests(mLastPendingHistDataRequestsEntry).Next = NullIndex
End If

mHistDataRequests(pIndex).State = StateRequestQueued

If pPriority Then
    gLog "Queue priority entry " & pIndex & _
                "; FirstPriorityEntry=" & mFirstPendingPriorityHistDataRequestsEntry & _
                "; LastPriorityEntry=" & mLastPendingPriorityHistDataRequestsEntry, _
                ModuleName, ProcName, , LogLevelDetail
Else
    gLog "Queue entry " & pIndex & _
                "; FirstEntry=" & mFirstPendingHistDataRequestsEntry & _
                "; LastEntry=" & mLastPendingHistDataRequestsEntry, _
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

With mHistDataRequests(pIndex)
    .State = StateFree
    .Next = mFirstFreeHistDataRequestsEntry
    .State = StateFree
    .EarliestSubmissionTime = MaxDate
    .Request = lBlankRequest
    .DelayReason = ""
    Set .Requester = Nothing
    .RequesterID = 0
    .RequestType = 0
End With

mFirstFreeHistDataRequestsEntry = pIndex

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

lIndex = mFirstPendingPriorityHistDataRequestsEntry
Do While lIndex <> NullIndex And _
        mConcurrentHistDataRequests < MaxConcurrentHistDataRequests
    submitEntry lIndex
    mConcurrentHistDataRequests = mConcurrentHistDataRequests + 1
    lIndex = mHistDataRequests(lIndex).Next
Loop

lIndex = mFirstPendingHistDataRequestsEntry
Do While lIndex <> NullIndex And _
        mConcurrentHistDataRequests < MaxConcurrentHistDataRequests
    submitEntry lIndex
    mConcurrentHistDataRequests = mConcurrentHistDataRequests + 1
    lIndex = mHistDataRequests(lIndex).Next
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

If mHistDataRequests(pIndex).EarliestSubmissionTime = MaxDate Then
    mHistDataRequests(pIndex).EarliestSubmissionTime = mPacer.GetEarliestSubmissionTime(mHistDataRequests(pIndex).Request, lReason)
    mHistDataRequests(pIndex).DelayReason = lReason
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub submitEntry(ByVal pIndex As Long)
Const ProcName As String = "submitEntry"
On Error GoTo Err

If mHistDataRequests(pIndex).RequestType = HistoricalDataRequestTypes.HistoricalDataRequestTypeHistoricalBars Then
    mTwsAPI.RequestHistoricalData pIndex, mHistDataRequests(pIndex).Request
    mHistDataRequests(pIndex).State = StateRequestInProgress
Else
    mTwsAPI.RequestRealtimeBars pIndex, mHistDataRequests(pIndex).Request.Contract, 5, "TRADES", False
End If
unqueueEntry pIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub unqueueEntry(ByVal pIndex As Long)
Const ProcName As String = "unqueueEntry"
On Error GoTo Err

Assert (mFirstPendingHistDataRequestsEntry <> NullIndex And mLastPendingHistDataRequestsEntry <> NullIndex) Or _
    (mFirstPendingPriorityHistDataRequestsEntry <> NullIndex And mLastPendingPriorityHistDataRequestsEntry <> NullIndex), "No pending entries"

If removeEntry(pIndex, mFirstPendingPriorityHistDataRequestsEntry, mLastPendingPriorityHistDataRequestsEntry) Then
    gLogger.Log "Unqueue priority entry " & pIndex & _
                "; FirstPriorityEntry=" & mFirstPendingPriorityHistDataRequestsEntry & _
                "; LastPriorityEntry=" & mLastPendingPriorityHistDataRequestsEntry, _
                ProcName, ModuleName, LogLevelMediumDetail
ElseIf removeEntry(pIndex, mFirstPendingHistDataRequestsEntry, mLastPendingHistDataRequestsEntry) Then
    gLogger.Log "Unqueue entry " & pIndex & _
                "; FirstEntry=" & mFirstPendingHistDataRequestsEntry & _
                "; LastEntry=" & mLastPendingHistDataRequestsEntry, _
                ProcName, ModuleName, LogLevelMediumDetail
Else
    Assert False, "Entry not found"
End If
            
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


