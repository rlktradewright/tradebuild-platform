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

Private Const ModuleName                            As String = "HistDataRequester"

'Private Const MaxConcurrentHistDataRequests        As Long = 50 ' limit imposed by IB
Private Const MaxConcurrentHistDataRequests         As Long = 5

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTwsAPI                          As TwsAPI
Attribute mTwsAPI.VB_VarHelpID = -1

Private mRequests()                                 As HistoricalDataRequestTableEntry
Private mRequestsIndex                              As Long
Private mFirstFreeRequestsEntry                     As Long
Private mFirstPendingRequestsEntry                  As Long
Private mLastPendingRequestsEntry                   As Long
Private mLastPriorityPendingRequestsEntry           As Long

Private mTaskContext                                As TaskContext

Private mConcurrentHistDataRequests                 As Long

Private mIsRequestPacingDisabled                    As Boolean

Private mPacer                                      As HistDataRequestPacer

Private mCurrRequestTask                            As HistDataRequestTask

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mRequests(7) As HistoricalDataRequestTableEntry
mFirstFreeRequestsEntry = NullIndex
mFirstPendingRequestsEntry = NullIndex
mLastPendingRequestsEntry = NullIndex
mLastPriorityPendingRequestsEntry = NullIndex
End Sub

'@================================================================================
' IHistDataConsumer Interface Members
'@================================================================================

Private Sub IHistDataConsumer_EndHistData()
Const ProcName As String = "IHistDataConsumer_EndHistData"
On Error GoTo Err

mCurrRequestTask.EndHistData
Set mCurrRequestTask = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IHistDataConsumer_NotifyBar(pBar As TWSBar)
Const ProcName As String = "IHistDataConsumer_NotifyBar"
On Error GoTo Err

mCurrRequestTask.NotifyBar pBar

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IHistDataConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IHistDataConsumer_NotifyError"
On Error GoTo Err

Set mCurrRequestTask = Nothing
mRequests(pRequestId).Requester.NotifyError pErrorCode, pErrorMsg

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IHistDataConsumer_StartHistData(pStartInfo As TWSHistDataStartInfo)
Const ProcName As String = "IHistDataConsumer_StartHistData"
On Error GoTo Err

Set mCurrRequestTask = mRequests(pStartInfo.RequestId).Requester

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

doSchedule

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

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
                ByVal pRequestId As Long)
Const ProcName As String = "CancelHistoricalDataRequest"
On Error GoTo Err

If mRequests(pRequestId).Cancelling Then Exit Sub
mRequests(pRequestId).Cancelling = True

If mRequests(pRequestId).ActualSubmissionTime = 0 Then
    unqueueEntry pRequestId
    releaseEntry pRequestId
Else
    mTwsAPI.CancelHistoricalData pRequestId
    
    ' it's possible that Tws has already sent the data but we haven't received it yet,
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

Public Sub DisableHistoricalDataRequestPacing()
Const ProcName As String = "DisableHistoricalDataRequestPacing"
On Error GoTo Err

mIsRequestPacingDisabled = True
Set mPacer = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Set mTwsAPI = Nothing
If Not mTaskContext Is Nothing Then mTaskContext.Finish Empty, False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise(ByVal pTwsAPI As TwsAPI)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTwsAPI = pTwsAPI
Set mPacer = New HistDataRequestPacer

StartTask Me, PriorityNormal

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Function RequestHistoricalData( _
                ByRef pRequest As HistoricalDataRequest, _
                ByVal pPriority As Boolean, _
                ByVal pRequester As HistDataRequestTask) As Long
Const ProcName As String = "RequestHistoricalData"
On Error GoTo Err

Dim lIndex As Long

lIndex = allocateEntry
mRequests(lIndex).Request = pRequest
Set mRequests(lIndex).Requester = pRequester

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

If mTwsAPI.ConnectionState <> ConnConnected Then
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
        If mLastPendingRequestsEntry = mLastPriorityPendingRequestsEntry Then mLastPendingRequestsEntry = pIndex
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

gLogger.Log "Queue entry " & pIndex & _
            "; FirstEntry=" & mFirstPendingRequestsEntry & _
            "; LastEntry=" & mLastPendingRequestsEntry & _
            "; LastPriorityEntry=" & mLastPriorityPendingRequestsEntry, _
            ProcName, ModuleName, LogLevelDetail
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function releaseEntry(ByVal pIndex As Long)
Const ProcName As String = "releaseEntry"
On Error GoTo Err

Dim lBlankRequest As HistoricalDataRequest

mRequests(pIndex).Next = mFirstFreeRequestsEntry
mFirstFreeRequestsEntry = pIndex
mRequests(pIndex).Cancelling = False
mRequests(pIndex).ActualSubmissionTime = 0
mRequests(pIndex).EarliestSubmissionTime = 0
mRequests(pIndex).Request = lBlankRequest
Set mRequests(pIndex).Requester = Nothing

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub schedulePaced()
Const ProcName As String = "schedulePaced"
On Error GoTo Err

Dim lIndex As Long
Dim lCurrIndex As Long
Dim lDelay As Long
Dim lTimestamp As Date
Dim lEarliest As Date

lTimestamp = GetTimestamp
lEarliest = MaxDate
lIndex = mFirstPendingRequestsEntry
lDelay = mPacer.DelayTillNextSubmission

Do While lDelay = 0 And lIndex <> NullIndex And mConcurrentHistDataRequests < MaxConcurrentHistDataRequests
    lCurrIndex = lIndex
    lIndex = mRequests(lIndex).Next
    If mRequests(lCurrIndex).EarliestSubmissionTime = 0 Then mRequests(lCurrIndex).EarliestSubmissionTime = mPacer.GetEarliestSubmissionTime(mRequests(lCurrIndex).Request)
    If mRequests(lCurrIndex).EarliestSubmissionTime < lTimestamp Then
        
        mTwsAPI.RequestHistoricalData lCurrIndex, mRequests(lCurrIndex).Request
        mRequests(lCurrIndex).ActualSubmissionTime = mPacer.RecordSubmission(mRequests(lCurrIndex).Request)
        
        mConcurrentHistDataRequests = mConcurrentHistDataRequests + 1
        
        unqueueEntry lCurrIndex
        lDelay = mPacer.DelayTillNextSubmission
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

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

gLogger.Log "Unqueue entry " & pIndex & _
            "; FirstEntry=" & mFirstPendingRequestsEntry & _
            "; LastEntry=" & mLastPendingRequestsEntry & _
            "; LastPriorityEntry=" & mLastPriorityPendingRequestsEntry, _
            ProcName, ModuleName, LogLevelNormal
            
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


