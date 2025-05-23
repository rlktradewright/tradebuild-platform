VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractsTwsRequester"
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

Private Enum ContractInfoRequestTypes
    ContractInfoRequestTypeNone
    ContractInfoRequestTypeContracts
    ContractInfoRequestTypeOptionParameters
End Enum

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

Private Type ContractRequestTableEntry
    Index                   As Long ' of this entry
    State                   As States
    RequestType             As ContractInfoRequestTypes
    ContractSpec            As IContractSpecifier
    Requester               As IContractDetailsConsumer
    NumberReturned          As Long
    PendingContracts        As EnumerableCollection
    Next                    As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ContractsTwsRequester"

Private Const MaxConcurrentContractDetailsRequests  As Long = 20

'@================================================================================
' Member variables
'@================================================================================

Private mTwsAPI                                     As TwsAPI
Attribute mTwsAPI.VB_VarHelpID = -1

Private mContractCache                              As ContractCache

Private mRequests()                                 As ContractRequestTableEntry
Private mRequestsIndex                              As Long
Private mFirstFreeRequestsEntry                     As Long
Private mFirstRequestEntry                          As Long
Private mLastRequestEntry                           As Long

Private mFirstPendingContractEntry                  As Long
Private mLastPendingContractEntry                   As Long

Private mTaskContext                                As TaskContext

Private mConcurrentRequests                         As Long
Private mMaxConcurrentRequests                      As Long

Private mIsTwsConnectedToIBServers                  As Boolean

Private mRulesetRequests                            As EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mRequests(7) As ContractRequestTableEntry
mFirstFreeRequestsEntry = NullIndex
mFirstRequestEntry = NullIndex
mLastRequestEntry = NullIndex
mFirstPendingContractEntry = NullIndex
mLastPendingContractEntry = NullIndex
Set mRulesetRequests = New EnumerableCollection
End Sub

'@================================================================================
' IContractDetailsConsumer Interface Members
'@================================================================================

Private Sub IContractDetailsConsumer_EndContracts(ByVal pRequestId As Long)
Const ProcName As String = "IContractDetailsConsumer_EndContracts"
On Error GoTo Err

If mRequests(pRequestId).NumberReturned = 1 Then
    mRequests(pRequestId).Requester.EndContracts pRequestId
    releaseEntry pRequestId
Else
    mRequests(pRequestId).State = StateRequestCompleted
End If
CompleteRequest

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_EndOptionParameters(ByVal pRequestId As Long)
Const ProcName As String = "IContractDetailsConsumer_EndOptionParameters"
On Error GoTo Err

mRequests(pRequestId).Requester.EndOptionParameters pRequestId
releaseEntry pRequestId

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyContract( _
                ByVal pRequestId As Long, _
                ByVal pTwsContract As TwsContract)
Const ProcName As String = "IContractDetailsConsumer_NotifyContract"
On Error GoTo Err

If mRequests(pRequestId).State = StateRequestCancelling Then
    ' continue to process returned contracts, so they get added to
    ' the cache
End If

If mRequests(pRequestId).NumberReturned = 0 Then
    mRequests(pRequestId).NumberReturned = 1
    mRequests(pRequestId).Requester.NotifyContract pRequestId, pTwsContract
Else
    mRequests(pRequestId).NumberReturned = mRequests(pRequestId).NumberReturned + 1
    If mRequests(pRequestId).PendingContracts Is Nothing Then
        Set mRequests(pRequestId).PendingContracts = New EnumerableCollection
        addEntryToPendingContractQueue pRequestId
    End If
    mRequests(pRequestId).PendingContracts.Add pTwsContract
    mTaskContext.Continue
End If

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IContractDetailsConsumer_NotifyError"
On Error GoTo Err

If mRequests(pRequestId).State <> StateRequestCancelling Then mRequests(pRequestId).Requester.NotifyError pRequestId, pErrorCode, pErrorMsg
CompleteRequest

If mRequests(pRequestId).NumberReturned > 1 Then
    removeEntryFromRequestQueue pRequestId
End If

releaseEntry pRequestId

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyMarketRule( _
                ByVal pMarketRuleId As Long, _
                ByRef pPriceIncrements() As TwsPriceIncrement)
Const ProcName As String = "IContractDetailsConsumer_NotifyMarketRule"
On Error GoTo Err

GIBEnhApi.Logger.Log "Market rule notified: " & pMarketRuleId, ProcName, ModuleName, LogLevelMediumDetail

Dim lRequestEntry As MarketRuleRequestEntry
lRequestEntry = mRulesetRequests.Item(CStr(pMarketRuleId))

Dim lRuleset As New PriceIncrementRuleset
lRuleset.Initialise pMarketRuleId

Dim i As Long
For i = 0 To UBound(pPriceIncrements)
    Dim lPriceIncrement As TwsPriceIncrement
    lPriceIncrement = pPriceIncrements(i)
    lRuleset.Add lPriceIncrement.LowEdge * lRequestEntry.PriceMagnifier, _
                    lPriceIncrement.Increment * lRequestEntry.PriceMagnifier
Next

Dim lFutureBuilder As FutureBuilder
Set lFutureBuilder = lRequestEntry.RulesetFutureBuilder
lFutureBuilder.Value = lRuleset
lFutureBuilder.Complete

mRulesetRequests.Remove CStr(pMarketRuleId)

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyOptionParameter( _
                ByVal pRequestId As Long, _
                ByVal pExchange As String, _
                ByVal pUnderlyingContractId As Long, _
                ByVal pTradingClass As String, _
                ByVal pMultiplier As Double, _
                pExpirations() As String, _
                pStrikes() As Double)
Const ProcName As String = "IContractDetailsConsumer_NotifyOptionParameter"
On Error GoTo Err

If mRequests(pRequestId).State = StateRequestCancelling Then Exit Sub

mRequests(pRequestId).Requester.NotifyOptionParameter _
                pRequestId, _
                pExchange, _
                pUnderlyingContractId, _
                pTradingClass, _
                pMultiplier, _
                pExpirations, _
                pStrikes

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITwsConnectionStateListener Interface Members
'@================================================================================

Private Sub ITwsConnectionStateListener_NotifyAPIConnectionStateChange(ByVal pSource As Object, ByVal pState As ApiConnectionStates, ByVal pMessage As String)
Const ProcName As String = "ITwsConnectionStateListener_NotifyAPIConnectionStateChange"
On Error GoTo Err

GIBEnhApi.Log "Connection state :" & TwsConnectionStateToString(pState) & "; " & pMessage, _
                ModuleName, ProcName, , LogLevelDetail

Select Case pState
Case ApiConnNotConnected
Case ApiConnConnecting

Case ApiConnConnected
    ' we'll assume that TWS has a server connection: if not,
    ' it will soon tell us
    mIsTwsConnectedToIBServers = True
    schedule
Case ApiConnFailed

End Select

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionClosed(ByVal pSource As Object)
mIsTwsConnectedToIBServers = False
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionRecovered(ByVal pSource As Object, ByVal pDataLost As Boolean)
Const ProcName As String = "ITwsConnectionStateListener_NotifyIBServerConnectionRecovered"
On Error GoTo Err

mIsTwsConnectedToIBServers = True
schedule

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

If mFirstPendingContractEntry = NullIndex Then
    mTaskContext.Suspend -1
    Exit Sub
End If

Dim lIndex As Long
lIndex = mFirstPendingContractEntry

With mRequests(lIndex)
    Dim en As Enumerator
    Set en = .PendingContracts.Enumerator
    If en.MoveNext Then
        Dim lTwsContract As TwsContract
        Set lTwsContract = en.Current
        .Requester.NotifyContract lIndex, lTwsContract
        en.Remove
    ElseIf .State = StateRequestCompleted Then
        .Requester.EndContracts lIndex
        removeEntryFromPendingContractQueue lIndex
        releaseEntry lIndex
    End If
End With
 
'If (mFirstRequestEntry = NullIndex Or mConcurrentRequests = mMaxConcurrentRequests) And _
'    mFirstPendingContractEntry = NullIndex _
'Then mTaskContext.Suspend -1

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
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

Friend Sub CancelRequest(ByVal pRequestId As Long)
Const ProcName As String = "CancelRequest"
On Error GoTo Err

Select Case mRequests(pRequestId).State
Case StateFree

Case StateAllocated
    releaseEntry pRequestId
Case StateRequestQueued
    removeEntryFromRequestQueue pRequestId
    releaseEntry pRequestId
Case StateRequestInProgress
    mRequests(pRequestId).State = StateRequestCancelling
Case StateRequestCancelling

End Select

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTwsAPI As TwsAPI, _
                ByVal pContractCache As ContractCache)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTwsAPI = pTwsAPI
Set mContractCache = pContractCache

mMaxConcurrentRequests = MaxConcurrentContractDetailsRequests

GIBEnhApi.Log "Max concurrent contract details requests", ModuleName, ProcName, CStr(mMaxConcurrentRequests), LogLevelDetail

StartTask Me, PriorityNormal

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function RequestContractDetails( _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pRequester As IContractDetailsConsumer) As Long
Const ProcName As String = "RequestContractDetails"
On Error GoTo Err

Dim lIndex As Long
lIndex = allocateEntry
mRequests(lIndex).RequestType = ContractInfoRequestTypeContracts
Set mRequests(lIndex).ContractSpec = pContractSpecifier
Set mRequests(lIndex).Requester = pRequester

GIBEnhApi.Log "Queueing request for", ModuleName, ProcName, pContractSpecifier.Key, LogLevelMediumDetail
addEntryToRequestQueue lIndex

RequestContractDetails = lIndex

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Function RequestOptionParameters( _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pRequester As IContractDetailsConsumer) As Long
Const ProcName As String = "RequestOptionParameters"
On Error GoTo Err

Dim lIndex As Long
lIndex = allocateEntry
mRequests(lIndex).RequestType = ContractInfoRequestTypeOptionParameters
Set mRequests(lIndex).ContractSpec = pContractSpecifier
Set mRequests(lIndex).Requester = pRequester

addEntryToRequestQueue lIndex

RequestOptionParameters = lIndex

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Function RequestPriceIncrementRuleset( _
                ByVal pRulesetID As Long, _
                ByVal pPriceMagnifier As Long) As IFuture
Const ProcName As String = "RequestPriceIncrementRuleset"
On Error GoTo Err

Dim lRuleset As PriceIncrementRuleset
Set lRuleset = mContractCache.GetPriceIncrementRuleset(pRulesetID)
If Not lRuleset Is Nothing Then
    ' we've already cached the ruleset
    Set RequestPriceIncrementRuleset = CreateFuture(lRuleset)
Else
    Dim lRulesetFutureBuilder As FutureBuilder
    If mContractCache.GetRulesetFutureBuilder(pRulesetID, lRulesetFutureBuilder) Then
        ' the ruleset is currently being fetched from the API
    Else
        Dim lRequestEntry As MarketRuleRequestEntry
        Set lRulesetFutureBuilder = New FutureBuilder
        Set lRequestEntry.RulesetFutureBuilder = lRulesetFutureBuilder
        lRequestEntry.PriceMagnifier = pPriceMagnifier
        lRequestEntry.RulesetID = Abs(pRulesetID)
        mRulesetRequests.Add lRequestEntry, CStr(Abs(pRulesetID))
        mContractCache.SetRulesetFutureBuilder pRulesetID, lRequestEntry.RulesetFutureBuilder
        mTwsAPI.RequestMarketRule Abs(pRulesetID)
    End If
    Set RequestPriceIncrementRuleset = lRulesetFutureBuilder.Future
End If

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addEntryToPendingContractQueue( _
                ByVal pIndex As Long)
Const ProcName As String = "addEntryToPendingContractQueue"
On Error GoTo Err

' NB: don't change the state of the entry
queueEntry pIndex, mFirstPendingContractEntry, mLastPendingContractEntry
mTaskContext.Continue

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addEntryToRequestQueue( _
                ByVal pIndex As Long)
Const ProcName As String = "addEntryToRequestQueue"
On Error GoTo Err

queueEntry pIndex, mFirstRequestEntry, mLastRequestEntry
mRequests(pIndex).State = StateRequestQueued

schedule

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function allocateEntry() As Long
Const ProcName As String = "allocateEntry"
On Error GoTo Err

If mRequestsIndex <= UBound(mRequests) Then
    allocateEntry = mRequestsIndex
    mRequestsIndex = mRequestsIndex + 1
ElseIf mFirstFreeRequestsEntry <> NullIndex Then
    allocateEntry = mFirstFreeRequestsEntry
    mFirstFreeRequestsEntry = mRequests(mFirstFreeRequestsEntry).Next
Else
    ReDim Preserve mRequests(2 * (UBound(mRequests) + 1) - 1) As ContractRequestTableEntry
    allocateEntry = mRequestsIndex
    mRequestsIndex = mRequestsIndex + 1
End If

mRequests(allocateEntry).State = StateAllocated
mRequests(allocateEntry).Index = allocateEntry

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub CompleteRequest()
Const ProcName As String = "completeRequest"
On Error GoTo Err

If mConcurrentRequests > 0 Then mConcurrentRequests = mConcurrentRequests - 1
schedule
mTaskContext.Continue

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub queueEntry( _
                ByVal pIndex As Long, _
                ByRef pFirstIndex As Long, _
                ByRef pLastIndex As Long)
Const ProcName As String = "queueEntry"
On Error GoTo Err

If pFirstIndex = NullIndex Then pFirstIndex = pIndex
If pLastIndex <> NullIndex Then mRequests(pLastIndex).Next = pIndex
pLastIndex = pIndex
mRequests(pLastIndex).Next = NullIndex

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function releaseEntry(ByVal pIndex As Long)
Const ProcName As String = "releaseEntry"
On Error GoTo Err

mRequests(pIndex).Next = mFirstFreeRequestsEntry
mFirstFreeRequestsEntry = pIndex
mRequests(pIndex).State = StateFree
mRequests(pIndex).RequestType = ContractInfoRequestTypeNone
Set mRequests(pIndex).ContractSpec = Nothing
Set mRequests(pIndex).Requester = Nothing
mRequests(pIndex).NumberReturned = 0
Set mRequests(pIndex).PendingContracts = Nothing

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub removeEntryFromPendingContractQueue( _
                ByVal pIndex As Long)
Const ProcName As String = "removeEntryFromPendingContractQueue"
On Error GoTo Err

unqueueEntry pIndex, mFirstPendingContractEntry, mLastPendingContractEntry

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeEntryFromRequestQueue( _
                ByVal pIndex As Long)
Const ProcName As String = "addEntryToRequestQueue"
On Error GoTo Err

unqueueEntry pIndex, mFirstRequestEntry, mLastRequestEntry

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub schedule()
Const ProcName As String = "schedule"
On Error GoTo Err

If mFirstRequestEntry = NullIndex Then Exit Sub

GIBEnhApi.Log "Scheduling contract requests", ModuleName, ProcName, , LogLevelMediumDetail

If mTwsAPI.ConnectionState <> TwsConnConnected Then
    GIBEnhApi.Log "Can't schedule contract requests: not connected to TWS", ModuleName, ProcName, , LogLevelMediumDetail
    Exit Sub
End If

If Not mIsTwsConnectedToIBServers Then
    GIBEnhApi.Log "Can't schedule contract requests: TWS not connected to IB servers", ModuleName, ProcName, , LogLevelMediumDetail
    Exit Sub
End If

Dim lIndex As Long: lIndex = mFirstRequestEntry
Dim lCurrIndex As Long

Do While lIndex <> NullIndex And mConcurrentRequests < mMaxConcurrentRequests
    lCurrIndex = lIndex
    lIndex = mRequests(lIndex).Next
    If mRequests(lCurrIndex).RequestType = ContractInfoRequestTypeContracts Then
        If GIBEnhApi.Logger.IsLoggable(LogLevelDetail) Then GIBEnhApi.Log "Requesting contract details for", ModuleName, ProcName, mRequests(lCurrIndex).ContractSpec.ToString, LogLevelMediumDetail
        mTwsAPI.RequestContractDetails lCurrIndex, _
                                        ContractSpecToTwsContractSpec(mRequests(lCurrIndex).ContractSpec)
    ElseIf mRequests(lCurrIndex).RequestType = ContractInfoRequestTypeOptionParameters Then
        If GIBEnhApi.Logger.IsLoggable(LogLevelDetail) Then GIBEnhApi.Log "Requesting option parameters for", ModuleName, ProcName, mRequests(lCurrIndex).ContractSpec.ToString, LogLevelMediumDetail
        mTwsAPI.RequestOptionParameters lCurrIndex, _
                                        mRequests(lCurrIndex).ContractSpec.Symbol, _
                                        IIf(mRequests(lCurrIndex).ContractSpec.SecType = TwsSecTypeFuture, _
                                            mRequests(lCurrIndex).ContractSpec.Exchange, _
                                            ""), _
                                        mRequests(lCurrIndex).ContractSpec.SecType, _
                                        mRequests(lCurrIndex).ContractSpec.ProviderProperties.GetParameterValue(ProviderPropertyContractID)
    Else
        Assert False, "Invalid request type"
    End If
    mConcurrentRequests = mConcurrentRequests + 1
        
    removeEntryFromRequestQueue lCurrIndex
    mRequests(lCurrIndex).State = StateRequestInProgress
Loop

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub unqueueEntry( _
                ByVal pIndex As Long, _
                ByRef pFirstIndex As Long, _
                ByRef pLastIndex As Long)
Const ProcName As String = "unqueueEntry"
On Error GoTo Err

Assert pFirstIndex <> NullIndex And pLastIndex <> NullIndex, "No entries in queue"

Dim lIndex As Long: lIndex = pFirstIndex
Dim lPrevIndex As Long: lPrevIndex = NullIndex
Do While lIndex <> pIndex
    lPrevIndex = lIndex
    lIndex = mRequests(lIndex).Next
    Assert lIndex <> NullIndex, "Can't find entry to be removed"
Loop

If lIndex = pFirstIndex Then pFirstIndex = mRequests(lIndex).Next
If lIndex = pLastIndex Then pLastIndex = lPrevIndex

If lPrevIndex <> NullIndex Then mRequests(lPrevIndex).Next = mRequests(lIndex).Next

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub




