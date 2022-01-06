VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractsFetcher"
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

Implements IContractDetailsConsumer
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

Private Const ModuleName                            As String = "ContractsFetcher"

'@================================================================================
' Member variables
'@================================================================================

Private mRequestManager                             As ContractsRequestManager

Private mContractSpecifier                          As IContractSpecifier
Private mContractRequester                          As ContractsTwsRequester

Private mTaskContext                                As TaskContext

Private mContractsForCachingFutureBuilder           As FutureBuilder
Private mContractsForCachingBuilder                 As ContractsBuilder

Private mCachedContracts                            As IContracts
Private mResolvedContracts                          As EnumerableCollection
Private mTwsContracts                               As EnumerableCollection

Private mContractCache                              As ContractCache

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mTwsRequestId                               As Long

Private mCancelled                                  As Boolean

Private mFinishedRetrieval                          As Boolean

Private mFromCache                                  As Boolean

Private mNumberOfUnresolvedContracts                As Long

Private mReturnTwsContracts                         As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
mTwsRequestId = NullIndex
End Sub

'@================================================================================
' IContractDetailsConsumer Interface Members
'@================================================================================

Private Sub IContractDetailsConsumer_EndContracts(ByVal pRequestId As Long)
Const ProcName As String = "IContractDetailsConsumer_EndContracts"
On Error GoTo Err

gLogger.Log "End contracts notified", ProcName, ModuleName, LogLevelDetail

If mTwsRequestId = NullIndex Then
    ' there must have been a cancellation before the EndContracts
    ' notification was received
    Exit Sub
End If

If mReturnTwsContracts Then
    mRequestManager.NotifyFetchCompleted Me
Else
    mFinishedRetrieval = True
    mTaskContext.Continue
End If
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_EndOptionParameters(ByVal pRequestId As Long)

End Sub

Private Sub IContractDetailsConsumer_NotifyContract(ByVal pRequestId As Long, ByVal pTwsContract As TwsContract)
Const ProcName As String = "IContractDetailsConsumer_NotifyContract"
On Error GoTo Err

gLogger.Log "Contract notified: " & pTwsContract.Specifier.ToString, ProcName, ModuleName, LogLevelDetail
processTwsContract pTwsContract

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IContractDetailsConsumer_NotifyError"
On Error GoTo Err

gLogger.Log "Error notified: " & pErrorCode & ": " & pErrorMsg, ProcName, ModuleName, LogLevelDetail

mContractsForCachingFutureBuilder.Fail pErrorCode, pErrorMsg, ProcName

mTwsRequestId = NullIndex

mRequestManager.NotifyFetchError Me, pErrorCode, pErrorMsg
mTaskContext.Finish Empty, False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyMarketRule(ByVal pMarketRuleId As Long, pPriceIncrement() As IBAPIV100.TwsPriceIncrement)

End Sub

Private Sub IContractDetailsConsumer_NotifyOptionParameter(ByVal pRequestId As Long, ByVal pExchange As String, ByVal pUnderlyingContractId As Long, ByVal pTradingClass As String, ByVal pMultiplier As Long, pExpirations() As String, pStrikes() As Double)

End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()
Cancel
End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    Cancel
    mTaskContext.Finish Empty, True
    Exit Sub
End If

If mFromCache Then
    Static sContractsEn As Enumerator
    Static sTwsContractsEn As Enumerator
    
    Dim lDone As Boolean

    If Not mReturnTwsContracts Then
        If sContractsEn Is Nothing Then Set sContractsEn = mCachedContracts.Enumerator
        If sContractsEn.MoveNext Then
            mRequestManager.NotifyContract sContractsEn.Current
        Else
            lDone = True
        End If
    Else
        If sTwsContractsEn Is Nothing Then Set sTwsContractsEn = mTwsContracts.Enumerator
        If sTwsContractsEn.MoveNext Then
            mRequestManager.NotifyTwsContract sTwsContractsEn.Current
        Else
            lDone = True
        End If
    End If
    
    If lDone Then
        gLog "Finished fetching cached contract details for", ModuleName, ProcName, mContractSpecifier.ToString, LogLevelDetail
        mTaskContext.Finish Empty, False
        mRequestManager.NotifyFetchCompleted Me
    End If
Else
    Dim lContractsEn As Enumerator
    Set lContractsEn = mResolvedContracts.Enumerator
    
    If lContractsEn.MoveNext Then
        mRequestManager.NotifyContract lContractsEn.Current
        lContractsEn.Remove
    ElseIf mFinishedRetrieval And mNumberOfUnresolvedContracts = 0 Then
        gLog "Finished fetching contract details from IBAPI for", ModuleName, ProcName, mContractSpecifier.ToString, LogLevelDetail
        mTaskContext.Finish Empty, False
        mContractsForCachingFutureBuilder.Value = mContractsForCachingBuilder.Contracts
        mContractsForCachingFutureBuilder.Complete
        mRequestManager.NotifyFetchCompleted Me
    Else
        mTaskContext.Suspend -1
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsFaulted Then
    mRequestManager.NotifyFetchError Me, ev.Future.ErrorNumber, ev.Future.ErrorMessage
    Exit Sub
End If

If ev.Future.IsCancelled Then Exit Sub

If TypeOf ev.Future.Value Is PriceIncrementRuleset Then
    Dim lContract As IContract
    Set lContract = ev.ContinuationData
    gLogger.Log "Resolved contract: " & lContract.Specifier.ToString, ProcName, ModuleName, LogLevelDetail
    mResolvedContracts.Add lContract
    mNumberOfUnresolvedContracts = mNumberOfUnresolvedContracts - 1
    mTaskContext.Continue
Else
    Set mCachedContracts = ev.Future.Value
    mFromCache = True
    gLogger.Log "Await completed for " & mCachedContracts.Count & " cached contracts for: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
    StartTask Me, PriorityNormal
End If

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

Friend Sub Cancel()
Const ProcName As String = "Cancel"
On Error GoTo Err

mCancelled = True
If mTwsRequestId <> NullIndex Then
    mContractRequester.CancelRequest mTwsRequestId
    mTwsRequestId = NullIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FetchContracts( _
                ByVal pRequestManager As ContractsRequestManager, _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pReturnTwsContracts As Boolean, _
                ByVal pCookie As Variant)
Const ProcName As String = "FetchContracts"
On Error GoTo Err

Set mRequestManager = pRequestManager
Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
Set mContractSpecifier = pContractSpecifier
mReturnTwsContracts = pReturnTwsContracts

Dim lCachedContractsFutureBuilder As FutureBuilder

If mContractCache.GetContracts( _
                        mContractSpecifier, _
                        mCachedContracts, _
                        mTwsContracts) Then
    gLog "Got contract details from cache for", ModuleName, ProcName, mContractSpecifier.ToString, LogLevelDetail
    mFromCache = True
    StartTask Me, PriorityNormal
ElseIf mContractCache.GetContractsFutureBuilder( _
                                    mContractSpecifier, _
                                    lCachedContractsFutureBuilder, _
                                    mTwsContracts) Then
    gLog "Awaiting cached contract details for", ModuleName, ProcName, mContractSpecifier.ToString, LogLevelDetail
    mFutureWaiter.Add lCachedContractsFutureBuilder.Future
Else
    gLog "Fetching contract details from IBAPI for", ModuleName, ProcName, mContractSpecifier.ToString, LogLevelDetail
    Set mResolvedContracts = New EnumerableCollection
    fetchContractsFromIBApi mContractSpecifier
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ReplayContracts( _
                ByVal pRequestManager As ContractsRequestManager, _
                ByVal pContracts As IContracts, _
                ByVal pTwsContracts As EnumerableCollection)
Const ProcName As String = "ReplayContracts"
On Error GoTo Err

Set mRequestManager = pRequestManager
Set mCachedContracts = pContracts
Set mContractSpecifier = mCachedContracts.ContractSpecifier
Set mTwsContracts = pTwsContracts

mFromCache = True
StartTask Me, PriorityNormal

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addCacheEntries( _
                ByVal pContract As IContract, _
                ByVal pTwsContract As TwsContract)
Const ProcName As String = "addCacheEntries"
On Error GoTo Err

addCacheEntry pContract.Specifier.Key, pContract, pTwsContract

Dim lSpec As IContractSpecifier
Set lSpec = CreateContractSpecifier(LocalSymbol:=pContract.Specifier.LocalSymbol, _
                                    Exchange:=pContract.Specifier.Exchange)
addCacheEntry lSpec.Key, pContract, pTwsContract

Set lSpec = CreateContractSpecifier(LocalSymbol:=pContract.Specifier.LocalSymbol, _
                                    Exchange:=pContract.Specifier.Exchange, _
                                    SecType:=pContract.Specifier.SecType)
addCacheEntry lSpec.Key, pContract, pTwsContract

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addCacheEntry( _
                ByVal pKey As String, _
                ByVal pContract As IContract, _
                ByVal pTwsContract As TwsContract)
Const ProcName As String = "addCacheEntry"
On Error GoTo Err

If pKey = mContractSpecifier.Key Then
ElseIf mContractCache.ContainsContract(pKey) Then
Else
    mContractCache.AddContract pKey, pContract, pTwsContract
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fetchContractsFromIBApi( _
                ByVal pContractSpec As IContractSpecifier)
Const ProcName As String = "fetchContractsFromIBApi"
On Error GoTo Err

Set mTwsContracts = New EnumerableCollection
Set mContractsForCachingFutureBuilder = New FutureBuilder
Set mContractsForCachingBuilder = New ContractsBuilder
mContractsForCachingBuilder.ContractSpecifier = pContractSpec

mContractCache.SetContractsFutureBuilder pContractSpec, mContractsForCachingFutureBuilder, mTwsContracts
If gLogger.IsLoggable(LogLevelDetail) Then gLog "Request contract details for", ModuleName, ProcName, pContractSpec.ToString, LogLevelDetail
mTwsRequestId = mContractRequester.RequestContractDetails(pContractSpec, Me)

If Not mReturnTwsContracts Then StartTask Me, PriorityNormal

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processPriceIncrementRuleset( _
                ByVal pContract As IContract, _
                ByVal pTwsContract As TwsContract)
Const ProcName As String = "processPriceIncrementRuleset"
On Error GoTo Err

Dim lRulesetID As Long
lRulesetID = CLng(pContract.ProviderProperties.GetParameterValue("MarketRuleID", "0"))

If lRulesetID = 0 Then
    gLogger.Log "Resolved contract (no price increment ruleset): " & pContract.Specifier.ToString, ProcName, ModuleName, LogLevelDetail
    mResolvedContracts.Add pContract
    mTaskContext.Continue
    Exit Sub
End If

Dim lPriceMagnifier As Long
lPriceMagnifier = CLng(pContract.ProviderProperties.GetParameterValue("PriceMagnifier", "1"))

Dim lRulesetFuture As IFuture
Set lRulesetFuture = mContractRequester.RequestPriceIncrementRuleset(lRulesetID, lPriceMagnifier)
If lRulesetFuture.IsAvailable Then
    gLogger.Log "Resolved contract: " & pContract.Specifier.ToString, ProcName, ModuleName, LogLevelDetail
    mResolvedContracts.Add pContract
    mTaskContext.Continue
Else
    gLogger.Log "Unresolved contract: " & pContract.Specifier.ToString, ProcName, ModuleName, LogLevelDetail
    mNumberOfUnresolvedContracts = mNumberOfUnresolvedContracts + 1
    mFutureWaiter.Add lRulesetFuture, CVar(pContract)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTwsContract( _
                ByVal pTwsContract As TwsContract)
Const ProcName As String = "processTwsContract"
On Error GoTo Err

If mCancelled Then Exit Sub

If Not IsValidExchangeCode(pTwsContract.Specifier.Exchange) Then
    gLog "Unknown Exchange code received from Tws", pMsgQualifier:=pTwsContract.Specifier.Exchange, pModName:=ModuleName, pProcName:=ProcName
    Exit Sub
End If
    
Dim lContract As IContract
Set lContract = gTwsContractToContract(pTwsContract)

If Not mContractsForCachingBuilder.Contracts.Contains(lContract.Specifier) Then
    addCacheEntries lContract, pTwsContract
    mTwsContracts.Add pTwsContract
    If mReturnTwsContracts Then mRequestManager.NotifyTwsContract pTwsContract
    mContractsForCachingBuilder.Add lContract
    processPriceIncrementRuleset lContract, pTwsContract
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub





