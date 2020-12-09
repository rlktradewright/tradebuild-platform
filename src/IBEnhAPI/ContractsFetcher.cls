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

Private mContractsBuilder                           As ContractsBuilder

Private mCachedContracts                            As IContracts
Private mCachedTwsContracts                         As EnumerableCollection

Private mContractsFutureBuilder                     As FutureBuilder
Attribute mContractsFutureBuilder.VB_VarHelpID = -1

Private mTwsContracts                               As EnumerableCollection

Private mContractCache                              As ContractCache

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mTwsRequestId                               As Long

Private mCancelled                                  As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
Set mContractsBuilder = New ContractsBuilder
Set mTwsContracts = New EnumerableCollection
mTwsRequestId = NullIndex
End Sub

'@================================================================================
' IContractDetailsConsumer Interface Members
'@================================================================================

Private Sub IContractDetailsConsumer_EndContracts(ByVal pRequestId As Long)
Const ProcName As String = "IContractDetailsConsumer_EndContracts"
On Error GoTo Err

If mTwsRequestId = NullIndex Then
    ' there must have been a cancellation before the EndContracts
    ' notification was received
    Exit Sub
End If

mRequestManager.NotifyFetchCompleted Me
mContractsFutureBuilder.Value = mContractsBuilder.Contracts
mContractsFutureBuilder.Complete
Set mRequestManager = Nothing
mTwsRequestId = NullIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_EndOptionParameters(ByVal pRequestId As Long)

End Sub

Private Sub IContractDetailsConsumer_NotifyContract(ByVal pRequestId As Long, ByVal pTwsContract As TwsContract)
Const ProcName As String = "IContractDetailsConsumer_NotifyContract"
On Error GoTo Err

processTwsContract pTwsContract

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IContractDetailsConsumer_NotifyError"
On Error GoTo Err

If pErrorCode = 200 Then
    mContractsFutureBuilder.Value = mContractsBuilder.Contracts
    mContractsFutureBuilder.Complete
End If

mTwsRequestId = NullIndex

mRequestManager.NotifyFetchError Me, pErrorCode, pErrorMsg

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

Static sContractsEn As Enumerator
If sContractsEn Is Nothing Then Set sContractsEn = mCachedContracts.Enumerator

Static sTwsContractsEn As Enumerator
If sTwsContractsEn Is Nothing Then Set sTwsContractsEn = mCachedTwsContracts.Enumerator

If sContractsEn.MoveNext Then
    mRequestManager.NotifyContract sContractsEn.Current
ElseIf sTwsContractsEn.MoveNext Then
    mRequestManager.NotifyTwsContract sTwsContractsEn.Current
Else
    mTaskContext.Finish Empty, False
    mRequestManager.NotifyFetchCompleted Me
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
    
Set mCachedContracts = ev.Future.Value
StartTask Me, PriorityNormal

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
                ByVal pContractSpecifier As IContractSpecifier)
Const ProcName As String = "FetchContracts"
On Error GoTo Err

Set mRequestManager = pRequestManager
Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
Set mContractSpecifier = pContractSpecifier

If mContractCache.GetContracts( _
                        mContractSpecifier, _
                        mCachedContracts, _
                        mCachedTwsContracts) Then
    StartTask Me, PriorityNormal
Else
    fetchContractsFromIBApi mContractSpecifier
End If

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

If mContractCache.GetContractsFutureBuilder( _
                                    pContractSpec, _
                                    mContractsFutureBuilder, _
                                    mCachedTwsContracts) Then
    mFutureWaiter.Add mContractsFutureBuilder.Future
Else
    Set mContractsFutureBuilder = New FutureBuilder
    mContractCache.SetContractsFutureBuilder pContractSpec, mContractsFutureBuilder, mTwsContracts
    If gLogger.IsLoggable(LogLevelDetail) Then gLog "Request contract details for", ModuleName, ProcName, pContractSpec.ToString, LogLevelDetail
    mTwsRequestId = mContractRequester.RequestContractDetails(pContractSpec, Me)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processPriceIncrementRuleset( _
                ByVal pContract As IContract)
Const ProcName As String = "processPriceIncrementRuleset"
On Error GoTo Err

Dim lRulesetID As Long
lRulesetID = CLng(pContract.ProviderProperties.GetParameterValue("MarketRuleID", "0"))

Dim lPriceMagnifier As Long
lPriceMagnifier = CLng(pContract.ProviderProperties.GetParameterValue("PriceMagnifier", "1"))

mContractRequester.RequestPriceIncrementRuleset lRulesetID, lPriceMagnifier

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

mTwsContracts.Add pTwsContract
If Not mContractsBuilder.Contracts.Contains(lContract.Specifier) Then
    mContractsBuilder.Add lContract
End If

addCacheEntries lContract, pTwsContract
processPriceIncrementRuleset lContract

mRequestManager.NotifyContract lContract
mRequestManager.NotifyTwsContract pTwsContract

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub





