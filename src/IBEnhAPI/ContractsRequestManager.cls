VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractsRequestManager"
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

Private Const ModuleName                            As String = "ContractsRequestManager"

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpecifier                          As IContractSpecifier
Private mContractRequester                          As ContractsTwsRequester

Private mNumOutstandingRequests                     As Long

Private mFetchers                                   As EnumerableCollection

Private WithEvents mResultsContractsFutureBuilder   As FutureBuilder
Attribute mResultsContractsFutureBuilder.VB_VarHelpID = -1
Private mResultsContractsBuilder                    As New ContractsBuilder

Private mListener                                   As IContractFetchListener

Private mCookie                                     As Variant

Private mReturnTwsContracts                         As Boolean

Private mTwsContracts                               As EnumerableCollection

Private mContractCache                              As ContractCache

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
Set mTwsContracts = New EnumerableCollection
Set mResultsContractsFutureBuilder = New FutureBuilder
Set mFetchers = New EnumerableCollection
End Sub

Private Sub Class_Terminate()
Const ProcName As String = "Class_Terminate"
gLogger.Log "ContractsRequestManager terminated for: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelMediumDetail
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then

ElseIf ev.Future.IsFaulted Then

Else
    gLogger.Log "Contracts future available for: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelMediumDetail
    If Not mListener Is Nothing Then mListener.FetchCompleted mCookie
End If

Set mSelfRef = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mResultsContractsFutureBuilder Event Handlers
'@================================================================================

Private Sub mResultsContractsFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mResultsContractsFutureBuilder_Cancelled"
On Error GoTo Err

doCancel

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

Friend Function Fetch( _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pListener As IContractFetchListener, _
                ByVal pCookie As Variant, _
                ByVal pReturnTwsContracts As Boolean) As IFuture
Const ProcName As String = "Fetch"
On Error GoTo Err

Set mSelfRef = Me

Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
Set mContractSpecifier = pContractSpecifier
Set mListener = pListener
mResultsContractsFutureBuilder.Cookie = pCookie

mResultsContractsBuilder.ContractSpecifier = pContractSpecifier

gSetVariant mCookie, pCookie

mReturnTwsContracts = pReturnTwsContracts

gLogger.Log "Fetch " & IIf(mReturnTwsContracts, "TWS ", "") & "contracts for: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelMediumDetail

Dim lCachedContracts As IContracts
Dim lCachedTwsContracts As EnumerableCollection

If Not mContractCache.GetContracts( _
                mContractSpecifier, _
                lCachedContracts, _
                lCachedTwsContracts) Then
    ' the required contracts are not already in the cache
    Select Case mContractSpecifier.SecType
    Case SecTypeNone
        requestContractsForUnspecifiedSectype
    Case SecTypeCash
        requestContractsForCash
    Case Else
        requestContracts mContractSpecifier
    End Select
    mFutureWaiter.Add mResultsContractsFutureBuilder.Future
    Set Fetch = mResultsContractsFutureBuilder.Future

ElseIf lCachedContracts.Count = 1 Then
    ' there is only one contract for this specifier, already in the cache. So notify
    ' it to the listener and return the relevant result
    gLogger.Log "One contract in cache for contract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
    Dim lFuture As IFuture
    If mReturnTwsContracts Then
        Set lFuture = CreateFuture(lCachedTwsContracts, mCookie)
        mFutureWaiter.Add lFuture
    Else
        If Not mListener Is Nothing Then mListener.NotifyContract mCookie, lCachedContracts.ItemAtIndex(1)
        Set lFuture = CreateFuture(lCachedContracts, mCookie)
        mFutureWaiter.Add lFuture
    End If
    Set Fetch = lFuture
Else
    ' the required contracts are in the cache but there are more than 1 of them
    
    If mReturnTwsContracts Then
        gLogger.Log lCachedTwsContracts.Count & " tws contracts in cache for contract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
        Set Fetch = CreateFuture(lCachedTwsContracts, mCookie)
    Else
        gLogger.Log lCachedContracts.Count & " contracts in cache for contract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
        ' so just play them back to the listener
        replayCachedContracts lCachedContracts, lCachedTwsContracts
        mFutureWaiter.Add mResultsContractsFutureBuilder.Future
        Set Fetch = mResultsContractsFutureBuilder.Future
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function FetchSorted( _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByRef pSortkeys() As ContractSortKeyIds, _
                ByVal pSortDescending As Boolean, _
                ByVal pCookie As Variant, _
                ByVal pReturnTwsContracts As Boolean) As IFuture
Const ProcName As String = "FetchSorted"
On Error GoTo Err

Dim ar() As ContractSortKeyIds
ar = pSortkeys
mResultsContractsBuilder.SortKeys = ar
mResultsContractsBuilder.SortAscending = Not pSortDescending

Set FetchSorted = Fetch(pContractRequester, _
                    pContractCache, _
                    pContractSpecifier, _
                    Nothing, _
                    pCookie, _
                    pReturnTwsContracts)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub NotifyContract( _
            ByVal pContract As IContract)
Const ProcName As String = "NotifyContract"
On Error GoTo Err

If Not mResultsContractsBuilder.Contracts.Contains(pContract.Specifier) Then
    mResultsContractsBuilder.Add pContract
    If Not mListener Is Nothing Then mListener.NotifyContract mCookie, pContract
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyFetchCompleted( _
                ByVal pFetcher As ContractsFetcher)
Const ProcName As String = "NotifyFetchCompleted"
On Error GoTo Err

mNumOutstandingRequests = mNumOutstandingRequests - 1
gLogger.Log "Fetch completed: mNumOutstandingRequests=" & mNumOutstandingRequests, ProcName, ModuleName, LogLevelMediumDetail

If mNumOutstandingRequests = 0 Then processEnd
requestCompleted pFetcher

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyFetchError( _
                ByVal pFetcher As ContractsFetcher, _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "NotifyFetchError"
On Error GoTo Err

mNumOutstandingRequests = mNumOutstandingRequests - 1

If pErrorCode <> 200 Then
    requestCompleted pFetcher
    
    cancelOutstandingRequests
        
    mResultsContractsFutureBuilder.Fail pErrorCode, pErrorMsg, ""
    If Not mListener Is Nothing Then mListener.FetchFailed mCookie, pErrorCode, pErrorMsg, ""
    Set mSelfRef = Nothing
ElseIf mNumOutstandingRequests = 0 Then
    processEnd
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTwsContract( _
            ByVal pTwsContract As TwsContract)
Const ProcName As String = "NotifyTwsContract"
On Error GoTo Err

If Not mTwsContracts.Contains(pTwsContract) Then
    mTwsContracts.Add pTwsContract
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub cancelOutstandingRequests()
Const ProcName As String = "cancelOutstandingRequests"
On Error GoTo Err

Dim en As Enumerator: Set en = mFetchers.Enumerator
Do While en.MoveNext
    Dim lFetcher As ContractsFetcher
    Set lFetcher = en.Current
    lFetcher.Cancel
Loop
mFetchers.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doCancel()
Const ProcName As String = "doCancel"
On Error GoTo Err

cancelOutstandingRequests

mResultsContractsFutureBuilder.Cancel
If Not mListener Is Nothing Then mListener.FetchCancelled mCookie

Set mSelfRef = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processEnd()
Const ProcName As String = "processEnd"
On Error GoTo Err

If mReturnTwsContracts Then
    gLogger.Log "Retrieved " & _
                mTwsContracts.Count & _
                " TWS contracts for: " & mContractSpecifier.ToString, _
                ProcName, ModuleName, LogLevelMediumDetail
    If mResultsContractsFutureBuilder.Future.IsPending Then
        mResultsContractsFutureBuilder.Value = mTwsContracts
        mResultsContractsFutureBuilder.Complete
    End If
Else
    gLogger.Log "Retrieved " & mResultsContractsBuilder.Contracts.Count & " contracts for: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelMediumDetail

    If mResultsContractsFutureBuilder.Future.IsPending Then
        mResultsContractsFutureBuilder.Value = mResultsContractsBuilder.Contracts
        mResultsContractsFutureBuilder.Complete
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub replayCachedContracts( _
                pCachedContracts As IContracts, _
                pCachedTwsContracts As EnumerableCollection)
Const ProcName As String = "replayCachedContracts"
On Error GoTo Err

Dim lFetcher As New ContractsFetcher
mFetchers.Add lFetcher
lFetcher.ReplayContracts Me, pCachedContracts, pCachedTwsContracts

mNumOutstandingRequests = mNumOutstandingRequests + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub requestCompleted( _
                ByVal pFetcher As ContractsFetcher)
Const ProcName As String = "requestCompleted"
On Error GoTo Err

mFetchers.Remove pFetcher
gLogger.Log "Removed ContractsFetcher for: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelMediumDetail

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub requestContracts(ByVal pContractSpec As IContractSpecifier)
Const ProcName As String = "requestContracts"
On Error GoTo Err

Dim lFetcher As New ContractsFetcher
mFetchers.Add lFetcher

lFetcher.FetchContracts Me, _
                        mContractRequester, _
                        mContractCache, _
                        pContractSpec, _
                        mReturnTwsContracts, _
                        mCookie

mNumOutstandingRequests = mNumOutstandingRequests + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub requestContractsForCash()
Const ProcName As String = "requestContractsForCash"
On Error GoTo Err

requestContracts CreateContractSpecifier( _
                                        mContractSpecifier.LocalSymbol, _
                                        mContractSpecifier.Symbol, _
                                        mContractSpecifier.TradingClass, _
                                        mContractSpecifier.Exchange, _
                                        SecTypeCash, _
                                        mContractSpecifier.CurrencyCode, _
                                        "", _
                                        mContractSpecifier.Multiplier)
If mContractSpecifier.LocalSymbol = "" Then
    requestContracts CreateContractSpecifier( _
                                        mContractSpecifier.Symbol, _
                                        "", _
                                        mContractSpecifier.TradingClass, _
                                        mContractSpecifier.Exchange, _
                                        SecTypeCash, _
                                        mContractSpecifier.CurrencyCode, _
                                        "", _
                                        mContractSpecifier.Multiplier)
End If

If mContractSpecifier.Symbol = "" Then
    requestContracts CreateContractSpecifier( _
                                        "", _
                                        mContractSpecifier.LocalSymbol, _
                                        mContractSpecifier.TradingClass, _
                                        mContractSpecifier.Exchange, _
                                        SecTypeCash, _
                                        mContractSpecifier.CurrencyCode, _
                                        "", _
                                        mContractSpecifier.Multiplier)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName & "(LocalSymbol=" & mContractSpecifier.LocalSymbol & "; Symbol=" & mContractSpecifier.Symbol & ")", ModuleName
End Sub

Private Sub requestContractsForUnspecifiedSectype()
Const ProcName As String = "requestContractsForUnspecifiedSectype"
On Error GoTo Err

' no Sectype supplied, so we'll place a separate request for each of the following:
'       Futures
'       Indexes
'       Stocks
'       Cash
' but we restrict options in this mechanism because there can be thousands of
' them, and warrants are not included because they are hardly used so it's not
' worth making an extra request every time.

If Not gIsSmartExchange(mContractSpecifier.Exchange) Then
    requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.TradingClass, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeFuture, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier)
    
    If mContractSpecifier.LocalSymbol = "" Then _
        requestContracts CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                        "", _
                                                        mContractSpecifier.TradingClass, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeFuture, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        mContractSpecifier.Expiry, _
                                                        mContractSpecifier.Multiplier)
    
    If mContractSpecifier.Symbol = "" Then _
        requestContracts CreateContractSpecifier("", _
                                                        mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.TradingClass, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeFuture, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        mContractSpecifier.Expiry, _
                                                        mContractSpecifier.Multiplier)
End If

If mContractSpecifier.Expiry = "" And _
    mContractSpecifier.Strike = 0 And _
    mContractSpecifier.Right = OptNone _
Then
    If Not gIsSmartExchange(mContractSpecifier.Exchange) Then
        requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.Symbol, _
                                                        mContractSpecifier.TradingClass, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeIndex, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier)
    
        If mContractSpecifier.LocalSymbol = "" Then _
            requestContracts CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                        "", _
                                                        mContractSpecifier.TradingClass, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeIndex, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier)
    
        If mContractSpecifier.Symbol = "" Then _
            requestContracts CreateContractSpecifier("", _
                                                        mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.TradingClass, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeIndex, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier)
        
        requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.Symbol, _
                                                        mContractSpecifier.TradingClass, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeCash, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier)

        If mContractSpecifier.LocalSymbol = "" Then _
            requestContracts CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                        "", _
                                                        mContractSpecifier.TradingClass, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeCash, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier)
    
        If mContractSpecifier.Symbol = "" Then _
            requestContracts CreateContractSpecifier("", _
                                                        mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.TradingClass, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeCash, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier)
    
    End If
    
    requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.TradingClass, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeStock, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    "", _
                                                    mContractSpecifier.Multiplier)

    If mContractSpecifier.LocalSymbol = "" Then _
        requestContracts CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                    "", _
                                                    mContractSpecifier.TradingClass, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeStock, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    "", _
                                                    mContractSpecifier.Multiplier)

    If mContractSpecifier.Symbol = "" Then _
        requestContracts CreateContractSpecifier("", _
                                                    mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.TradingClass, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeStock, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    "", _
                                                    mContractSpecifier.Multiplier)

Else
    
    ' for Options and Future Options, one of the following sets of information must be
    ' specified, to avoid fetching the entire option chain:
    '     - localsymbol and exchange (should return 0 or 1 contract)
    '     [- exchange and expiry (returns all strikes and rights): this is
    '        no longer allowed, but can be achieved by specifying the relevant
    '        sectype]
    '     - exchange and strike (returns all expiries for that strike)
    '     - expiry and strike (returns all rights and exchanges for that strike)
    If (mContractSpecifier.LocalSymbol <> "" And mContractSpecifier.Exchange <> "") Or _
        (mContractSpecifier.Exchange <> "" And mContractSpecifier.Strike <> 0#) Or _
        (mContractSpecifier.Expiry <> "" And mContractSpecifier.Strike <> 0#) _
    Then
        requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.TradingClass, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeOption, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier, _
                                                    mContractSpecifier.Strike, _
                                                    mContractSpecifier.Right)
        requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.TradingClass, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeFuturesOption, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier, _
                                                    mContractSpecifier.Strike, _
                                                    mContractSpecifier.Right)
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

