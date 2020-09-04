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

Private mContractsBuilder                           As ContractsBuilder
Private mCachedContracts                            As IContracts

Private mNumOutstandingRequests                     As Long

Private mFetchers                                   As EnumerableCollection

Private WithEvents mResultContractsFutureBuilder    As FutureBuilder
Attribute mResultContractsFutureBuilder.VB_VarHelpID = -1
Private mCacheContractsFutureBuilder                As FutureBuilder
Attribute mCacheContractsFutureBuilder.VB_VarHelpID = -1

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
Set mContractsBuilder = New ContractsBuilder
Set mResultContractsFutureBuilder = New FutureBuilder
Set mCacheContractsFutureBuilder = New FutureBuilder
Set mTwsContracts = New EnumerableCollection
Set mFetchers = New EnumerableCollection
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

If Not mListener Is Nothing Then mListener.FetchCompleted mCookie
Set mSelfRef = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mResultContractsFutureBuilder Event Handlers
'@================================================================================

Private Sub mResultContractsFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mResultContractsFutureBuilder_Cancelled"
On Error GoTo Err

doCancel

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ContractsFuture() As IFuture
Set ContractsFuture = mResultContractsFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Fetch( _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pListener As IContractFetchListener, _
                ByVal pCookie As Variant, _
                ByVal pReturnTwsContracts As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mSelfRef = Me

Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
Set mContractSpecifier = pContractSpecifier
Set mListener = pListener

gSetVariant mCookie, pCookie
mResultContractsFutureBuilder.Cookie = mCookie

mReturnTwsContracts = pReturnTwsContracts

Set mCachedContracts = mContractCache.GetContracts(mContractSpecifier)
If mCachedContracts Is Nothing Then
ElseIf mCachedContracts.Count = 1 Then
    If Not mListener Is Nothing Then mListener.NotifyContract mCookie, mCachedContracts.ItemAtIndex(1)
    If Not mListener Is Nothing Then mListener.FetchCompleted mCookie
    If mReturnTwsContracts Then
        mResultContractsFutureBuilder.Value = mContractCache.GetTwsContracts(mContractSpecifier)
        mResultContractsFutureBuilder.Complete
    Else
        mResultContractsFutureBuilder.Value = mCachedContracts
        mResultContractsFutureBuilder.Complete
    End If
    Exit Sub
End If

If mContractSpecifier.SecType <> SecTypeNone Then
    requestContracts mContractSpecifier
ElseIf Not mCachedContracts Is Nothing Then
     requestContracts mContractSpecifier
Else
    mContractCache.SetContractsFutureBuilder mContractSpecifier, _
                                            mCacheContractsFutureBuilder, _
                                            mTwsContracts
    requestContractsForUnspecifiedSectype
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyContract( _
            ByVal pContract As IContract)
Const ProcName As String = "NotifyContract"
On Error GoTo Err

If Not mContractsBuilder.Contracts.Contains(pContract.Specifier) Then
    mContractsBuilder.Add pContract
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

requestCompleted pFetcher
If mNumOutstandingRequests = 0 Then processEnd

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

requestCompleted pFetcher

If pErrorCode <> 200 Then
    
    cancelOutstandingRequests
        
    mResultContractsFutureBuilder.Fail pErrorCode, pErrorMsg, ""
    mCacheContractsFutureBuilder.Fail pErrorCode, pErrorMsg, ""
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

mTwsContracts.Add pTwsContract

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

mResultContractsFutureBuilder.Cancel
mCacheContractsFutureBuilder.Cancel
If Not mListener Is Nothing Then mListener.FetchCancelled mCookie

Set mSelfRef = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processEnd()
Const ProcName As String = "processEnd"
On Error GoTo Err

If mResultContractsFutureBuilder.Future.IsPending Then
    mFutureWaiter.Add mResultContractsFutureBuilder.Future
    If mReturnTwsContracts Then
        mResultContractsFutureBuilder.Value = mTwsContracts
        mResultContractsFutureBuilder.Complete
    Else
        mResultContractsFutureBuilder.Value = mContractsBuilder.Contracts
        mResultContractsFutureBuilder.Complete
    End If
End If

If mReturnTwsContracts Then
    mCacheContractsFutureBuilder.Value = mTwsContracts
    mCacheContractsFutureBuilder.Complete
Else
    mCacheContractsFutureBuilder.Value = mContractsBuilder.Contracts
    mCacheContractsFutureBuilder.Complete
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub requestCompleted( _
                ByVal pFetcher As ContractsFetcher)
Const ProcName As String = "requestCompleted"
On Error GoTo Err

mFetchers.Remove pFetcher
mNumOutstandingRequests = mNumOutstandingRequests - 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub requestContracts(ByVal pContractSpec As IContractSpecifier)
Const ProcName As String = "requestContracts"
On Error GoTo Err

Dim lFetcher As New ContractsFetcher
mFetchers.Add lFetcher
lFetcher.FetchContracts Me, mContractRequester, mContractCache, pContractSpec

mNumOutstandingRequests = mNumOutstandingRequests + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub requestContractsForNonSMART()
Const ProcName As String = "requestContractsForNonSMART"
On Error GoTo Err

Dim s As String

If IsContractSpecOffsetExpiry(mContractSpecifier) Then
    
    ' only Futures, Options and future options can specify expiry in this format
    
    ' for options and future options, the exchange, strike and right must be specified, since we're targetting
    ' a single contract
    If mContractSpecifier.Strike <> 0# Or mContractSpecifier.Right <> OptNone Then
        If mContractSpecifier.Exchange = "" Or mContractSpecifier.Strike = 0# Or mContractSpecifier.Right = OptNone Then
            s = "Exchange, Strike and Right must be specified for offset options"
            If Not mListener Is Nothing Then mListener.FetchFailed mCookie, -1, s, ""
            Set mSelfRef = Nothing
        Else
            requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                            mContractSpecifier.Symbol, _
                                                            mContractSpecifier.Exchange, _
                                                            SecTypeOption, _
                                                            mContractSpecifier.CurrencyCode, _
                                                            "", _
                                                            mContractSpecifier.Multiplier, _
                                                            mContractSpecifier.Strike, _
                                                            mContractSpecifier.Right)
            requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                            mContractSpecifier.Symbol, _
                                                            mContractSpecifier.Exchange, _
                                                            SecTypeFuturesOption, _
                                                            mContractSpecifier.CurrencyCode, _
                                                            "", _
                                                            mContractSpecifier.Multiplier, _
                                                            mContractSpecifier.Strike, _
                                                            mContractSpecifier.Right)
        End If
    Else
        requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.Symbol, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeFuture, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier)
        
        If mContractSpecifier.LocalSymbol = "" Then _
            requestContracts CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                            "", _
                                                            mContractSpecifier.Exchange, _
                                                            SecTypeFuture, _
                                                            mContractSpecifier.CurrencyCode, _
                                                            "", _
                                                            mContractSpecifier.Multiplier)
    
        If mContractSpecifier.Symbol = "" Then _
            requestContracts CreateContractSpecifier("", _
                                                            mContractSpecifier.LocalSymbol, _
                                                            mContractSpecifier.Exchange, _
                                                            SecTypeFuture, _
                                                            mContractSpecifier.CurrencyCode, _
                                                            "", _
                                                            mContractSpecifier.Multiplier)
    End If
Else
    
    requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeFuture, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier)
    
    If mContractSpecifier.LocalSymbol = "" Then _
        requestContracts CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                        "", _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeFuture, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        mContractSpecifier.Expiry, _
                                                        mContractSpecifier.Multiplier)
    
    If mContractSpecifier.Symbol = "" Then _
        requestContracts CreateContractSpecifier("", _
                                                        mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeFuture, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        mContractSpecifier.Expiry, _
                                                        mContractSpecifier.Multiplier)
    
    requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeIndex, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier)
    
    requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeStock, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier)
    
    If mContractSpecifier.LocalSymbol = "" Then _
        requestContracts CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                        "", _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeStock, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        mContractSpecifier.Expiry, _
                                                        mContractSpecifier.Multiplier)
    
    If mContractSpecifier.Symbol = "" Then _
        requestContracts CreateContractSpecifier("", _
                                                        mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeStock, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        mContractSpecifier.Expiry, _
                                                        mContractSpecifier.Multiplier)
    
    requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeCash, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier)

    ' for Options and Future Options, one of the following sets of information must be specified, to
    ' avoid fetching the entire option chain:
    '     - localsymbol and exchange (should return 0 or 1 contract)
    '     - exchange and expiry (returns all strikes and rights)
    '     - exchange and strike (returns all expiries for that strike)
    '     - expiry and strike (returns all rights and exchanges for that strike)
    If mContractSpecifier.LocalSymbol <> "" And mContractSpecifier.Exchange <> "" Or _
        mContractSpecifier.Exchange <> "" And mContractSpecifier.Expiry <> "" Or _
        mContractSpecifier.Exchange <> "" And mContractSpecifier.Strike <> 0# Or _
        mContractSpecifier.Expiry <> "" And mContractSpecifier.Strike <> 0# _
    Then
        requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeOption, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier, _
                                                    mContractSpecifier.Strike, _
                                                    mContractSpecifier.Right)
        requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
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

Private Sub requestContractsForSMART()
Const ProcName As String = "requestContractsForSMART"
On Error GoTo Err

' SMART exchanges only apply to stocks, options and future options
If IsContractSpecOffsetExpiry(mContractSpecifier) Then
    
    ' only options and future options can specify expiry in this format.
    ' for options, the strike and right must be specified, since we're targetting
    ' a single contract
    If mContractSpecifier.Strike = 0# Or mContractSpecifier.Right = OptNone Then
        Dim s As String: s = "Strike and Right must be specified for offset options"
        If Not mListener Is Nothing Then mListener.FetchFailed mCookie, -1, s, ""
        Set mSelfRef = Nothing
    Else
        requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.Symbol, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeOption, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier)
    
        requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.Symbol, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeFuturesOption, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier)
    End If
Else
    requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeStock, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    "", _
                                                    mContractSpecifier.Multiplier)
    
    If mContractSpecifier.LocalSymbol = "" Then
        requestContracts CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                        "", _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeStock, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier)
    End If

    If mContractSpecifier.Symbol = "" Then
        requestContracts CreateContractSpecifier("", _
                                                        mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeStock, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier)
    End If

    ' for Options and Future Options, one of the following sets of information must be specified, to
    ' avoid fetching the entire option chain (note that we already know the exchange):
    '     - localsymbol (should return 0 or 1 contract)
    '     - expiry (returns all strikes and rights for the SMART exchange)
    '     - strike (returns all expiries for that strike)
    '     - expiry and strike (returns all rights for that strike)
    If mContractSpecifier.LocalSymbol <> "" Or _
        mContractSpecifier.Expiry <> "" Or _
        mContractSpecifier.Strike <> 0# Or _
        mContractSpecifier.Expiry <> "" And mContractSpecifier.Strike <> 0# _
    Then
        requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeOption, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier, _
                                                    mContractSpecifier.Strike, _
                                                    mContractSpecifier.Right)
        requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
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

Private Sub requestContractsForUnspecifiedSectype()
Const ProcName As String = "requestContractsForUnspecifiedSectype"
On Error GoTo Err

' no Sectype supplied, so we'll place a separate request for each of the following:
'       Futures
'       Indexes
'       Stocks
'       Cash
' but we don't include options in this mechanism because there can be thousands of
' them

If gIsSmartExchange(mContractSpecifier.Exchange) Then
    requestContractsForSMART
Else
    requestContractsForNonSMART
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

