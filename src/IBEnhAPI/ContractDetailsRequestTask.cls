VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractDetailsRequestTask"
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

Private Const ModuleName                            As String = "ContractDetailsRequestTask"

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpecifier                          As IContractSpecifier
Private mContractRequester                          As ContractDetailsRequester

Private mTaskContext                                As TaskContext

Private mContractsBuilder                           As ContractsBuilder

Private mNumOutstandingRequests                     As Long

Private mRequestId                                  As Long
Private mRequestId1                                 As Long
Private mRequestId2                                 As Long
Private mRequestId3                                 As Long
Private mRequestId4                                 As Long
Private mRequestId5                                 As Long
Private mRequestId6                                 As Long
Private mRequestId7                                 As Long
Private mRequestId8                                 As Long
Private mRequestId9                                 As Long
Private mRequestId10                                As Long

Private WithEvents mContractsFutureBuilder          As FutureBuilder
Attribute mContractsFutureBuilder.VB_VarHelpID = -1

Private mListener                                   As IContractFetchListener

Private mCookie                                     As Variant

Private mReturnTwsContracts                         As Boolean

Private mTwsContracts                               As EnumerableCollection

Private mContractCache                              As TwsContractCache

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
Set mContractsBuilder = New ContractsBuilder
Set mContractsFutureBuilder = New FutureBuilder
Set mTwsContracts = New EnumerableCollection
mRequestId = NullIndex
mRequestId1 = NullIndex
mRequestId2 = NullIndex
mRequestId3 = NullIndex
mRequestId4 = NullIndex
mRequestId5 = NullIndex
mRequestId6 = NullIndex
mRequestId7 = NullIndex
mRequestId8 = NullIndex
mRequestId9 = NullIndex
mRequestId10 = NullIndex
End Sub

'@================================================================================
' IContractDetailsConsumer Interface Members
'@================================================================================

Private Sub IContractDetailsConsumer_EndContractDetails(ByVal pRequestId As Long)
Const ProcName As String = "IContractDetailsConsumer_EndContractDetails"
On Error GoTo Err

requestCompleted pRequestId
mNumOutstandingRequests = mNumOutstandingRequests - 1
If mNumOutstandingRequests = 0 Then processEnd

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_EndOptionParameters(ByVal pRequestIs As Long)

End Sub

Private Sub IContractDetailsConsumer_NotifyContract(ByVal pRequestId As Long, pContractDetails As TwsContractDetails)
Const ProcName As String = "IContractDetailsConsumer_NotifyContract"
On Error GoTo Err

processContract pContractDetails

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IContractDetailsConsumer_NotifyError"
On Error GoTo Err

requestCompleted pRequestId
mNumOutstandingRequests = mNumOutstandingRequests - 1

If mContractsFutureBuilder.Future.IsCancelled Then Exit Sub

If pErrorCode <> 200 Then
    
    cancelOutstandingRequests
    mTaskContext.Error pErrorCode, pErrorMsg
    
    mContractsFutureBuilder.Fail pErrorCode, pErrorMsg, ""
    If Not mListener Is Nothing Then mListener.FetchFailed mCookie, pErrorCode, pErrorMsg, ""

ElseIf mNumOutstandingRequests = 0 Then
    mFutureWaiter.Add mContractsFutureBuilder.Future
    If mReturnTwsContracts Then
        mContractsFutureBuilder.Value = mTwsContracts
        mContractsFutureBuilder.Complete
    Else
        mContractsFutureBuilder.Value = mContractsBuilder.Contracts
        mContractsFutureBuilder.Complete
    End If
End If

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
doCancel
End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    doCancel
    Exit Sub
End If

If mContractSpecifier.SecType = SecTypeNone Then
    ' no Sectype supplied, so we'll place a separate Request for each of the following:
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
Else
    Dim lTwsContract As TwsContractDetails
    Set lTwsContract = mContractCache.Item(mContractSpecifier)
    
    If Not lTwsContract Is Nothing Then
        gLogger.Log "TwsContractDetails retrieved from cache for IContract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
        processContract lTwsContract
        processEnd
        mTaskContext.Suspend -1
        Exit Sub
    End If
    
    mRequestId = requestContracts(mContractSpecifier)
End If

mTaskContext.Suspend -1

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
' mContractsFutureBuilder Event Handlers
'@================================================================================

Private Sub mContractsFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mContractsFutureBuilder_Cancelled"
On Error GoTo Err

doCancel

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not mListener Is Nothing Then mListener.FetchCompleted mCookie
mTaskContext.Finish Empty, False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ContractsFuture() As IFuture
Set ContractsFuture = mContractsFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractCache As TwsContractCache, _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pListener As IContractFetchListener, _
                ByVal pCookie As Variant, _
                ByVal pReturnTwsContracts As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
Set mContractSpecifier = pContractSpecifier
Set mListener = pListener

gSetVariant mCookie, pCookie
mContractsFutureBuilder.Cookie = mCookie

mReturnTwsContracts = pReturnTwsContracts

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

cancelRequest mRequestId
cancelRequest mRequestId1
cancelRequest mRequestId2
cancelRequest mRequestId3
cancelRequest mRequestId4
cancelRequest mRequestId5
cancelRequest mRequestId6
cancelRequest mRequestId7
cancelRequest mRequestId8
cancelRequest mRequestId9
cancelRequest mRequestId10

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub cancelRequest(ByVal pRequestId As Long)
Const ProcName As String = "cancelRequest"
On Error GoTo Err

If pRequestId <> NullIndex Then mContractRequester.cancelRequest pRequestId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doCancel()
Const ProcName As String = "doCancel"
On Error GoTo Err

cancelOutstandingRequests

mTaskContext.Finish Nothing, True

mContractsFutureBuilder.Cancel
If Not mListener Is Nothing Then mListener.FetchCancelled mCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processContract(ByVal pContractDetails As TwsContractDetails)
Const ProcName As String = "processContract"
On Error GoTo Err

If mContractsFutureBuilder.Future.IsCancelled Then Exit Sub

If IsValidExchangeCode(pContractDetails.Summary.Exchange) Then
    If mReturnTwsContracts Then
        mTwsContracts.Add pContractDetails
    Else
        Dim lContract As IContract
        Set lContract = gTwsContractDetailsToContract(pContractDetails)
        
        If Not mContractsBuilder.Contracts.Contains(lContract.Specifier) Then
            mContractsBuilder.Add lContract
            If Not mListener Is Nothing Then mListener.NotifyContract mCookie, lContract
        End If
    End If
Else
    gLog "Unknown Exchange code received from Tws", pMsgQualifier:=pContractDetails.Summary.Exchange, pModName:=ModuleName, pProcName:=ProcName
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processEnd()
Const ProcName As String = "processEnd"
On Error GoTo Err

If mContractsFutureBuilder.Future.IsCancelled Then Exit Sub

mFutureWaiter.Add mContractsFutureBuilder.Future
If mReturnTwsContracts Then
    mContractsFutureBuilder.Value = mTwsContracts
    mContractsFutureBuilder.Complete
Else
    mContractsFutureBuilder.Value = mContractsBuilder.Contracts
    mContractsFutureBuilder.Complete
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function requestContracts(ByVal pContractSpec As IContractSpecifier) As Long
Const ProcName As String = "requestContracts"
On Error GoTo Err

If gLogger.IsLoggable(LogLevelDetail) Then gLog "Request contract details for", ModuleName, ProcName, pContractSpec.ToString, LogLevelDetail
requestContracts = mContractRequester.RequestContractDetails(pContractSpec, Me)
mNumOutstandingRequests = mNumOutstandingRequests + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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
            mTaskContext.Error -1, s
        Else
            mRequestId1 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                            mContractSpecifier.Symbol, _
                                                            mContractSpecifier.Exchange, _
                                                            SecTypeOption, _
                                                            mContractSpecifier.CurrencyCode, _
                                                            "", _
                                                            mContractSpecifier.Multiplier, _
                                                            mContractSpecifier.Strike, _
                                                            mContractSpecifier.Right))
            mRequestId2 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                            mContractSpecifier.Symbol, _
                                                            mContractSpecifier.Exchange, _
                                                            SecTypeFuturesOption, _
                                                            mContractSpecifier.CurrencyCode, _
                                                            "", _
                                                            mContractSpecifier.Multiplier, _
                                                            mContractSpecifier.Strike, _
                                                            mContractSpecifier.Right))
        End If
    Else
        mRequestId1 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.Symbol, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeFuture, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier))
        
        If mContractSpecifier.LocalSymbol = "" Then _
            mRequestId2 = requestContracts(CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                            "", _
                                                            mContractSpecifier.Exchange, _
                                                            SecTypeFuture, _
                                                            mContractSpecifier.CurrencyCode, _
                                                            "", _
                                                            mContractSpecifier.Multiplier))
    
        If mContractSpecifier.Symbol = "" Then _
            mRequestId3 = requestContracts(CreateContractSpecifier("", _
                                                            mContractSpecifier.LocalSymbol, _
                                                            mContractSpecifier.Exchange, _
                                                            SecTypeFuture, _
                                                            mContractSpecifier.CurrencyCode, _
                                                            "", _
                                                            mContractSpecifier.Multiplier))
    End If
Else
    
    mRequestId1 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeFuture, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier))
    
    If mContractSpecifier.LocalSymbol = "" Then _
        mRequestId2 = requestContracts(CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                        "", _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeFuture, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        mContractSpecifier.Expiry, _
                                                        mContractSpecifier.Multiplier))
    
    If mContractSpecifier.Symbol = "" Then _
        mRequestId3 = requestContracts(CreateContractSpecifier("", _
                                                        mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeFuture, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        mContractSpecifier.Expiry, _
                                                        mContractSpecifier.Multiplier))
    
    mRequestId4 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeIndex, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier))
    
    mRequestId5 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeStock, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier))
    
    If mContractSpecifier.LocalSymbol = "" Then _
        mRequestId6 = requestContracts(CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                        "", _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeStock, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        mContractSpecifier.Expiry, _
                                                        mContractSpecifier.Multiplier))
    
    If mContractSpecifier.Symbol = "" Then _
        mRequestId7 = requestContracts(CreateContractSpecifier("", _
                                                        mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeStock, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        mContractSpecifier.Expiry, _
                                                        mContractSpecifier.Multiplier))
    
    mRequestId8 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeCash, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier))

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
        mRequestId9 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeOption, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier, _
                                                    mContractSpecifier.Strike, _
                                                    mContractSpecifier.Right))
        mRequestId10 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeFuturesOption, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier, _
                                                    mContractSpecifier.Strike, _
                                                    mContractSpecifier.Right))
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
        mTaskContext.Error -1, s
    Else
        mRequestId1 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.Symbol, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeOption, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier))
    
        mRequestId2 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.Symbol, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeFuturesOption, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier))
    End If
Else
    mRequestId1 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeStock, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    "", _
                                                    mContractSpecifier.Multiplier))
    
    If mContractSpecifier.LocalSymbol = "" Then
        mRequestId2 = requestContracts(CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                        "", _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeStock, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier))
    End If

    If mContractSpecifier.Symbol = "" Then
        mRequestId3 = requestContracts(CreateContractSpecifier("", _
                                                        mContractSpecifier.LocalSymbol, _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeStock, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        "", _
                                                        mContractSpecifier.Multiplier))
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
        mRequestId4 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeOption, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier, _
                                                    mContractSpecifier.Strike, _
                                                    mContractSpecifier.Right))
        mRequestId5 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeFuturesOption, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier, _
                                                    mContractSpecifier.Strike, _
                                                    mContractSpecifier.Right))
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub requestCompleted(ByVal pRequestId As Long)
If pRequestId = mRequestId Then mRequestId = NullIndex
If pRequestId = mRequestId1 Then mRequestId1 = NullIndex
If pRequestId = mRequestId2 Then mRequestId2 = NullIndex
If pRequestId = mRequestId3 Then mRequestId3 = NullIndex
If pRequestId = mRequestId4 Then mRequestId4 = NullIndex
If pRequestId = mRequestId5 Then mRequestId5 = NullIndex
If pRequestId = mRequestId6 Then mRequestId6 = NullIndex
If pRequestId = mRequestId7 Then mRequestId7 = NullIndex
End Sub

