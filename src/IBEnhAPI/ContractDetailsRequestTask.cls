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

Private WithEvents mContractsFutureBuilder          As FutureBuilder
Attribute mContractsFutureBuilder.VB_VarHelpID = -1

Private mListener                                   As IContractFetchListener

Private mCookie                                     As Variant

Private mReturnTwsContracts                         As Boolean

Private mTwsContracts                               As EnumerableCollection

Private mContractCache                              As TwsContractCache

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mContractsBuilder = New ContractsBuilder
Set mContractsFutureBuilder = New FutureBuilder
Set mTwsContracts = New EnumerableCollection
mRequestId = NullIndex
mRequestId1 = NullIndex
mRequestId2 = NullIndex
mRequestId3 = NullIndex
mRequestId4 = NullIndex
mRequestId5 = NullIndex
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

If pErrorCode <> 200 Then
    
    cancelOutstandingRequests
    mTaskContext.Error pErrorCode, pErrorMsg
    
    mContractsFutureBuilder.Fail pErrorCode, pErrorMsg, ""
    If Not mListener Is Nothing Then mListener.FetchFailed mCookie, pErrorCode, pErrorMsg, ""

ElseIf mNumOutstandingRequests = 0 Then
    If mReturnTwsContracts Then
        mContractsFutureBuilder.Value = mTwsContracts
        mContractsFutureBuilder.Complete
    Else
        mContractsFutureBuilder.Value = mContractsBuilder.Contracts
        mContractsFutureBuilder.Complete
    End If
    
    If Not mListener Is Nothing Then mListener.FetchCompleted mCookie
    mTaskContext.Finish Empty, False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

If mContractSpecifier.Sectype = SecTypeNone Then
    ' no Sectype supplied, so we'll place a separate Request for each of the following:
    '       Futures
    '       Indexes
    '       Stocks
    '       Cash
    mRequestId1 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeFuture, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry))
    
    If mContractSpecifier.LocalSymbol = "" Then
        mRequestId2 = requestContracts(CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                        "", _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeFuture, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        mContractSpecifier.Expiry))
    End If
    
    mRequestId3 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeIndex, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry))

    mRequestId4 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeStock, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry))

    mRequestId5 = requestContracts(CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeCash, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry))

Else
    Dim lTwsContract As TwsContractDetails
    Set lTwsContract = mContractCache.Item(mContractSpecifier)
    
    If Not lTwsContract Is Nothing Then
        gLogger.Log "TwsContractDetails retrieved from cache for IContract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
        processContract lTwsContract
        processEnd
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

If mRequestId <> NullIndex Then mContractRequester.CancelRequest mRequestId
If mRequestId1 <> NullIndex Then mContractRequester.CancelRequest mRequestId1
If mRequestId2 <> NullIndex Then mContractRequester.CancelRequest mRequestId2
If mRequestId3 <> NullIndex Then mContractRequester.CancelRequest mRequestId3
If mRequestId4 <> NullIndex Then mContractRequester.CancelRequest mRequestId4
If mRequestId5 <> NullIndex Then mContractRequester.CancelRequest mRequestId5

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

If IsValidExchangeCode(pContractDetails.Summary.Exchange) Then
    If mReturnTwsContracts Then
        mTwsContracts.Add pContractDetails
    Else
        Dim lContract As IContract
        Set lContract = gTwsContractDetailsToContract(pContractDetails)
        
        mContractsBuilder.Add lContract
        If Not mListener Is Nothing Then mListener.NotifyContract mCookie, lContract
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

If mReturnTwsContracts Then
    mContractsFutureBuilder.Value = mTwsContracts
    mContractsFutureBuilder.Complete
Else
    mContractsFutureBuilder.Value = mContractsBuilder.Contracts
    mContractsFutureBuilder.Complete
End If

If Not mListener Is Nothing Then mListener.FetchCompleted mCookie
mTaskContext.Finish Empty, False

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

Private Sub requestCompleted(ByVal pRequestId As Long)
If pRequestId = mRequestId Then mRequestId = NullIndex
If pRequestId = mRequestId1 Then mRequestId1 = NullIndex
If pRequestId = mRequestId2 Then mRequestId2 = NullIndex
If pRequestId = mRequestId3 Then mRequestId3 = NullIndex
If pRequestId = mRequestId4 Then mRequestId4 = NullIndex
If pRequestId = mRequestId5 Then mRequestId5 = NullIndex
End Sub
