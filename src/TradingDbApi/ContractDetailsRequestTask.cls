VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractDtlsRequestTask"
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

Implements IContractFetchListener
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

Private Const ModuleName                            As String = "ContractDtlsRequestTask"

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpecifier                          As IContractSpecifier

Private mTaskContext                                As TaskContext

Private mNumOutstandingRequests                     As Long

Private mTradingDB                                  As TradingDB

Private mContractsBuilder                           As IContractsBuilder

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private WithEvents mContractsFutureBuilder          As FutureBuilder
Attribute mContractsFutureBuilder.VB_VarHelpID = -1

Private mListener                                   As IContractFetchListener

Private mCookie                                     As Variant

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mContractsBuilder = New ContractsBuilder
Set mContractsFutureBuilder = New FutureBuilder
Set mFutureWaiter = New FutureWaiter

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IContractFetchListener Interface Members
'@================================================================================

Private Sub IContractFetchListener_FetchCancelled(ByVal pCookie As Variant)
Const ProcName As String = "IContractFetchListener_FetchCancelled"
On Error GoTo Err

If mContractsFutureBuilder.Future.IsPending Then mContractsFutureBuilder.Cancel
If Not mListener Is Nothing Then mListener.FetchCancelled pCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_FetchCompleted(ByVal pCookie As Variant)
Const ProcName As String = "IContractFetchListener_FetchCompleted"
On Error GoTo Err

mNumOutstandingRequests = mNumOutstandingRequests - 1

If mNumOutstandingRequests > 0 Then Exit Sub

gLogger.Log "Received contract details for", ProcName, ModuleName, LogLevelDetail, mContractSpecifier.ToString

mContractsFutureBuilder.Value = mContractsBuilder.Contracts
mContractsFutureBuilder.Complete

mFutureWaiter.Add mContractsFutureBuilder.Future

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_FetchFailed(ByVal pCookie As Variant, ByVal pErrorCode As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
Const ProcName As String = "IContractFetchListener_FetchFailed"
On Error GoTo Err

If Not mListener Is Nothing Then mListener.FetchFailed pCookie, pErrorCode, pErrorMessage, pErrorSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_NotifyContract(ByVal pCookie As Variant, ByVal pContract As ContractUtils27.IContract)
Const ProcName As String = "IContractFetchListener_NotifyContract"
On Error GoTo Err

mContractsBuilder.Add pContract
If Not mListener Is Nothing Then mListener.NotifyContract pCookie, pContract

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

If mTradingDB Is Nothing Then
    mTaskContext.Suspend -1
    Exit Sub
End If

If mContractSpecifier.SecType = SecTypeNone Then
    ' no Sectype supplied, so we'll place a separate Request for each of the following:
    '       Futures
    '       Indexes
    '       Stocks
    '       Cash
    requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeFuture, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier)
    
    If mContractSpecifier.LocalSymbol = "" Then
        requestContracts CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                        "", _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeFuture, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        mContractSpecifier.Expiry, _
                                                        mContractSpecifier.Multiplier)
    End If
    
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

    requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeCash, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry, _
                                                    mContractSpecifier.Multiplier)

Else
    requestContracts mContractSpecifier
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

If ev.Future.IsCancelled Then
    doCancel
ElseIf ev.Future.IsFaulted Then
    doFail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
    mFutureWaiter.Clear
ElseIf TypeOf ev.Future.Value Is TradingDB Then
    Set mTradingDB = ev.Future.Value
    mTaskContext.Continue
ElseIf ev.Future Is mContractsFutureBuilder.Future Then
    If Not mListener Is Nothing Then mListener.FetchCompleted mCookie
    mTaskContext.Finish Empty, False
End If
    
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ContractsFuture() As IFuture
Const ProcName As String = "ContractsFuture"
On Error GoTo Err

Set ContractsFuture = mContractsFutureBuilder.Future

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pTradingDBFuture As IFuture, _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pListener As IContractFetchListener, _
                ByVal pCookie As Variant)
Const ProcName As String = "Initialise"
On Error GoTo Err

If pTradingDBFuture.IsAvailable Then
    Set mTradingDB = pTradingDBFuture.Value
Else
    mFutureWaiter.Add pTradingDBFuture
End If

Set mContractSpecifier = pContractSpecifier
Set mListener = pListener

mContractsBuilder.ContractSpecifier = mContractSpecifier

gSetVariant mCookie, pCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub doCancel()
Const ProcName As String = "doCancel"
On Error GoTo Err

mFutureWaiter.Cancel

If Not mListener Is Nothing Then mListener.FetchCancelled mCookie

If mContractsFutureBuilder.Future.IsPending Then mContractsFutureBuilder.Cancel

mTaskContext.Finish Nothing, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doFail( _
                ByVal pErrorNumber As Long, _
                ByVal pErrorMessage As String, _
                ByVal pErrorSource As String)
Const ProcName As String = "doFail"
On Error GoTo Err

If Not mListener Is Nothing Then mListener.FetchFailed mCookie, pErrorNumber, pErrorMessage, pErrorSource
If mContractsFutureBuilder.Future.IsPending Then mContractsFutureBuilder.Fail pErrorNumber, pErrorMessage, pErrorSource

mTaskContext.Error pErrorNumber, pErrorMessage, pErrorSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub requestContracts(ByVal pContractSpec As IContractSpecifier)
Const ProcName As String = "requestContracts"
On Error GoTo Err

If gLogger.IsLoggable(LogLevelDetail) Then gLogger.Log "Request contract details for", ProcName, ModuleName, LogLevelDetail, pContractSpec.ToString
mFutureWaiter.Add mTradingDB.FetchContractsAsync(pContractSpec, Me, mCookie)
mNumOutstandingRequests = mNumOutstandingRequests + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




