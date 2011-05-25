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

Private Const ModuleName                            As String = "ContractDetailsRequestTask"

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpecifier                          As ContractSpecifier
Private mContractRequester                          As ContractDetailsRequester

Private mTaskContext                                As TaskContext

Private mContractsBuilder                           As ContractsBuilder

Private mNumOutstandingRequests                     As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IContractDetailsConsumer Interface Members
'@================================================================================

Private Sub IContractDetailsConsumer_EndContractDetails(ByVal pRequestId As Long)
Const ProcName As String = "IContractDetailsConsumer_EndContractDetails"
On Error GoTo Err

EndRequest pRequestId
mNumOutstandingRequests = mNumOutstandingRequests - 1
If mNumOutstandingRequests = 0 Then mTaskContext.Finish mContractsBuilder.Contracts, False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyContract(ByVal pRequestId As Long, pContractDetails As TwsContractDetails)
Const ProcName As String = "IContractDetailsConsumer_NotifyContract"
On Error GoTo Err

If IsValidExchangeCode(pContractDetails.Summary.Exchange) Then
    mContractsBuilder.AddContract gTwsContractDetailsToContract(pContractDetails)
Else
    gLog "Unknown Exchange code received from Tws", pMsgQualifier:=pContractDetails.Summary.Exchange, pModName:=ModuleName, pProcName:=ProcName
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IContractDetailsConsumer_NotifyError"
On Error GoTo Err

EndRequest pRequestId
mNumOutstandingRequests = mNumOutstandingRequests - 1

If mNumOutstandingRequests = 0 Then
    If mContractsBuilder.Contracts.Count > 0 Then
        mTaskContext.Finish mContractsBuilder.Contracts, False
    Else
        mTaskContext.Error pErrorCode, pErrorMsg
    End If
End If

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

If mContractSpecifier.Sectype = SecTypeNone Then
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
                                                    mContractSpecifier.Expiry)
    
    If mContractSpecifier.LocalSymbol = "" Then
        requestContracts CreateContractSpecifier(mContractSpecifier.Symbol, _
                                                        "", _
                                                        mContractSpecifier.Exchange, _
                                                        SecTypeFuture, _
                                                        mContractSpecifier.CurrencyCode, _
                                                        mContractSpecifier.Expiry)
    End If
    
    requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeIndex, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry)

    requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeStock, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry)

    requestContracts CreateContractSpecifier(mContractSpecifier.LocalSymbol, _
                                                    mContractSpecifier.Symbol, _
                                                    mContractSpecifier.Exchange, _
                                                    SecTypeCash, _
                                                    mContractSpecifier.CurrencyCode, _
                                                    mContractSpecifier.Expiry)

Else
    requestContracts mContractSpecifier
End If

mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Let Task_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.Name
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractsBuilder As ContractsBuilder)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractRequester = pContractRequester
Set mContractsBuilder = pContractsBuilder
Set mContractSpecifier = mContractsBuilder.Contracts.ContractSpecifier

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function EndRequest(ByVal pRequestId As Long)
Const ProcName As String = "EndRequest"
On Error GoTo Err

mContractRequester.CompleteRequest pRequestId

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function requestContracts(ByVal pContractSpec As ContractSpecifier) As Long
Const ProcName As String = "requestContracts"
On Error GoTo Err

requestContracts = mContractRequester.RequestContractDetails(pContractSpec, Me)
mNumOutstandingRequests = mNumOutstandingRequests + 1

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

