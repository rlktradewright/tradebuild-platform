VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractStore"
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

Implements IContractStore

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

Private Const ModuleName                            As String = "ContractStore"

Private Const MyCapabilities                        As Long = ContractStoreCapabilities.ContractStoreCanQuery

'@================================================================================
' Member variables
'@================================================================================

Private mTradingDBFuture                            As IFuture

Private mFutureWaiter                               As New FutureWaiter

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IContractStore Interface Members
'@================================================================================

Private Function IContractStore_FetchContracts( _
                ByVal pContractSpecifier As IContractSpecifier, _
                Optional ByVal pListener As IContractFetchListener, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "IContractStore_FetchContracts"
On Error GoTo Err

Set IContractStore_FetchContracts = FetchContracts(pContractSpecifier, pListener, pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IContractStore_Finish()
Const ProcName As String = "IContractStore_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IContractStore_Supports(ByVal pCapabilities As ContractStoreCapabilities) As Boolean
IContractStore_Supports = Supports(pCapabilities)
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Function FetchContracts( _
                ByVal pContractSpecifier As IContractSpecifier, _
                Optional ByVal pListener As IContractFetchListener, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchContracts"
On Error GoTo Err

Dim lContractDtlsRequestTask As New ContractDtlsRequestTask
lContractDtlsRequestTask.Initialise mTradingDBFuture, pContractSpecifier, pListener, pCookie
StartTask lContractDtlsRequestTask, PriorityNormal, , pCookie

Set FetchContracts = lContractDtlsRequestTask.ContractsFuture
mFutureWaiter.Add FetchContracts

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

mFutureWaiter.Cancel

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise(ByVal pTradingDBFuture As IFuture)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTradingDBFuture = pTradingDBFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Supports(ByVal pCapabilities As ContractStoreCapabilities) As Boolean
Supports = ((MyCapabilities And pCapabilities) = pCapabilities)
End Function

'@================================================================================
' Helper Functions
'@================================================================================




