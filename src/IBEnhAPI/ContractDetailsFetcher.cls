VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractDetailsFetcher"
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

Private Const ModuleName                            As String = "ContractDetailsFetcher"

'@================================================================================
' Member variables
'@================================================================================

Private mContractRequester                          As ContractDetailsRequester
Private mContractCache                              As TwsContractCache
Private mContractsFutures                           As New EnumerableCollection

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

Set FetchContracts = gFetchContracts(mContractRequester, mContractCache, pContractSpecifier, pListener, pCookie, False)
mContractsFutures.Add FetchContracts

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Dim f As IFuture
For Each f In mContractsFutures
    If f.IsPending Then f.Cancel
Next

Set mContractRequester = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractCache As TwsContractCache)
Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
End Sub

Public Function Supports(ByVal pCapabilities As ContractStoreCapabilities) As Boolean
Const Capabilities As Long = ContractStoreCapabilities.ContractStoreCanQuery
Supports = ((Capabilities And pCapabilities) = pCapabilities)
End Function

'@================================================================================
' Helper Functions
'@================================================================================




