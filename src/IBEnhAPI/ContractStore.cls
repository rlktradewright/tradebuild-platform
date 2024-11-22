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

'@================================================================================
' Member variables
'@================================================================================

Private mContractRequester                          As ContractsTwsRequester
Private mContractCache                              As ContractCache
Private mContractsFutures                           As New EnumerableCollection
Private mStrikesFutures                             As New EnumerableCollection
Private mExpiriesFutures                            As New EnumerableCollection
Private mClientID                                   As Long
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

Set IContractStore_FetchContracts = GIBEnhancedApi.FetchContracts( _
                                                mContractRequester, _
                                                mContractCache, _
                                                pContractSpecifier, _
                                                pListener, _
                                                pCookie, _
                                                False, _
                                                mClientID)
mContractsFutures.Add IContractStore_FetchContracts

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function IContractStore_FetchContractsSorted( _
                ByVal pContractSpec As IContractSpecifier, _
                ByRef pSortkeys() As ContractSortKeyIds, _
                Optional ByVal pSortDescending As Boolean = False, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "IContractStore_FetchContractsSorted"
On Error GoTo Err

Set IContractStore_FetchContractsSorted = GIBEnhancedApi.FetchContractsSorted( _
                                                mContractRequester, _
                                                mContractCache, _
                                                pContractSpec, _
                                                pSortkeys, _
                                                pSortDescending, _
                                                pCookie, _
                                                False, _
                                                mClientID)

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function IContractStore_FetchOptionExpiries( _
                ByVal pUnderlyingContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                Optional ByVal pStrike As Double = 0#, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "IContractStore_FetchOptionExpiries"
On Error GoTo Err

Set IContractStore_FetchOptionExpiries = FetchOptionExpiries(mContractRequester, mContractCache, pUnderlyingContractSpecifier, pExchange, pStrike, pCookie)

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function IContractStore_FetchOptionStrikes( _
                ByVal pUnderlyingContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                Optional ByVal pExpiry As String, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "IContractStore_FetchOptionStrikes"
On Error GoTo Err

Set IContractStore_FetchOptionStrikes = FetchOptionStrikes(mContractRequester, mContractCache, pUnderlyingContractSpecifier, pExchange, pExpiry, pCookie)

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IContractStore_Finish()
Const ProcName As String = "IContractStore_Finish"
On Error GoTo Err

Dim f As IFuture
For Each f In mContractsFutures
    If f.IsPending Then f.Cancel
Next

For Each f In mExpiriesFutures
    If f.IsPending Then f.Cancel
Next

For Each f In mStrikesFutures
    If f.IsPending Then f.Cancel
Next

Set mContractRequester = Nothing

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IContractStore_Supports(ByVal pCapabilities As ContractStoreCapabilities) As Boolean
Const Capabilities As Long = ContractStoreCapabilities.ContractStoreCanQuery Or _
                            ContractStoreCapabilities.ContractStoreOptionExpiries Or _
                            ContractStoreCapabilities.ContractStoreOptionStrikes
                            
IContractStore_Supports = ((Capabilities And pCapabilities) = pCapabilities)
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

Friend Sub Initialise( _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
                ByVal pClientID As Long)
Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
mClientID = pClientID
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




