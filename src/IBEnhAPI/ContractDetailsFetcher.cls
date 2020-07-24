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
Private mContractCache                              As ContractCache
Private mContractsFutures                           As New EnumerableCollection
Private mStrikesFutures                             As New EnumerableCollection
Private mExpiriesFutures                            As New EnumerableCollection

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

Private Function IContractStore_FetchOptionExpiries( _
                ByVal pUnderlyingContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                Optional ByVal pStrike As Double = 0#, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "IContractStore_FetchOptionExpiries"
On Error GoTo Err

Set IContractStore_FetchOptionExpiries = FetchOptionExpiries(pUnderlyingContractSpecifier, pExchange, pStrike, pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IContractStore_FetchOptionStrikes( _
                ByVal pUnderlyingContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                Optional ByVal pExpiry As String, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "IContractStore_FetchOptionStrikes"
On Error GoTo Err

Set IContractStore_FetchOptionStrikes = FetchOptionStrikes(pUnderlyingContractSpecifier, pExchange, pExpiry, pCookie)

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

Public Function FetchOptionExpiries( _
                ByVal pUnderlyingContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                Optional ByVal pStrike As Double = 0#, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchOptionExpiries"
On Error GoTo Err

Set FetchOptionExpiries = gFetchOptionExpiries(mContractRequester, mContractCache, pUnderlyingContractSpecifier, pExchange, pStrike, pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchOptionStrikes( _
                ByVal pUnderlyingContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                Optional ByVal pExpiry As String, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchOptionStrikes"
On Error GoTo Err

Set FetchOptionStrikes = gFetchOptionStrikes(mContractRequester, mContractCache, pUnderlyingContractSpecifier, pExchange, pExpiry, pCookie)

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

For Each f In mExpiriesFutures
    If f.IsPending Then f.Cancel
Next

For Each f In mStrikesFutures
    If f.IsPending Then f.Cancel
Next

Set mContractRequester = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractCache As ContractCache)
Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
End Sub

Public Function Supports(ByVal pCapabilities As ContractStoreCapabilities) As Boolean
Const Capabilities As Long = ContractStoreCapabilities.ContractStoreCanQuery Or _
                            ContractStoreCapabilities.ContractStoreOptionExpiries Or _
                            ContractStoreCapabilities.ContractStoreOptionStrikes
                            
Supports = ((Capabilities And pCapabilities) = pCapabilities)
End Function

'@================================================================================
' Helper Functions
'@================================================================================




