VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "ContractCache"

'@================================================================================
' Member variables
'@================================================================================

Private mContracts                                  As New EnumerableCollection

Private mOptionParameters                           As New EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddContract( _
                ByVal pContract As IContract, _
                ByVal pTwsContractDetails As TwsContractDetails)
Const ProcName As String = "AddContract"
On Error GoTo Err

If pContract.Specifier.SecType = SecTypeCombo Then Exit Sub
If Not IsValidExchangeCode(pTwsContractDetails.Summary.Exchange) Then Exit Sub

Dim lEntry As ContractCacheEntry
If tryGetCacheEntry(pContract.Specifier.Key, lEntry) Then Exit Sub

lEntry.CacheDate = Now
Set lEntry.Contract = pContract
Set lEntry.TwsContractDetails = pTwsContractDetails
mContracts.Add lEntry, pContract.Specifier.Key
gLogger.Log "Added contract to cache with key", ProcName, ModuleName, LogLevelDetail, pContract.Specifier.Key

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddContractWithPartialKey( _
                ByVal pContract As IContract, _
                ByVal pTwsContractDetails As TwsContractDetails, _
                ByVal pKey As String)
Const ProcName As String = "AddContractWithPartialKey"
On Error GoTo Err

If pContract.Specifier.SecType = SecTypeCombo Then Exit Sub
If Not IsValidExchangeCode(pTwsContractDetails.Summary.Exchange) Then Exit Sub

Dim lEntry As ContractCacheEntry
If tryGetCacheEntry(pKey, lEntry) Then Exit Sub

lEntry.CacheDate = Now
Set lEntry.Contract = pContract
Set lEntry.TwsContractDetails = pTwsContractDetails
mContracts.Add lEntry, pKey
gLogger.Log "Added contract to cache with partial key", ProcName, ModuleName, LogLevelDetail, pKey

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddOptionParameter( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pExchange As String, _
                pExpirations() As String, _
                pStrikes() As Double)
Const ProcName As String = "AddOptionParameter"
On Error GoTo Err

Dim lParametersStore As OptionParametersStore
Dim lKey As String: lKey = generateOptionParametersKey(pContractSpec.Key, pExchange)
If mOptionParameters.TryItem(lKey, lParametersStore) Then
    If CLng(lParametersStore.CacheDate) <> CLng(Now) Then
        mOptionParameters.Remove lKey
        Set lParametersStore = Nothing
    End If
End If

If lParametersStore Is Nothing Then
    Set lParametersStore = New OptionParametersStore
    mOptionParameters.Add lParametersStore, lKey
End If

Dim i As Long
Dim j As Long
For i = 0 To UBound(pExpirations)
    For j = 0 To UBound(pStrikes)
        lParametersStore.AddParameter pStrikes(j), pExpirations(i)
    Next
Next

For i = 0 To UBound(pExpirations)
    lParametersStore.AddExpiry pExpirations(i)
Next

For j = 0 To UBound(pStrikes)
    lParametersStore.AddStrike pStrikes(j)
Next

lParametersStore.Complete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub EndOptionParameters( _
                ByVal pContractSpec As IContractSpecifier)
Const ProcName As String = "EndOptionParameters"
On Error GoTo Err

' the presence of this key in the collection indicates that all the
' relevant option parameters have been processed
mOptionParameters.Add Nothing, generateEndOptionParametersKey(pContractSpec)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function GetContract(ByVal pContractSpec As IContractSpecifier) As IContract
Const ProcName As String = "GetContract"
On Error GoTo Err

If pContractSpec.SecType = SecTypeCombo Then Exit Function

Dim lEntry As ContractCacheEntry
If Not tryGetCacheEntry(pContractSpec.Key, lEntry) Then Exit Function

Set GetContract = lEntry.Contract

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetTwsContractDetails(ByVal pContractSpec As IContractSpecifier) As TwsContractDetails
Const ProcName As String = "GetTwsContractDetails"
On Error GoTo Err

If pContractSpec.SecType = SecTypeCombo Then Exit Function

Dim lEntry As ContractCacheEntry
If Not tryGetCacheEntry(pContractSpec.Key, lEntry) Then Exit Function

Set GetTwsContractDetails = lEntry.TwsContractDetails

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetOptionExpiries( _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                ByVal pStrike As Double) As Expiries
Const ProcName As String = "GetOptionExpiries"
On Error GoTo Err

If Not mOptionParameters.Contains(generateEndOptionParametersKey(pContractSpecifier)) Then Exit Function

Dim lOptionParametersStore As OptionParametersStore
If mOptionParameters.TryItem( _
                generateOptionParametersKey(pContractSpecifier.Key, pExchange), _
                lOptionParametersStore) Then
    If pStrike = 0 Then
        Set GetOptionExpiries = lOptionParametersStore.GetExpiries()
    Else
        Set GetOptionExpiries = lOptionParametersStore.GetExpiriesForStrike(pStrike)
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetOptionStrikes( _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                ByVal pExpiry As String) As Strikes
Const ProcName As String = "GetOptionStrikes"
On Error GoTo Err

If Not mOptionParameters.Contains(generateEndOptionParametersKey(pContractSpecifier)) Then Exit Function

Dim lOptionParametersStore As OptionParametersStore
If mOptionParameters.TryItem( _
                generateOptionParametersKey(pContractSpecifier.Key, pExchange), _
                lOptionParametersStore) Then
    If pExpiry = "" Then
        Set GetOptionStrikes = lOptionParametersStore.GetStrikes()
    Else
        Set GetOptionStrikes = lOptionParametersStore.GetStrikesForExpiry(pExpiry)
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function generateEndOptionParametersKey( _
                ByVal pContractSpec As IContractSpecifier)
generateEndOptionParametersKey = pContractSpec.Key & "|END"
End Function

Private Function generateOptionParametersKey( _
                ByVal pContractKey As String, _
                ByVal pExchange As String)
Const ProcName As String = "generateOptionParametersKey"
On Error GoTo Err

generateOptionParametersKey = pContractKey & "|" & pExchange

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function tryGetCacheEntry( _
                ByVal pKey As String, _
                ByRef pCacheEntry As ContractCacheEntry) As Boolean
Const ProcName As String = "tryGetCacheEntry"
On Error GoTo Err

If Not mContracts.TryItem(pKey, pCacheEntry) Then
ElseIf CLng(pCacheEntry.CacheDate) <> CLng(Now) Then
    mContracts.Remove pKey
Else
    tryGetCacheEntry = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


