VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsContractCache"
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

Private Const ModuleName                            As String = "TwsContractCache"

'@================================================================================
' Member variables
'@================================================================================

Private mTwsContracts                               As New EnumerableCollection

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

Friend Sub Add(ByVal pContractDetails As TwsContractDetails)
Const ProcName As String = "Add"
On Error GoTo Err

If pContractDetails.Summary.Sectype = TwsSecTypeCombo Then Exit Sub
If Not IsValidExchangeCode(pContractDetails.Summary.Exchange) Then Exit Sub

Dim lSpec As IContractSpecifier
With pContractDetails.Summary
    Set lSpec = CreateContractSpecifier(.LocalSymbol, .Symbol, .Exchange, gTwsSecTypeToSecType(.Sectype), .CurrencyCode, .Expiry, .Strike, gTwsOptionRightToOptionRight(.OptRight))
End With

Dim lKey As String
lKey = lSpec.Key

Dim lEntry As ContractCacheEntry
If mTwsContracts.Contains(lKey) Then
    lEntry = mTwsContracts.Item(lKey)
    If CLng(lEntry.CacheDate) = CLng(Now) Then Exit Sub
    mTwsContracts.Remove lKey
End If

lEntry.CacheDate = Now
Set lEntry.ContractDetails = pContractDetails
mTwsContracts.Add lEntry, lKey
gLogger.Log "Added contract to cache with key", ProcName, ModuleName, LogLevelDetail, lKey

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function Item(ByVal pContractSpec As IContractSpecifier) As TwsContractDetails
Const ProcName As String = "Item"
On Error GoTo Err

If pContractSpec.Sectype = SecTypeCombo Then Exit Function

Dim lKey As String
lKey = pContractSpec.Key

Dim lEntry As ContractCacheEntry
If mTwsContracts.Contains(lKey) Then
    lEntry = mTwsContracts.Item(lKey)
    If CLng(lEntry.CacheDate) <> CLng(Now) Then Exit Function
    
    Set Item = lEntry.ContractDetails
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




