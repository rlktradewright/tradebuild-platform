VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IContractStore"
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

Private Const ModuleName                            As String = "IContractStore"

'@================================================================================
' Member variables
'@================================================================================

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

Public Function FetchContracts( _
                ByVal pContractSpecifier As IContractSpecifier, _
                Optional ByVal pListener As IContractFetchListener, _
                Optional ByVal pCookie As Variant) As IFuture
End Function

Public Function FetchContractsSorted( _
                ByVal pContractSpec As IContractSpecifier, _
                ByRef pSortkeys() As ContractSortKeyIds, _
                Optional ByVal pSortDescending As Boolean = False, _
                Optional ByVal pCookie As Variant) As IFuture
End Function

Public Function FetchOptionExpiries( _
                ByVal pUnderlyingContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                Optional ByVal pStrike As Double = 0#, _
                Optional ByVal pCookie As Variant) As IFuture
End Function

Public Function FetchOptionStrikes( _
                ByVal pUnderlyingContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                Optional ByVal pExpiry As String, _
                Optional ByVal pCookie As Variant) As IFuture
End Function

Public Sub Finish()
End Sub

Public Function Supports(ByVal pCapabilities As ContractStoreCapabilities) As Boolean
End Function

'@================================================================================
' Helper Functions
'@================================================================================




