Attribute VB_Name = "GContractStore"
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

Public Enum OptionParameterTypes
    OptionParameterTypeNone
    OptionParameterTypeExpiries
    OptionParameterTypeStrikes
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GContractStore"

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

Public Function FetchOptionExpiries( _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
                ByVal pUnderlyingContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                Optional ByVal pStrike As Double = 0#, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchOptionExpiries"
On Error GoTo Err

If GIBEnhApi.Logger.IsLoggable(LogLevelDetail) Then GIBEnhApi.Log "Fetching option expiries for", ModuleName, ProcName, pUnderlyingContractSpecifier.ToString, LogLevelDetail
Dim lFetcher As New OptionParametersRequester
Set FetchOptionExpiries = lFetcher.Fetch( _
                                        pContractRequester, _
                                        pContractCache, _
                                        pUnderlyingContractSpecifier, _
                                        OptionParameterTypeExpiries, _
                                        pExchange, _
                                        "", _
                                        pStrike, _
                                        pCookie)

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchOptionStrikes( _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
                ByVal pUnderlyingContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                Optional ByVal pExpiry As String, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchOptionStrikes"
On Error GoTo Err

If GIBEnhApi.Logger.IsLoggable(LogLevelDetail) Then GIBEnhApi.Log "Fetching option strikes for", ModuleName, ProcName, pUnderlyingContractSpecifier.ToString, LogLevelDetail
Dim lFetcher As New OptionParametersRequester
Set FetchOptionStrikes = lFetcher.Fetch( _
                                        pContractRequester, _
                                        pContractCache, _
                                        pUnderlyingContractSpecifier, _
                                        OptionParameterTypeStrikes, _
                                        pExchange, _
                                        pExpiry, _
                                        0#, _
                                        pCookie)

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




