VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockContractStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Private Const ModuleName                            As String = "MockContractStore"

'@================================================================================
' Member variables
'@================================================================================

Private mContracts                                  As IContracts

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Dim lContractsBuilder As New ContractsBuilder

lContractsBuilder.Add gCreateStockContract("MSFT")

lContractsBuilder.Add gCreateOptionContract("MSFT", _
                                        "MSFT  200731C00187500", _
                                        "SMART", _
                                        "20200731", _
                                        OptCall, _
                                        187.5)

lContractsBuilder.Add gCreateOptionContract("MSFT", _
                                        "MSFT  200731C00190000", _
                                        "SMART", _
                                        "20200731", _
                                        OptCall, _
                                        190#)

lContractsBuilder.Add gCreateOptionContract("MSFT", _
                                        "MSFT  200731C00192500", _
                                        "SMART", _
                                        "20200731", _
                                        OptCall, _
                                        192.5)

lContractsBuilder.Add gCreateOptionContract("MSFT", _
                                        "MSFT  200731C00195000", _
                                        "SMART", _
                                        "20200731", _
                                        OptCall, _
                                        195#)

lContractsBuilder.Add gCreateOptionContract("MSFT", _
                                        "MSFT  200731C00197500", _
                                        "SMART", _
                                        "20200731", _
                                        OptCall, _
                                        197.5)

lContractsBuilder.Add gCreateOptionContract("MSFT", _
                                        "MSFT  200731C00200000", _
                                        "SMART", _
                                        "20200731", _
                                        OptCall, _
                                        200#)

lContractsBuilder.Add gCreateOptionContract("MSFT", _
                                        "MSFT  200731C00202500", _
                                        "SMART", _
                                        "20200731", _
                                        OptCall, _
                                        202.5)

lContractsBuilder.Add gCreateOptionContract("MSFT", _
                                        "MSFT  200731C00205000", _
                                        "SMART", _
                                        "20200731", _
                                        OptCall, _
                                        205#)

lContractsBuilder.Add gCreateOptionContract("MSFT", _
                                        "MSFT  200731C00207500", _
                                        "SMART", _
                                        "20200731", _
                                        OptCall, _
                                        207.5)

lContractsBuilder.Add gCreateOptionContract("MSFT", _
                                        "MSFT  200731C00210000", _
                                        "SMART", _
                                        "20200731", _
                                        OptCall, _
                                        210#)

lContractsBuilder.Add gCreateOptionContract("MSFT", _
                                        "MSFT  200731C00212500", _
                                        "SMART", _
                                        "20200731", _
                                        OptCall, _
                                        212.5)

lContractsBuilder.Add gCreateOptionContract("MSFT", _
                                        "MSFT  200731C00215000", _
                                        "SMART", _
                                        "20200731", _
                                        OptCall, _
                                        215#)

Set mContracts = lContractsBuilder.Contracts
End Sub

'@================================================================================
' IContractStore Interface Members
'@================================================================================

Private Function IContractStore_FetchContracts(ByVal pContractSpecifier As IContractSpecifier, Optional ByVal pListener As IContractFetchListener, Optional ByVal pCookie As Variant) As IFuture
Dim lContractFetchTask As MockContractFetchTask
Set lContractFetchTask = New MockContractFetchTask
lContractFetchTask.Initialise pContractSpecifier, mContracts, pListener
Set IContractStore_FetchContracts = lContractFetchTask.ContractsFuture
StartTask lContractFetchTask, PriorityNormal, , , pCookie
End Function

Private Function IContractStore_FetchOptionExpiries(ByVal pUnderlyingContractSpecifier As IContractSpecifier, ByVal pExchange As String, Optional ByVal pStrike As Double = 0#, Optional ByVal pCookie As Variant) As IFuture
Set IContractStore_FetchOptionExpiries = gContract.gFetchOptionExpiries( _
                                                        pUnderlyingContractSpecifier, _
                                                        pExchange, _
                                                        pStrike, _
                                                        pCookie)
End Function

Private Function IContractStore_FetchOptionStrikes(ByVal pUnderlyingContractSpecifier As IContractSpecifier, ByVal pExchange As String, Optional ByVal pExpiry As String, Optional ByVal pCookie As Variant) As IFuture
Set IContractStore_FetchOptionStrikes = gContract.gFetchOptionStrikes( _
                                                        pUnderlyingContractSpecifier, _
                                                        pExchange, _
                                                        pExpiry, _
                                                        pCookie)
End Function

Private Sub IContractStore_Finish()
Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Not supported"
End Sub

Private Function IContractStore_Supports(ByVal pCapabilities As ContractStoreCapabilities) As Boolean
IContractStore_Supports = True
End Function

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




