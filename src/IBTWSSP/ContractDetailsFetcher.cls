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

Public Function FetchContractDetails( _
                ByVal pContractSpecifier As ContractSpecifier, _
                Optional ByVal pCookie As Variant) As TaskController
Const ProcName As String = "FetchContractDetails"
On Error GoTo Err

Dim lFetchTask As New ContractDetailsRequestTask

lFetchTask.Initialise mContractRequester, CreateContractsBuilder(pContractSpecifier)
Set FetchContractDetails = StartTask(lFetchTask, PriorityLow, , pCookie)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub Initialise( _
                ByVal pContractRequester As ContractDetailsRequester)
Set mContractRequester = pContractRequester
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




