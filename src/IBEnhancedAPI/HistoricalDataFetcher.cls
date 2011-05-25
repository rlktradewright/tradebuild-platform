VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistoricalDataFetcher"
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

Private Const ModuleName                            As String = "HistoricalDataFetcher"

'@================================================================================
' Member variables
'@================================================================================

Private mHistDataRequester                          As HistDataRequester
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

Public Function FetchBars( _
                ByVal pBarSpecifier As BarDataSpecifier, _
                Optional ByVal pCookie As Variant) As TaskController
Const ProcName As String = "FetchBars"
On Error GoTo Err

Dim lHistDataFetcher As New HistDataRequestTask

lHistDataFetcher.Initialise pBarSpecifier, mHistDataRequester, mContractRequester
Set FetchBars = StartTask(lHistDataFetcher, PriorityLow, , pCookie)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub Initialise( _
                ByVal pHistDataRequester As HistDataRequester, _
                ByVal pContractRequester As ContractDetailsRequester)
Set mHistDataRequester = pHistDataRequester
Set mContractRequester = pContractRequester
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




