VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WorkspaceUtils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
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

Private Const ModuleName                            As String = "WorkspaceUtils"

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

Public Function CreateWorkspaces( _
                ByVal pScopeName As String, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pOrderManager As OrderManager, _
                ByVal pStudyLibraryManager As StudyLibraryManager, _
                ByVal pHistoricalDataStore As IHistoricalDataStore, _
                Optional ByVal pOrderSubmitterFactoryLive As IOrderSubmitterFactory, _
                Optional ByVal pOrderSubmitterFactorySimulated As IOrderSubmitterFactory, _
                Optional ByVal pMoneyManager As IMoneyManager) As WorkSpaces
Const ProcName As String = "CreateWorkspaces"
On Error GoTo Err

Set CreateWorkspaces = New WorkSpaces
CreateWorkspaces.Initialise pScopeName, pMarketDataManager, pOrderManager, pStudyLibraryManager, pHistoricalDataStore, pOrderSubmitterFactoryLive, pOrderSubmitterFactorySimulated, pMoneyManager

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




