VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickerUtils"
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

Private Const ModuleName                            As String = "TickerUtils"

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

Public Function CreateTickers( _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pStudyLibraryManager As StudyLibraryManager, _
                Optional ByVal pHistDataStore As IHistoricalDataStore, _
                Optional ByVal pOrderManager As OrderManager, _
                Optional ByVal pOrderSubmitterFactoryLive As IOrderSubmitterFactory, _
                Optional ByVal pOrderSubmitterFactorySimulated As IOrderSubmitterFactory, _
                Optional ByVal pName As String, _
                Optional ByVal pScopeName As String, _
                Optional ByVal pMoneyManager As IMoneyManager) As Tickers
Const ProcName As String = "CreateTickers"
On Error GoTo Err

AssertArgument Not pMarketDataManager Is Nothing, "pMarketDataManager is Nothing"
AssertArgument Not pStudyLibraryManager Is Nothing, "pStudyLibraryManager is Nothing"

If pName = "" Then pName = GenerateGUIDString
Set CreateTickers = New Tickers
CreateTickers.Initialise pMarketDataManager, pStudyLibraryManager, pHistDataStore, pOrderManager, pOrderSubmitterFactoryLive, pOrderSubmitterFactorySimulated, pName, pScopeName, pMoneyManager

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




