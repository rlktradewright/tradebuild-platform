VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IStrategyHostModel"
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

Private Const ModuleName                            As String = "IStrategyHostModel"

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

Public Property Let Contract(ByVal Value As IContract)
End Property

Public Property Get Contract() As IContract
End Property

Public Property Let ContractStorePrimary(ByVal Value As IContractStore)
End Property

Public Property Get ContractStorePrimary() As IContractStore
End Property

Public Property Let ContractStoreSecondary(ByVal Value As IContractStore)
End Property

Public Property Get ContractStoreSecondary() As IContractStore
End Property

Public Property Let HistoricalDataStoreInput(ByVal Value As IHistoricalDataStore)
End Property

Public Property Get HistoricalDataStoreInput() As IHistoricalDataStore
End Property

Public Property Let IsTickReplay(ByVal Value As Boolean)
End Property

Public Property Get IsTickReplay() As Boolean
End Property

Public Property Let LogDummyProfitProfile(ByVal Value As Boolean)
End Property

Public Property Get LogDummyProfitProfile() As Boolean
End Property

Public Property Let LogParameters(ByVal Value As Boolean)
End Property

Public Property Get LogParameters() As Boolean
End Property

Public Property Let LogProfitProfile(ByVal Value As Boolean)
End Property

Public Property Get LogProfitProfile() As Boolean
End Property

Public Property Let OrderSubmitterFactoryLive(ByVal Value As IOrderSubmitterFactory)
End Property

Public Property Get OrderSubmitterFactoryLive() As IOrderSubmitterFactory
End Property

Public Property Let OrderSubmitterFactorySimulated(ByVal Value As IOrderSubmitterFactory)
End Property

Public Property Get OrderSubmitterFactorySimulated() As IOrderSubmitterFactory
End Property

Public Property Let RealtimeTickers(ByVal Value As Tickers)
End Property

Public Property Get RealtimeTickers() As Tickers
End Property

Public Property Let ResultsPath(ByVal Value As String)
End Property

Public Property Get ResultsPath() As String
End Property

Public Property Let SeparateSessions(ByVal Value As Boolean)
End Property

Public Property Get SeparateSessions() As Boolean
End Property

Public Property Let ShowChart(ByVal Value As Boolean)
End Property

Public Property Get ShowChart() As Boolean
End Property

Public Property Let StopStrategyFactoryClassName(ByVal Value As String)
End Property

Public Property Get StopStrategyFactoryClassName() As String
End Property

Public Property Let StrategyClassName(ByVal Value As String)
End Property

Public Property Get StrategyClassName() As String
End Property

Public Property Let StudyLibraryManager(ByVal Value As StudyLibraryManager)
End Property

Public Property Get StudyLibraryManager() As StudyLibraryManager
End Property

Public Property Let Symbol(ByVal Value As IContractSpecifier)
End Property

Public Property Get Symbol() As IContractSpecifier
End Property

Public Property Let Ticker(ByVal Value As Ticker)
End Property

Public Property Get Ticker() As Ticker
End Property

Public Property Let TickFileSpecifiers(ByVal Value As TickFileSpecifiers)
End Property

Public Property Get TickFileSpecifiers() As TickFileSpecifiers
End Property

Public Property Let TickfileStoreInput(ByVal Value As ITickfileStore)
End Property

Public Property Get TickfileStoreInput() As ITickfileStore
End Property

Public Property Let UseLiveBroker(ByVal Value As Boolean)
End Property

Public Property Get UseLiveBroker() As Boolean
End Property

Public Property Let UseMoneyManagement(ByVal Value As Boolean)
End Property

Public Property Get UseMoneyManagement() As Boolean
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




