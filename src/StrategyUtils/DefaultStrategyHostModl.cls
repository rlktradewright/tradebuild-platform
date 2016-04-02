VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultStrategyHostModl"
Attribute VB_GlobalNameSpace = False
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

Implements IStrategyHostModel

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

Private Const ModuleName                            As String = "DefaultStrategyHostModl"

'@================================================================================
' Member variables
'@================================================================================

Private mContract As IContract
Private mContractStorePrimary As IContractStore
Private mContractStoreSecondary As IContractStore
Private mHistoricalDataStoreInput As IHistoricalDataStore
Private mIsTickReplay As Boolean
Private mLogDummyProfitProfile As Boolean
Private mLogParameters As Boolean
Private mLogProfitProfile As Boolean
Private mOrderSubmitterFactoryLive As IOrderSubmitterFactory
Private mOrderSubmitterFactorySimulated As IOrderSubmitterFactory
Private mRealtimeTickers As Tickers
Private mResultsPath As String
Private mSeparateSessions As Boolean
Private mShowChart As Boolean
Private mStopStrategyFactoryClassName As String
Private mStrategyClassName As String
Private mStudyLibraryManager As StudyLibraryManager
Private mSymbol As IContractSpecifier
Private mTicker As Ticker
Private mTickFileSpecifiers As TickFileSpecifiers
Private mTickfileStoreInput As ITickfileStore
Private mUseLiveBroker As Boolean
Private mUseMoneyManagement As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IStrategyHostModel Interface Members
'@================================================================================

Private Property Let IStrategyHostModel_Contract(ByVal Value As IContract)
Set mContract = Value
End Property

Private Property Get IStrategyHostModel_Contract() As IContract
Set IStrategyHostModel_Contract = mContract
End Property

Private Property Let IStrategyHostModel_ContractStorePrimary(ByVal Value As IContractStore)
Set mContractStorePrimary = Value
End Property

Private Property Get IStrategyHostModel_ContractStorePrimary() As IContractStore
Set IStrategyHostModel_ContractStorePrimary = mContractStorePrimary
End Property

Private Property Let IStrategyHostModel_ContractStoreSecondary(ByVal Value As IContractStore)
Set mContractStoreSecondary = Value
End Property

Private Property Get IStrategyHostModel_ContractStoreSecondary() As IContractStore
Set IStrategyHostModel_ContractStoreSecondary = mContractStoreSecondary
End Property

Private Property Let IStrategyHostModel_HistoricalDataStoreInput(ByVal Value As IHistoricalDataStore)
Set mHistoricalDataStoreInput = Value
End Property

Private Property Get IStrategyHostModel_HistoricalDataStoreInput() As IHistoricalDataStore
Set IStrategyHostModel_HistoricalDataStoreInput = mHistoricalDataStoreInput
End Property

Private Property Let IStrategyHostModel_IsTickReplay(ByVal Value As Boolean)
mIsTickReplay = Value
End Property

Private Property Get IStrategyHostModel_IsTickReplay() As Boolean
IStrategyHostModel_IsTickReplay = mIsTickReplay
End Property

Private Property Let IStrategyHostModel_LogDummyProfitProfile(ByVal Value As Boolean)
mLogDummyProfitProfile = Value
End Property

Private Property Get IStrategyHostModel_LogDummyProfitProfile() As Boolean
IStrategyHostModel_LogDummyProfitProfile = mLogDummyProfitProfile
End Property

Private Property Let IStrategyHostModel_LogParameters(ByVal Value As Boolean)
mLogParameters = Value
End Property

Private Property Get IStrategyHostModel_LogParameters() As Boolean
IStrategyHostModel_LogParameters = mLogParameters
End Property

Private Property Let IStrategyHostModel_LogProfitProfile(ByVal Value As Boolean)
mLogProfitProfile = Value
End Property

Private Property Get IStrategyHostModel_LogProfitProfile() As Boolean
IStrategyHostModel_LogProfitProfile = mLogProfitProfile
End Property

Private Property Let IStrategyHostModel_OrderSubmitterFactoryLive(ByVal Value As IOrderSubmitterFactory)
Set mOrderSubmitterFactoryLive = Value
End Property

Private Property Get IStrategyHostModel_OrderSubmitterFactoryLive() As IOrderSubmitterFactory
Set IStrategyHostModel_OrderSubmitterFactoryLive = mOrderSubmitterFactoryLive
End Property

Private Property Let IStrategyHostModel_OrderSubmitterFactorySimulated(ByVal Value As IOrderSubmitterFactory)
Set mOrderSubmitterFactorySimulated = Value
End Property

Private Property Get IStrategyHostModel_OrderSubmitterFactorySimulated() As IOrderSubmitterFactory
Set IStrategyHostModel_OrderSubmitterFactorySimulated = mOrderSubmitterFactorySimulated
End Property

Private Property Let IStrategyHostModel_RealtimeTickers(ByVal Value As Tickers)
Set mRealtimeTickers = Value
End Property

Private Property Get IStrategyHostModel_RealtimeTickers() As Tickers
Set IStrategyHostModel_RealtimeTickers = mRealtimeTickers
End Property

Private Property Let IStrategyHostModel_ResultsPath(ByVal Value As String)
mResultsPath = Value
End Property

Private Property Get IStrategyHostModel_ResultsPath() As String
IStrategyHostModel_ResultsPath = mResultsPath
End Property

Private Property Let IStrategyHostModel_SeparateSessions(ByVal Value As Boolean)
mSeparateSessions = Value
End Property

Private Property Get IStrategyHostModel_SeparateSessions() As Boolean
IStrategyHostModel_ShowChart = mSeparateSessions
End Property

Private Property Let IStrategyHostModel_ShowChart(ByVal Value As Boolean)
mShowChart = Value
End Property

Private Property Get IStrategyHostModel_ShowChart() As Boolean
IStrategyHostModel_ShowChart = mShowChart
End Property

Private Property Let IStrategyHostModel_StopStrategyFactoryClassName(ByVal Value As String)
mStopStrategyFactoryClassName = Value
End Property

Private Property Get IStrategyHostModel_StopStrategyFactoryClassName() As String
IStrategyHostModel_StopStrategyFactoryClassName = mStopStrategyFactoryClassName
End Property

Private Property Let IStrategyHostModel_StrategyClassName(ByVal Value As String)
mStrategyClassName = Value
End Property

Private Property Get IStrategyHostModel_StrategyClassName() As String
IStrategyHostModel_StrategyClassName = mStrategyClassName
End Property

Private Property Let IStrategyHostModel_StudyLibraryManager(ByVal Value As StudyLibraryManager)
Set mStudyLibraryManager = Value
End Property

Private Property Get IStrategyHostModel_StudyLibraryManager() As StudyLibraryManager
Set IStrategyHostModel_StudyLibraryManager = mStudyLibraryManager
End Property

Private Property Let IStrategyHostModel_Symbol(ByVal Value As IContractSpecifier)
Set mSymbol = Value
End Property

Private Property Get IStrategyHostModel_Symbol() As IContractSpecifier
Set IStrategyHostModel_Symbol = mSymbol
End Property

Private Property Let IStrategyHostModel_Ticker(ByVal Value As Ticker)
Set mTicker = Value
End Property

Private Property Get IStrategyHostModel_Ticker() As Ticker
Set IStrategyHostModel_Ticker = mTicker
End Property

Private Property Let IStrategyHostModel_TickFileSpecifiers(ByVal Value As TickFileSpecifiers)
Set mTickFileSpecifiers = Value
End Property

Private Property Get IStrategyHostModel_TickFileSpecifiers() As TickFileSpecifiers
Set IStrategyHostModel_TickFileSpecifiers = mTickFileSpecifiers
End Property

Private Property Let IStrategyHostModel_TickfileStoreInput(ByVal Value As ITickfileStore)
Set mTickfileStoreInput = Value
End Property

Private Property Get IStrategyHostModel_TickfileStoreInput() As ITickfileStore
Set IStrategyHostModel_TickfileStoreInput = mTickfileStoreInput
End Property

Private Property Let IStrategyHostModel_UseLiveBroker(ByVal Value As Boolean)
mUseLiveBroker = Value
End Property

Private Property Get IStrategyHostModel_UseLiveBroker() As Boolean
IStrategyHostModel_UseLiveBroker = mUseLiveBroker
End Property

Private Property Let IStrategyHostModel_UseMoneyManagement(ByVal Value As Boolean)
mUseMoneyManagement = Value
End Property

Private Property Get IStrategyHostModel_UseMoneyManagement() As Boolean
IStrategyHostModel_UseMoneyManagement = mUseMoneyManagement
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




