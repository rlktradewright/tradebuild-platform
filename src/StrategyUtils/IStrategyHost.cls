VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IStrategyHost"
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

Private Const ModuleName                            As String = "IStrategyHost"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class public sub  Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX public sub  Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ContractStorePrimary() As IContractStore
End Property

Public Property Get ContractStoreSecondary() As IContractStore
End Property

Public Property Get HistoricalDataStoreInput() As IHistoricalDataStore
End Property

Public Property Get LogDummyProfitProfile() As Boolean
End Property

Public Property Get LogParameters() As Boolean
End Property

Public Property Get LogProfitProfile() As Boolean
End Property

Public Property Get OrderSubmitterFactoryLive() As IOrderSubmitterFactory
End Property

Public Property Get OrderSubmitterFactorySimulated() As IOrderSubmitterFactory
End Property

Public Property Get RealtimeTickers() As Tickers
End Property

Public Property Get ResultsPath() As String
End Property

Public Property Get StudyLibraryManager() As StudyLibraryManager
End Property

Public Property Get TickfileStoreInput() As ITickfileStore
End Property

Public Property Get UseMoneyManagement() As Boolean
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddTimeframe( _
                ByVal pTimeframe As Timeframe)
End Sub

Public Sub ChartStudyValue(ByVal pStudy As IStudy, _
                ByVal pValueName As String, _
                ByVal pTimeframe As Timeframe)
End Sub

Public Sub ContractInvalid(ByVal pMessage As String)
End Sub

Public Sub NotifyReplayEvent(ev As NotificationEventData)
End Sub

Public Sub NotifyReplayCompleted()
End Sub

Public Sub NotifyReplayProgress( _
                ByVal tickfileTimestamp As Date, _
                ByVal eventsPlayed As Long, _
                ByVal percentComplete As Single)
End Sub

Public Sub NotifyReplayStarted()
End Sub

Public Sub NotifyTickfileCompleted(ByVal pTickfile As TickfileSpecifier, ByVal pEventsPlayed As Long)
End Sub

Public Sub TickerCreated(ByVal pTicker As Ticker)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




