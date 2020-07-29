VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockMarketDataManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

Implements IMarketDataManager

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

Private Const ModuleName                            As String = "MockMarketDataManager"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IMarketDataManager Interface Members
'@================================================================================

Private Property Get IMarketDataManager_Capabilities() As MarketDataManagerCapabilities

End Property

Private Function IMarketDataManager_CreateMarketDataSource( _
                ByVal pContractFuture As IFuture, _
                ByVal pIsRecoverable As Boolean, _
                Optional ByVal pKey As String, _
                Optional ByVal pUseLocalTimeForStudies As Boolean, _
                Optional ByVal pStateChangeListener As IStateChangeListener) As IMarketDataSource

End Function

Private Function IMarketDataManager_CreateSnapshotMarketDataSource( _
                ByVal pContractFuture As IFuture, _
                ByVal pIsRecoverable As Boolean, _
                Optional ByVal pKey As String, _
                Optional ByVal pUseLocalTimeForStudies As Boolean, _
                Optional ByVal pStateChangeListener As IStateChangeListener, _
                Optional ByVal pRefreshInterval As Long = 0&) As IMarketDataSource

End Function

Private Property Get IMarketDataManager_DataSources() As EnumerableCollection

End Property

Private Sub IMarketDataManager_Finish()

End Sub

Private Function IMarketDataManager_GetMarketDataSource( _
                ByVal pKey As String) As IMarketDataSource

End Function

Private Sub IMarketDataManager_LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)

End Sub

Private Sub IMarketDataManager_RemoveFromConfig()

End Sub

Private Function IMarketDataManager_Supports( _
                ByVal pCapability As MarketDataManagerCapabilities) As Boolean

End Function

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




