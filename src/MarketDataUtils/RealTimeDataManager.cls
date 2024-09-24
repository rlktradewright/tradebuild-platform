VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RealTimeDataManager"
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

Private Const ModuleName                            As String = "RealtimeDataManager"

Private Const MyCapabilities                        As Long = MarketDataManagerCapabilities.MarketDataManagerStreaming + _
                                                                MarketDataManagerCapabilities.MarketDataManagerSnapshots + _
                                                                MarketDataManagerCapabilities.MarketDataManagerProvidesDelayedData

'@================================================================================
' Member variables
'@================================================================================

Private mDataManager                                As MarketDataManager

Private mConfig                                     As ConfigurationSection

Private mDefaultStateChangeListener                 As IStateChangeListener

Private mDataDelaySeconds                           As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IMarketDataManager Interface Members
'@================================================================================

Private Property Get IMarketDataManager_Capabilities() As MarketDataManagerCapabilities
IMarketDataManager_Capabilities = MyCapabilities
End Property

Private Function IMarketDataManager_CreateMarketDataSource( _
                ByVal pContractFuture As IFuture, _
                ByVal pIsRecoverable As Boolean, _
                Optional ByVal pKey As String, _
                Optional ByVal pUseLocalTimeForStudies As Boolean, _
                Optional ByVal pStateChangeListener As IStateChangeListener, _
                Optional ByVal pOffsetExpiry As String) As IMarketDataSource
Const ProcName As String = "IMarketDataManager_CreateMarketDataSource"
On Error GoTo Err

Set IMarketDataManager_CreateMarketDataSource = CreateMarketDataSource( _
                                                    pContractFuture, _
                                                    pIsRecoverable, _
                                                    pKey, _
                                                    pUseLocalTimeForStudies, _
                                                    pStateChangeListener, _
                                                    pOffsetExpiry)

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function IMarketDataManager_CreateSnapshotMarketDataSource( _
                ByVal pContractFuture As IFuture, _
                ByVal pIsRecoverable As Boolean, _
                Optional ByVal pKey As String, _
                Optional ByVal pUseLocalTimeForStudies As Boolean, _
                Optional ByVal pStateChangeListener As IStateChangeListener, _
                Optional ByVal pRefreshInterval As Long = 0&, _
                Optional ByVal pOffsetExpiry As String) As IMarketDataSource
Const ProcName As String = "IMarketDataManager10_CreateSnapshotMarketDataSource"
On Error GoTo Err

Set IMarketDataManager_CreateSnapshotMarketDataSource = CreateSnapshotMarketDataSource( _
                                                    pContractFuture, _
                                                    pIsRecoverable, _
                                                    pKey, _
                                                    pUseLocalTimeForStudies, _
                                                    pStateChangeListener, _
                                                    pRefreshInterval, _
                                                    pOffsetExpiry)

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Private Property Let IMarketDataManager_DataDelaySeconds(ByVal Value As Long)
mDataDelaySeconds = Value
End Property

Private Property Get IMarketDataManager_DataDelaySeconds() As Long
IMarketDataManager_DataDelaySeconds = mDataDelaySeconds
End Property

Private Property Get IMarketDataManager_DataSources() As EnumerableCollection
Set IMarketDataManager_DataSources = DataSources
End Property

Private Sub IMarketDataManager_Finish()
Const ProcName As String = "IMarketDataManager_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IMarketDataManager_GetMarketDataSource(ByVal pKey As String) As IMarketDataSource
Const ProcName As String = "IMarketDataManager_GetMarketDataSource"
On Error GoTo Err

Set IMarketDataManager_GetMarketDataSource = GetMarketDataSource(pKey)

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IMarketDataManager_LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)
Const ProcName As String = "IMarketDataManager_LoadFromConfig"
On Error GoTo Err

LoadFromConfig pConfig

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataManager_RemoveFromConfig()
Const ProcName As String = "IMarketDataManager_RemoveFromConfig"
On Error GoTo Err

RemoveFromConfig

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IMarketDataManager_Supports(ByVal pCapability As MarketDataManagerCapabilities) As Boolean
IMarketDataManager_Supports = (MyCapabilities And pCapability) = pCapability
End Function

'@================================================================================
' Properties
'@================================================================================

Public Property Get DataSources() As EnumerableCollection
Set DataSources = mDataManager.DataSources
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateMarketDataSource( _
                ByVal pContractFuture As IFuture, _
                ByVal pIsRecoverable As Boolean, _
                Optional ByVal pKey As String, _
                Optional ByVal pUseLocalTimeForStudies As Boolean, _
                Optional ByVal pStateChangeListener As IStateChangeListener, _
                Optional ByVal pOffsetExpiry As String) As IMarketDataSource
Const ProcName As String = "CreateMarketDataSource"
On Error GoTo Err

If pStateChangeListener Is Nothing Then Set pStateChangeListener = mDefaultStateChangeListener

Set CreateMarketDataSource = mDataManager.CreateMarketDataSource( _
                                                pContractFuture, _
                                                CreateClockFuture(pContractFuture), _
                                                pKey, _
                                                pStateChangeListener, _
                                                pIsRecoverable, _
                                                pUseLocalTimeForStudies, _
                                                NullIndex, _
                                                False, _
                                                0, _
                                                pOffsetExpiry)

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateSnapshotMarketDataSource( _
                ByVal pContractFuture As IFuture, _
                ByVal pIsRecoverable As Boolean, _
                Optional ByVal pKey As String, _
                Optional ByVal pUseLocalTimeForStudies As Boolean, _
                Optional ByVal pStateChangeListener As IStateChangeListener, _
                Optional ByVal pRefreshInterval As Long = 0&, _
                Optional ByVal pOffsetExpiry As String) As IMarketDataSource
Const ProcName As String = "CreateSnapshotMarketDataSource"
On Error GoTo Err

If pStateChangeListener Is Nothing Then Set pStateChangeListener = mDefaultStateChangeListener

Set CreateSnapshotMarketDataSource = mDataManager.CreateMarketDataSource( _
                                                pContractFuture, _
                                                CreateClockFuture(pContractFuture), _
                                                pKey, _
                                                pStateChangeListener, _
                                                pIsRecoverable, _
                                                pUseLocalTimeForStudies, _
                                                NullIndex, _
                                                True, _
                                                pRefreshInterval, _
                                                pOffsetExpiry)

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

mDataManager.Finish

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetMarketDataSource(ByVal pKey As String) As IMarketDataSource
Const ProcName As String = "GetMarketDataSource"
On Error GoTo Err

Set GetMarketDataSource = mDataManager.GetMarketDataSourceByKey(pKey)

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pFactory As IMarketDataFactory, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore, _
                ByVal pStudyLibManager As StudyLibraryManager, _
                ByVal pOptions As MarketDataSourceOptions, _
                ByVal pDefaultStateChangeListener As IStateChangeListener, _
                ByVal pNumberOfMarketDepthRows As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mDefaultStateChangeListener = pDefaultStateChangeListener

Set mDataManager = New MarketDataManager
mDataManager.Initialise pFactory, _
                        pPrimaryContractStore, _
                        pSecondaryContractStore, _
                        pStudyLibManager, _
                        pOptions, _
                        pNumberOfMarketDepthRows

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = pConfig
mDataManager.LoadFromConfig pConfig

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

mDataManager.RemoveFromConfig

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================






