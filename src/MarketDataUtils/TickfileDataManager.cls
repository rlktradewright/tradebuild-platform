VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileDataManager"
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

Implements IMarketDataFactory
Implements IMarketDataManager
Implements IStateChangeListener

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

Private Const ModuleName                            As String = "TickfileDataManager"

Private Const MyCapabilities                        As Long = MarketDataManagerCapabilities.MarketDataManagerStreaming

'@================================================================================
' Member variables
'@================================================================================

Private mReplayController                           As ReplayController
Attribute mReplayController.VB_VarHelpID = -1

Private mDataManager                                As MarketDataManager

Private mOptions                                    As MarketDataSourceOptions
Private mNumberOfMarketDepthRows                    As Long

Private mDefaultStateChangeListener                 As IStateChangeListener

Private mPrimaryContractStore                       As IContractStore
Private mSecondaryContractStore                     As IContractStore

Private mTickStreams                                As New EnumerableCollection

Private mDataDelaySeconds                           As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IMarketDataFactory Interface Members
'@================================================================================

Private Function IMarketDataFactory_CreateMarketDataStream(ByVal pContractFuture As IFuture, ByVal pErrorListener As IMarketDataErrorListener, Optional ByVal pStreamId As Long = -1&) As ITickStream
Const ProcName As String = "IMarketDataFactory_CreateMarketDataStream"
On Error GoTo Err

GMktData.Logger.Log "Retrieving tick stream for contract future: " & GetObjectKey(pContractFuture), ProcName, ModuleName

Dim lStream As ITickStream
Set lStream = mTickStreams.Item(GetObjectKey(pContractFuture))

Assert Not lStream Is Nothing, "Specified contract future does not identify an existing tick stream"

Set IMarketDataFactory_CreateMarketDataStream = lStream


Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function IMarketDataFactory_CreateMarketDepthStream(ByVal pContractFuture As IFuture, ByVal pNumberOfRows As Long, ByVal pErrorListener As IMarketDataErrorListener, Optional ByVal pStreamId As Long = -1&) As ITickStream
Const ProcName As String = "IMarketDataFactory_CreateMarketDepthStream"
On Error GoTo Err

Dim lStream As ITickStream
Set lStream = mTickStreams.Item(GetObjectKey(pContractFuture))

Assert Not lStream Is Nothing, "Specified contract future does not identify an existing tick stream"

Set IMarketDataFactory_CreateMarketDepthStream = lStream

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function IMarketDataFactory_CreateSnapshotMarketDataStream( _
                ByVal pContractFuture As IFuture, _
                ByVal pErrorListener As IMarketDataErrorListener, _
                Optional ByVal pStreamId As Long = -1&) As ITickStream

Const ProcName As String = "IMarketDataFactory_CreateSnapshotMarketDataStream"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, "Not supported"

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IMarketDataFactory_MaxStreamId() As Long
IMarketDataFactory_MaxStreamId = &H7FFFFFFF
End Property

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

Set IMarketDataManager_CreateMarketDataSource = CreateMarketDataSource(pContractFuture, pIsRecoverable, pKey, pUseLocalTimeForStudies, pStateChangeListener)

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

Err.Raise ErrorCodes.ErrUnsupportedOperationException, "Not supported"

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

Err.Raise ErrorCodes.ErrUnsupportedOperationException, "Not supported"

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataManager_RemoveFromConfig()
Const ProcName As String = "IMarketDataManager_RemoveFromConfig"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, "Not supported"

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IMarketDataManager_Supports(ByVal pCapability As MarketDataManagerCapabilities) As Boolean
IMarketDataManager_Supports = (MyCapabilities And pCapability) = pCapability
End Function

'@================================================================================
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As TWUtilities40.StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

If ev.State <> TickStreamStateFinished Then Exit Sub

Dim lStream As TickStream
Set lStream = ev.Source

mTickStreams.Remove GetObjectKey(lStream.ContractFuture)

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get DataSources() As EnumerableCollection
Set DataSources = mDataManager.DataSources
End Property

Public Property Get ReplayController() As ReplayController
Set ReplayController = mReplayController
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
                Optional ByVal pRecoveryContractSpec As IContractSpecifier) As IMarketDataSource
Const ProcName As String = "CreateMarketDataSource"
On Error GoTo Err

AssertArgument Not pContractFuture Is Nothing, "pContractFuture cannot be Nothing"
AssertArgument Not pIsRecoverable, "Recoverable tickfile-based market data sources not supported"
AssertArgument pRecoveryContractSpec Is Nothing, "Recovery contract spec not relevant for tickfile-based market data sources"

If pStateChangeListener Is Nothing Then Set pStateChangeListener = mDefaultStateChangeListener

AssertArgument mTickStreams.Contains(GetObjectKey(pContractFuture)), "No data for this contract"

Dim lTickStream As ITickStream
Set lTickStream = mTickStreams(GetObjectKey(pContractFuture))

Dim lDataSource As New MarketDataSource
Set lDataSource = mDataManager.CreateMarketDataSource( _
                                    pContractFuture, _
                                    lTickStream.ClockFuture, _
                                    pKey, _
                                    pStateChangeListener, _
                                    False, _
                                    pUseLocalTimeForStudies, _
                                    lTickStream.StreamId, _
                                    False, _
                                    0, _
                                    "")

Set CreateMarketDataSource = lDataSource

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
                ByVal pTickfileSpecifiers As TickFileSpecifiers, _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pPlaySequentially As Boolean, _
                ByVal pStudyLibManager As StudyLibraryManager, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore, _
                ByVal pOptions As MarketDataSourceOptions, _
                ByVal pDefaultStateChangeListener As IStateChangeListener, _
                ByVal pNumberOfMarketDepthRows As Long, _
                ByVal pReplaySpeed As Long, _
                ByVal pReplayProgressEventInterval As Long, _
                ByVal pTimestampAdjustmentStart As Double, _
                ByVal pTimestampAdjustmentEnd As Double)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore

mOptions = pOptions
Set mDefaultStateChangeListener = pDefaultStateChangeListener
mNumberOfMarketDepthRows = pNumberOfMarketDepthRows

Set mDataManager = New MarketDataManager
mDataManager.Initialise Me, _
                        pPrimaryContractStore, _
                        pSecondaryContractStore, _
                        pStudyLibManager, _
                        pOptions, _
                        pNumberOfMarketDepthRows

Set mReplayController = CreateTickfileReplayController(pTickfileStore, mPrimaryContractStore, mSecondaryContractStore, pReplaySpeed, pReplayProgressEventInterval, pTimestampAdjustmentStart, pTimestampAdjustmentEnd)

Dim lTickStream As ITickStream
If pPlaySequentially Then
    mReplayController.SequentialTickFileSpecifiers = pTickfileSpecifiers
    Set lTickStream = mReplayController.TickStream(0)
    lTickStream.AddStateChangeListener Me
    GMktData.Logger.Log "Add tick stream for contract future: " & GetObjectKey(lTickStream.ContractFuture), ProcName, ModuleName
    mTickStreams.Add lTickStream, GetObjectKey(lTickStream.ContractFuture)
Else
    mReplayController.TickFileSpecifiers = pTickfileSpecifiers
    Dim i As Long
    For i = 1 To pTickfileSpecifiers.Count
        Set lTickStream = mReplayController.TickStream(i - 1)
        GMktData.Logger.Log "Add tick stream for contract future: " & GetObjectKey(lTickStream.ContractFuture), ProcName, ModuleName
        mTickStreams.Add lTickStream, GetObjectKey(lTickStream.ContractFuture)
    Next
End If

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




