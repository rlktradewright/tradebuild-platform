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

Implements IGenericTickListener
Implements IMarketDataErrorListener
Implements IMarketDataFactory
Implements IMarketDataManager
Implements TaskCompletionListener

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

Private Const ConfigSectionMarketDataSource         As String = "MarketDataSource"

Private Const ConfigSettingOptions                  As String = "&Options"
Private Const ConfigSettingKey                      As String = "&Key"
Private Const ConfigSettingNumberOfMarketDepthRows  As String = "&NumberOfMarketDepthRows"

'@================================================================================
' Member variables
'@================================================================================

Private mFactory                                    As IMarketDataFactory
Private mPrimaryContractStore                       As IContractStore
Private mSecondaryContractStore                     As IContractStore

Private mDataManager                                As MarketDataManager

Private mConfig                                     As ConfigurationSection

Private mDefaultStateChangeListener                 As StateChangeListener

Private mContractFuture                             As IFuture

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

mDataManager.NotifyTick ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IMarketDataErrorListener Interface Members
'@================================================================================

Private Sub IMarketDataErrorListener_NotifyMarketDataError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "IMarketDataErrorListener_NotifyMarketDataError"
On Error GoTo Err

mDataManager.NotifyMarketDataError pStreamId, pErrorCode, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataErrorListener_NotifyMarketDepthError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "IMarketDataErrorListener_NotifyMarketDepthError"
On Error GoTo Err

mDataManager.NotifyMarketDataError pStreamId, pErrorCode, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IMarketDataFactory Interface Members
'@================================================================================

Private Function IMarketDataFactory_CreateMarketDataStream(ByVal pContract As IContract, ByVal pDataListener As TickUtils27.IGenericTickListener, ByVal pErrorListener As TickUtils27.IMarketDataErrorListener, Optional ByVal pStreamId As Long) As ITickStream
Const ProcName As String = "IMarketDataFactory_CreateMarketDataStream"
On Error GoTo Err

Set IMarketDataFactory_CreateMarketDataStream = mFactory.CreateMarketDataStream(pContract, pDataListener, pErrorListener, pStreamId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IMarketDataFactory_CreateMarketDepthStream(ByVal pContract As IContract, ByVal pNumberOfRows As Long, ByVal pDataListener As TickUtils27.IGenericTickListener, ByVal pErrorListener As IMarketDataErrorListener, Optional ByVal pStreamId As Long) As ITickStream
Const ProcName As String = "IMarketDataFactory_CreateMarketDepthStream"
On Error GoTo Err

Set IMarketDataFactory_CreateMarketDepthStream = mFactory.CreateMarketDepthStream(pContract, pNumberOfRows, pDataListener, pErrorListener, pStreamId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IMarketDataFactory_MaxStreamId() As Long
IMarketDataFactory_MaxStreamId = mFactory.MaxStreamId
End Property

'@================================================================================
' IMarketDataManager Interface Members
'@================================================================================

Private Function IMarketDataManager_CreateMarketDataSource( _
                Optional ByVal pContract As IContract, _
                Optional ByVal pContractSpec As IContractSpecifier, _
                Optional ByVal pKey As String, _
                Optional ByVal pStateChangeListener As StateChangeListener) As IMarketDataSource
Const ProcName As String = "IMarketDataManager_CreateMarketDataSource"
On Error GoTo Err

Set IMarketDataManager_CreateMarketDataSource = CreateMarketDataSource(pContract, pContractSpec, pKey)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IMarketDataManager_Finish()
Const ProcName As String = "IMarketDataManager_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IMarketDataManager_GetMarketDataSource(ByVal pKey As String) As IMarketDataSource
Const ProcName As String = "IMarketDataManager_GetMarketDataSource"
On Error GoTo Err

Set IMarketDataManager_GetMarketDataSource = GetMarketDataSource(pKey)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IMarketDataManager_PrimaryContractStore() As IContractStore
Set IMarketDataManager_PrimaryContractStore = PrimaryContractStore
End Property

Private Property Get IMarketDataManager_SecondaryContractStore() As IContractStore
Set IMarketDataManager_SecondaryContractStore = SecondaryContractStore
End Property

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "TaskCompletionListener_TaskCompleted"
On Error GoTo Err

Dim lIndex As Long
lIndex = ev.Cookie

If mContractFuture.IsCancelled Then
    mDataManager.NotifyError lIndex, MarketDataErrorCodes.MarketDataErrCancelled, "The contract fetch operation was cancelled"
ElseIf mContractFuture.IsFaulted Then
    mDataManager.NotifyError lIndex, MarketDataErrorCodes.MarketDataErrContractFetchError, mContractFuture.ErrorMessage
Else
    mDataManager.SetContract lIndex, mContractFuture.Contract
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get PrimaryContractStore() As IContractStore
Set PrimaryContractStore = mPrimaryContractStore
End Property

Public Property Get SecondaryContractStore() As IContractStore
Set SecondaryContractStore = mSecondaryContractStore
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateMarketDataSource( _
                Optional ByVal pContract As IContract, _
                Optional ByVal pContractSpec As IContractSpecifier, _
                Optional ByVal pKey As String, _
                Optional ByVal pStateChangeListener As StateChangeListener) As IMarketDataSource
Const ProcName As String = "CreateMarketDataSource"
On Error GoTo Err

If Not pContractSpec Is Nothing Then AssertArgument Not (mPrimaryContractStore Is Nothing And mSecondaryContractStore Is Nothing), "No contract store is available"

If pStateChangeListener Is Nothing Then Set pStateChangeListener = mDefaultStateChangeListener

Dim lDataSource As New MarketDataSource
Set lDataSource = mDataManager.CreateMarketDataSource(pContract, pContractSpec, pKey, pStateChangeListener)

If lDataSource.Contract Is Nothing And pContract Is Nothing Then getContract pContractSpec, lDataSource.StreamId

Set CreateMarketDataSource = lDataSource

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

mDataManager.Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetMarketDataSource(ByVal pKey As String) As IMarketDataSource
Const ProcName As String = "GetMarketDataSource"
On Error GoTo Err

Set GetMarketDataSource = mDataManager.GetMarketDataSourceByKey(pKey)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pFactory As IMarketDataFactory, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore, _
                ByVal pOptions As MarketDataSourceOptions, _
                ByVal pDefaultStateChangeListener As StateChangeListener, _
                ByVal pNumberOfMarketDepthRows As Long, _
                ByVal pConfig As ConfigurationSection)
Set mFactory = pFactory
Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore
Set mDefaultStateChangeListener = pDefaultStateChangeListener
Set mConfig = pConfig

Set mDataManager = New MarketDataManager
mDataManager.Initialise Me, pOptions, pNumberOfMarketDepthRows, pConfig

If Not mConfig Is Nothing Then mDataManager.LoadFromConfig mConfig
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

mDataManager.RemoveFromConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub getContract(ByVal pContractSpec As IContractSpecifier, ByVal pindex As Long)
Const ProcName As String = "getContract"
On Error GoTo Err

Set mContractFuture = FetchContract(pContractSpec, mPrimaryContractStore, mSecondaryContractStore, , , pindex)
mContractFuture.TaskController.AddTaskCompletionListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getConfigSection( _
                Optional ByVal pContract As IContract, _
                Optional ByVal pOptions As MarketDataSourceOptions = MarketDataSourceOptUseExchangeTimeZone, _
                Optional ByVal pKey As String, _
                Optional ByVal pNumberOfMarketDepthRows As Long) As ConfigurationSection
Const ProcName As String = "getConfigSection"
On Error GoTo Err

If mConfig Is Nothing Then Exit Function


Dim lConfig As ConfigurationSection
Set lConfig = mConfig.AddConfigurationSection(ConfigSectionMarketDataSource & "(" & pKey & ")")

If Not pContract Is Nothing Then SaveContractToConfig pContract, lConfig.AddConfigurationSection(ConfigSectionContract)
lConfig.SetSetting ConfigSettingKey, pKey
lConfig.SetSetting ConfigSettingOptions, pOptions
lConfig.SetSetting ConfigSettingNumberOfMarketDepthRows, pNumberOfMarketDepthRows

Set getConfigSection = lConfig

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function





