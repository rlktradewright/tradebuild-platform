VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradeBuildAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

Implements IErrorListener

'@================================================================================
' Events
'@================================================================================

Event Notification(ByRef ev As NotificationEventData)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TradeBuildAPI"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTradeBuildConfig                           As ConfigurationSection
Private mMarketDataSourcesConfig                    As ConfigurationSection

Private mMarketDataSourceOptions                    As MarketDataSourceOptions
Private mNumberOfMarketDepthRows                    As Long

Private mDefaultStateChangeListener                 As IStateChangeListener

Private mOrderManager                               As New OrderManager

Private mServiceProviderError                       As ServiceProviderError

Private mServiceProviders                           As ServiceProviders

Private mStudyLibraryManager                        As StudyLibraryManager

Private mBarFormatterLibManager                     As BarFormatterLibManager

Private mWorkspaces                                 As WorkSpaces

Private mDefaultWorkSpace                           As WorkSpace
Private mTickers                                    As Tickers

Private WithEvents mClosePositionsTC                As TaskController
Attribute mClosePositionsTC.VB_VarHelpID = -1

Private mSessionName                                As String
Private mScopeName                                  As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
UnhandledErrorHandler.AddErrorListener Me

Set mServiceProviders = New ServiceProviders
mServiceProviders.TradeBuildAPI = Me

Set mStudyLibraryManager = New StudyLibraryManager
Set mBarFormatterLibManager = New BarFormatterLibManager
End Sub

Private Sub Class_Terminate()
'TerminateTWUtilities
Debug.Print "TradeBuildAPI terminated"
End Sub

'@================================================================================
' IErrorListener Event Handlers
'@================================================================================

Private Sub IErrorListener_Notify(ev As ErrorEventData)
On Error Resume Next
mServiceProviders.RemoveAll
End Sub

'@================================================================================
' mClosePositionsTC Event Handlers
'@================================================================================

Private Sub mClosePositionsTC_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mClosePositionsTC_Completed"
On Error GoTo Err

Set mClosePositionsTC = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get AllOrdersSimulated() As Boolean
Const ProcName As String = "AllOrdersSimulated"
On Error GoTo Err

AllOrdersSimulated = (Not mServiceProviders.IsLiveOrderSubmitterConfigured)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BarFormatterLibManager() As BarFormatterLibManager
Set BarFormatterLibManager = mBarFormatterLibManager
End Property

Public Property Get ClosingPositions() As Boolean
Const ProcName As String = "ClosingPositions"
On Error GoTo Err

ClosingPositions = Not (mClosePositionsTC Is Nothing)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ContractStorePrimary() As IContractStore
Set ContractStorePrimary = ServiceProviders.ContractStorePrimary
End Property

Public Property Get ContractStoreSecondary() As IContractStore
Set ContractStoreSecondary = ServiceProviders.ContractStoreSecondary
End Property

Public Property Get HistoricalDataStoreInput() As IHistoricalDataStore
Set HistoricalDataStoreInput = ServiceProviders.HistoricalDataStoreInput
End Property

Public Property Get HistoricalDataStoreOutput() As IHistoricalDataStore
Set HistoricalDataStoreOutput = ServiceProviders.HistoricalDataStoreOutput
End Property

Public Property Get MarketDataManager() As IMarketDataManager
Set MarketDataManager = ServiceProviders.MarketDataManager
End Property

Public Property Get OrderSubmitterFactoryLive() As IOrderSubmitterFactory
Set OrderSubmitterFactoryLive = ServiceProviders.OrderSubmitterFactoryLive
End Property

Public Property Get OrderSubmitterFactorySimulated() As IOrderSubmitterFactory
Set OrderSubmitterFactorySimulated = ServiceProviders.OrderSubmitterFactorySimulated
End Property

Public Property Get StudyLibraryManager() As StudyLibraryManager
Set StudyLibraryManager = mStudyLibraryManager
End Property

Public Property Get SupportedInputTickfileFormats() As TickfileFormatSpecifier()
Const ProcName As String = "SupportedInputTickfileFormats"
On Error GoTo Err

SupportedInputTickfileFormats = ServiceProviders.SupportedInputTickfileFormats

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SupportedOutputTickfileFormats() As TickfileFormatSpecifier()
Const ProcName As String = "SupportedOutputTickfileFormats"
On Error GoTo Err

SupportedOutputTickfileFormats = ServiceProviders.SupportedOutputTickfileFormats

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TickfileStoreInput() As ITickfileStore
Set TickfileStoreInput = ServiceProviders.TickfileStoreInput
End Property

Public Property Get TickfileStoreOutput() As ITickfileStore
Set TickfileStoreOutput = ServiceProviders.TickfileStoreOutput
End Property

Friend Property Get DefaultStateChangeListener() As IStateChangeListener
Set DefaultStateChangeListener = mDefaultStateChangeListener
End Property

Public Property Get DefaultWorkSpace() As WorkSpace
Const ProcName As String = "DefaultWorkSpace"
On Error GoTo Err

Set DefaultWorkSpace = mDefaultWorkSpace

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get MarketDataSourceOptions() As MarketDataSourceOptions
MarketDataSourceOptions = mMarketDataSourceOptions
End Property

Public Property Get NumberOfMarketDepthRows() As Long
NumberOfMarketDepthRows = mNumberOfMarketDepthRows
End Property

Public Property Get OrderManager() As OrderManager
Set OrderManager = mOrderManager
End Property

Public Property Get ServiceProviders() As ServiceProviders
Const ProcName As String = "ServiceProviders"
On Error GoTo Err

Set ServiceProviders = mServiceProviders

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ScopeName() As String
ScopeName = mScopeName
End Property

Public Property Get SessionName() As String
SessionName = mSessionName
End Property

Public Property Get Tickers() As Tickers
Const ProcName As String = "Tickers"
On Error GoTo Err

Set Tickers = mTickers

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get WorkSpaces() As WorkSpaces
Const ProcName As String = "WorkSpaces"
On Error GoTo Err

Set WorkSpaces = mWorkspaces

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CloseAllPositions( _
                Optional ByVal pTypes As PositionTypes = PositionTypeAll, _
                Optional ByVal pFlags As ClosePositionFlags = ClosePositionCancelOrders Or ClosePositionWaitForCancel) As TaskController
Const ProcName As String = "CloseAllPositions"
On Error GoTo Err

Assert mClosePositionsTC Is Nothing, "Positions are already being closed"

Dim t As New ClosePositionsTask
t.Initialise mWorkspaces, pTypes, pFlags
Set mClosePositionsTC = StartTask(t, PriorityHigh)
Set CloseAllPositions = mClosePositionsTC

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub EndSession()
Const ProcName As String = "EndSession"
On Error GoTo Err

ServiceProviders.RemoveAll

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function FetchContracts( _
                ByVal pContractSpec As ContractSpecifier, _
                Optional ByVal pListener As IContractFetchListener, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchContracts"
On Error GoTo Err

Set FetchContracts = ContractUtils27.FetchContracts(pContractSpec, mServiceProviders.ContractStorePrimary, , mServiceProviders.ContractStoreSecondary, PriorityNormal, "LoadContracts/" & pContractSpec.ToString, pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetServiceProviderError() As ServiceProviderError
Const ProcName As String = "GetServiceProviderError"
On Error GoTo Err

Set GetServiceProviderError = mServiceProviderError
'Set mServiceProviderError = Nothing

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pScopeName As String, _
                ByVal pPermittedServiceProviderRoles As ServiceProviderRoles, _
                ByVal pDefaultStateChangeListener As IStateChangeListener, _
                ByVal pMarketDataSourceOptions As MarketDataSourceOptions, _
                ByVal pNumberOfMarketDepthRows As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Init

logMessage "Initialising TradeBuild", ProcName

mScopeName = pScopeName

mServiceProviders.PermittedServiceProviderRoles = pPermittedServiceProviderRoles
mMarketDataSourceOptions = pMarketDataSourceOptions
mNumberOfMarketDepthRows = pNumberOfMarketDepthRows
Set mDefaultStateChangeListener = pDefaultStateChangeListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseFromConfig( _
                ByVal pTradeBuildConfig As ConfigurationSection, _
                ByVal pScopeName As String, _
                ByVal pPermittedServiceProviderRoles As ServiceProviderRoles, _
                ByVal pDefaultStateChangeListener As IStateChangeListener)
Const ProcName As String = "InitialiseFromConfig"
On Error GoTo Err

AssertArgument Not pTradeBuildConfig Is Nothing, "pTradeBuildConfig is Nothing"

Set mTradeBuildConfig = pTradeBuildConfig

Init

logMessage "Initialising TradeBuild from config", ProcName, mTradeBuildConfig.Path

mScopeName = pScopeName

mServiceProviders.PermittedServiceProviderRoles = pPermittedServiceProviderRoles
Set mDefaultStateChangeListener = pDefaultStateChangeListener

Dim lInitialisedOk As Boolean
lInitialisedOk = True

Set mMarketDataSourcesConfig = setupMarketDataInfo(mTradeBuildConfig)

If Not setupServiceProviders(mTradeBuildConfig) Then
    logMessage "Failed loading service providers from configuration", ProcName
    lInitialisedOk = False
End If
If Not setupStudyLibraries(mTradeBuildConfig) Then
    logMessage "Failed loading study libraries from configuration", ProcName
    lInitialisedOk = False
End If
If Not setupBarFormatterLibraries(mTradeBuildConfig) Then
    logMessage "Failed loading bar formatter libraries from configuration", ProcName
    lInitialisedOk = False
End If

If lInitialisedOk Then
    logMessage "Loaded configuration", ProcName
Else
    Assert False, "Failed to load configuration"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseWithDefaultProviders( _
                ByVal pPermittedServiceProviderRoles As ServiceProviderRoles, _
                ByVal pScopeName As String, _
                ByVal pDefaultStateChangeListener As IStateChangeListener, _
                ByVal pTwsServer As String, _
                ByVal pTwsPort As Long, _
                ByVal pClientId As Long, _
                ByVal pConnectionRetryIntervalSecs As Long, _
                ByVal pTwsLogLevel As String, _
                ByVal pTickfilesPath As String, _
                ByVal pOrderPersistenceFilePath As String)
Const ProcName As String = "InitialiseWithDefaultProviders"
On Error GoTo Err

Init

logMessage "Initialising TradeBuild with default providers", ProcName

mScopeName = pScopeName

Set mDefaultStateChangeListener = pDefaultStateChangeListener

mServiceProviders.ConfigureDefaultServiceProviders pPermittedServiceProviderRoles, _
                                                    pTwsServer, _
                                                    pTwsPort, _
                                                    pClientId, _
                                                    pConnectionRetryIntervalSecs, _
                                                    pTwsLogLevel, _
                                                    pTickfilesPath, _
                                                    pOrderPersistenceFilePath

mStudyLibraryManager.AddStudyLibrary "CmnStudiesLib27.StudyLib", True, "Built-in"
mBarFormatterLibManager.AddBarFormatterLibrary "BarFormatters27.BarFormattersLib", True, "Built-in"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyEvent(ByRef ev As NotificationEventData)
Const ProcName As String = "NotifyEvent"
On Error GoTo Err

logMessage pMsg:="Notification", _
    pMsgQualifier:="(" & GApiNotifyCodeToString(ev.EventCode) & "): " & ev.EventMessage, _
    pProcName:=ProcName
RaiseEvent Notification(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Friend Sub NotifyFatalError( _
'                ByVal ErrorCode As Long, _
'                ByVal ErrorMessage As String, _
'                ByVal errorSource As String)
'Const ProcName As String = "NotifyFatalError"
'On Error Resume Next
'
'Dim ev As ErrorEventData
'
'logMessage pMsg:="Fatal error " & ErrorCode & vbCrLf & _
'                                ErrorMessage & vbCrLf & _
'                                "Occurred at:" & vbCrLf & _
'                                errorSource, _
'            pProcName:=ProcName, _
'            pLogLevel:=LogLevelSevere
'
'logMessage pMsg:="Killing program", _
'            pProcName:=ProcName, _
'            pLogLevel:=LogLevelSevere
'
'TerminateTWUtilities    ' to stop any outstanding tasks and timers
'
'Set ev.Source = Me
'ev.ErrorCode = ErrorCode
'ev.ErrorMessage = ErrorMessage
'ev.errorSource = errorSource
'
'' don't raise the event if the error indicates that a callee has disappeared - it
'' might be the one that was supposed to handle the event
'If (ErrorCode <> &H80010007 And ErrorCode <> &H80010012) Then RaiseEvent Error(ev)
'
'' calling END in a dll is not allowed, so we call the Win32 APIs to kill the
'' process. (Note that it is not feasible to exit back to the caller of this
'' procedure because another error of some sort will almost certainly occur.)
'
'TerminateProcess GetCurrentProcess, 1
'
'End Sub

Public Function RecoverOrders( _
                Optional ByVal pCookie As Object) As TaskController
Const ProcName As String = "RecoverOrders"
On Error GoTo Err

Assert Not ServiceProviders.MarketDataManager Is Nothing, "No realtime data service provider is configured"
Assert Not ServiceProviders.OrderPersistenceDataStore Is Nothing, "No order persistence service provider is configured"
Assert mScopeName <> "", "No scope name has been supplied"
Set RecoverOrders = mOrderManager.RecoverOrdersFromPreviousSession(mScopeName, _
                                                    ServiceProviders.OrderPersistenceDataStore, _
                                                    ServiceProviders.OrderRecoveryAgent, _
                                                    ServiceProviders.MarketDataManager, _
                                                    ServiceProviders.OrderSubmitterFactoryLive, _
                                                    pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub SetServiceProviderError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMessage As String, _
                ByVal pServiceProviderName As String)
Const ProcName As String = "SetServiceProviderError"
On Error GoTo Err

Dim ev As NotificationEventData

Set mServiceProviderError = New ServiceProviderError
mServiceProviderError.ErrorCode = pErrorCode
mServiceProviderError.Message = pErrorMessage
mServiceProviderError.ServiceProviderName = pServiceProviderName
mServiceProviderError.Timestamp = GetTimestamp

Set ev.Source = Me
ev.EventCode = ApiNotifyCodes.ApiNotifyServiceProviderError
ev.EventMessage = "A Service Provider error has occurred"
NotifyEvent ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function StartServiceProviders() As Boolean
Const ProcName As String = "StartServiceProviders"
On Error GoTo Err

StartServiceProviders = ServiceProviders.StartServiceProviders(mStudyLibraryManager, mSessionName)

If Not ServiceProviders.MarketDataManager Is Nothing Then
    If Not mMarketDataSourcesConfig Is Nothing Then MarketDataManager.LoadFromConfig mMarketDataSourcesConfig
    Set mWorkspaces = CreateWorkspaces(mScopeName, _
                        Me.MarketDataManager, _
                        mOrderManager, _
                        Me.StudyLibraryManager, _
                        Me.HistoricalDataStoreInput, _
                        Me.OrderSubmitterFactoryLive, _
                        Me.OrderSubmitterFactorySimulated, _
                        Nothing)
                        
    If Not mTradeBuildConfig Is Nothing Then mWorkspaces.LoadFromConfig mTradeBuildConfig.AddPrivateConfigurationSection(ConfigSectionWorkspaces)
    
    If mWorkspaces.Contains(DefaultWorkspaceName) Then
        Set mDefaultWorkSpace = mWorkspaces(DefaultWorkspaceName)
    Else
        Set mDefaultWorkSpace = mWorkspaces.Add(DefaultWorkspaceName)
    End If
    
    Set mTickers = DefaultWorkSpace.Tickers
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub FinishAllTickers()
Const ProcName As String = "FinishAllTickers"
On Error GoTo Err

Dim ws As WorkSpace
For Each ws In mWorkspaces
    ws.Tickers.Finish
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub Init()
Const ProcName As String = "Init"
On Error GoTo Err

mSessionName = GenerateGUIDString

Set mServiceProviders = New ServiceProviders
mServiceProviders.TradeBuildAPI = Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLogger.Log pMsg, pProcName, ModuleName, pLogLevel, pMsgQualifier
End Sub

Private Function setupMarketDataInfo(ByVal pTradeBuilfConfig As ConfigurationSection) As ConfigurationSection
Const ProcName As String = "setupMarketDataInfo"
On Error GoTo Err

Dim lMarketDataSourcesConfig As ConfigurationSection
Set lMarketDataSourcesConfig = pTradeBuilfConfig.AddConfigurationSection(ConfigSectionMarketDataSources)

Dim lOptions As MarketDataSourceOptions
If CBool(lMarketDataSourcesConfig.GetSetting(ConfigSettingNoImpliedTrades, "False")) Then lOptions = MarketDataSourceOptNoImpliedTrades
If CBool(lMarketDataSourcesConfig.GetSetting(ConfigSettingNoVolumeAdjustments, "False")) Then lOptions = lOptions Or MarketDataSourceOptNoVolumeAdjustments
If CBool(lMarketDataSourcesConfig.GetSetting(ConfigSettingUseExchangeTimezone, "True")) Then lOptions = lOptions Or MarketDataSourceOptUseExchangeTimeZone

mMarketDataSourceOptions = lOptions

mNumberOfMarketDepthRows = CLng(lMarketDataSourcesConfig.GetSetting(ConfigSettingNumberOfMarketDepthRows, "20"))

Set setupMarketDataInfo = lMarketDataSourcesConfig

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupServiceProviders( _
                ByVal pTradeBuildConfig As ConfigurationSection) As Boolean
Const ProcName As String = "setupServiceProviders"
On Error GoTo Err

On Error GoTo ConfigErr
ServiceProviders.ConfigurationSection = pTradeBuildConfig
ServiceProviders.LoadServiceProvidersFromConfiguration
setupServiceProviders = True

Exit Function

ConfigErr:
logMessage "Service provider configuration failed: " & Err.Description & vbCrLf & Err.Source, ProcName, , LogLevelSevere
setupServiceProviders = False

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupStudyLibraries( _
                ByVal pTradeBuildConfig As ConfigurationSection) As Boolean
Const ProcName As String = "setupStudyLibraries"
On Error GoTo Err

On Error GoTo ConfigErr

mStudyLibraryManager.LoadStudyLibraryConfiguration pTradeBuildConfig
setupStudyLibraries = True

Exit Function

ConfigErr:
logMessage "Study library configuration failed: " & Err.Description & vbCrLf & Err.Source, ProcName, , LogLevelSevere
setupStudyLibraries = False

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupBarFormatterLibraries( _
                ByVal pTradeBuildConfig As ConfigurationSection) As Boolean
Const ProcName As String = "setupBarFormatterLibraries"
On Error GoTo Err

On Error GoTo ConfigErr
mBarFormatterLibManager.LoadBarFormatterLibraryConfiguration pTradeBuildConfig
setupBarFormatterLibraries = True

Exit Function

ConfigErr:
logMessage "Bar formatter library configuration failed: " & Err.Description & vbCrLf & Err.Source, ProcName, , LogLevelSevere
setupBarFormatterLibraries = False

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



