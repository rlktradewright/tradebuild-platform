VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradeBuild"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'' A GlobalMultiUse class that provides a variety of useful functions.

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TradeBuild"

'@================================================================================
' Enums
'@================================================================================

Public Enum ApiNotifyCodes
    ApiNotifyServiceProviderError = 900
    
    ApiNotifySourceNotResponding
    ApiNotifyConnecting
    ApiNotifyConnected
    ApiNotifyCantConnect
    ApiNotifyRetryingConnection
    ApiNotifyReconnecting
    ApiNotifyLostConnection

    ApiNotifyNonSpecificNotification
End Enum

Public Enum ClosePositionFlags
    ClosePositionCancelOrders = 1
    ClosePositionWaitForCancel = 2
End Enum

Public Enum PositionTypes
    PositionTypeLive = 1
    PositionTypeSimulated = 2
    PositionTypeAll = 3
End Enum

Public Enum ServiceProviderRoles
    SPRoleRealtimeData = 1&
    SPRoleContractDataPrimary = 2&
    SPRoleContractDataSecondary = 4&
    SPRoleHistoricalDataInput = 8&
    SPRoleHistoricalDataOutput = &H10&
    SPRoleOrderSubmissionSimulated = &H20&
    SPRoleTickfileInput = &H40&
    SPRoleTickfileOutput = &H80&
    SPRoleOrderSubmissionLive = &H100&
    SPRoleOrderPersistence = &H200&
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type ClosePositionsResult
    NumberOfPositionsClosed     As Long
    NumberOfPositionsNotClosed  As Long
End Type

Public Type ServiceProviderDetails
    Comments                    As String
    Vendor                      As String
    EXEName                     As String
    FileDescription             As String
    LegalCopyright              As String
    LegalTrademarks             As String
    VersionMajor                As Long
    VersionMinor                As Long
    VersionRevision             As Long
    Path                        As String
    ProductName                 As String
End Type

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Randomize
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get SPNameContractDataPrimary() As String
SPNameContractDataPrimary = SrvcProviderNamePrimaryContractData
End Property

Public Property Get SPNameContractDataSecondary() As String
SPNameContractDataSecondary = SrvcProviderNameSecondaryContractData
End Property

Public Property Get SPNameHistoricalDataInput() As String
SPNameHistoricalDataInput = SrvcProviderNameHistoricalDataInput
End Property

Public Property Get SPNameHistoricalDataOutput() As String
SPNameHistoricalDataOutput = SrvcProviderNameHistoricalDataOutput
End Property

Public Property Get SPNameOrderSubmissionLive() As String
SPNameOrderSubmissionLive = SrvcProviderNameBrokerLive
End Property

Public Property Get SPNameOrderSubmissionSimulated() As String
SPNameOrderSubmissionSimulated = SrvcProviderNameBrokerSimulated
End Property

Public Property Get SPNameRealtimeData() As String
SPNameRealtimeData = SrvcProviderNameRealtimeData
End Property

Public Property Get SPNameTickfileInput() As String
SPNameTickfileInput = SrvcProviderNameTickfileInput
End Property

Public Property Get SPNameTickfileOutput() As String
SPNameTickfileOutput = SrvcProviderNameTickfileOutput
End Property

'Public Property Get SPProgIdQtBarData() As String
'SPProgIdQtBarData = ProgIdQtBarData
'End Property
'
'Public Property Get SPProgIdQtRealtimeData() As String
'SPProgIdQtRealtimeData = ProgIdQtRealtimeData
'End Property
'
'Public Property Get SPProgIdQtTickData() As String
'SPProgIdQtTickData = ProgIdQtTickData
'End Property

Public Property Get SPProgIdTbBarData() As String
SPProgIdTbBarData = ProgIdTbBarData
End Property

Public Property Get SPProgIdTbContractData() As String
SPProgIdTbContractData = ProgIdTbContractData
End Property

Public Property Get SPProgIdTbOrders() As String
SPProgIdTbOrders = ProgIdTbOrders
End Property

Public Property Get SPProgIdTbTickData() As String
SPProgIdTbTickData = ProgIdTbTickData
End Property

Public Property Get SPProgIdFileTickData() As String
SPProgIdFileTickData = ProgIdFileTickData
End Property

Public Property Get SPProgIdTwsBarData() As String
SPProgIdTwsBarData = ProgIdTwsBarData
End Property

Public Property Get SPProgIdTwsContractData() As String
SPProgIdTwsContractData = ProgIdTwsContractData
End Property

Public Property Get SPProgIdTwsOrders() As String
SPProgIdTwsOrders = ProgIdTwsOrders
End Property

Public Property Get SPProgIdTwsRealtimeData() As String
SPProgIdTwsRealtimeData = ProgIdTwsRealtimeData
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateTradeBuildAPI( _
                Optional ByVal pScopeName As String, _
                Optional ByVal pPermittedServiceProviderRoles As ServiceProviderRoles = _
                                SPRoleOrderSubmissionSimulated + _
                                SPRoleHistoricalDataInput + _
                                SPRoleContractDataPrimary + _
                                SPRoleRealtimeData, _
                Optional ByVal pDefaultStateChangeListener As IStateChangeListener, _
                Optional ByVal pMarketDataSourceOptions As MarketDataSourceOptions = MarketDataSourceOptions.MarketDataSourceOptUseExchangeTimeZone, _
                Optional ByVal pNumberOfMarketDepthRows As Long = 20) As TradeBuildAPI
Const ProcName As String = "CreateTradeBuildAPI"
On Error GoTo Err

Set CreateTradeBuildAPI = New TradeBuildAPI
CreateTradeBuildAPI.Initialise pScopeName, _
                                pPermittedServiceProviderRoles, _
                                pDefaultStateChangeListener, _
                                pMarketDataSourceOptions, _
                                pNumberOfMarketDepthRows

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTradeBuildAPIFromConfig( _
                ByVal pConfig As ConfigurationSection, _
                Optional ByVal pScopeName As String, _
                Optional ByVal pPermittedServiceProviderRoles As ServiceProviderRoles = _
                                SPRoleOrderSubmissionSimulated + _
                                SPRoleHistoricalDataInput + _
                                SPRoleContractDataPrimary + _
                                SPRoleRealtimeData, _
                Optional ByVal pDefaultStateChangeListener As IStateChangeListener) As TradeBuildAPI
Const ProcName As String = "CreateTradeBuildAPIFromConfig"
On Error GoTo Err

AssertArgument Not pConfig Is Nothing, "No configuration supplied"

Set CreateTradeBuildAPIFromConfig = New TradeBuildAPI
CreateTradeBuildAPIFromConfig.InitialiseFromConfig _
                pConfig.AddConfigurationSection(ConfigSectionTradeBuild), _
                pScopeName, _
                pPermittedServiceProviderRoles, _
                pDefaultStateChangeListener

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTradeBuildAPIWithDefaultProviders( _
                Optional ByVal pPermittedServiceProviderRoles As ServiceProviderRoles = _
                                SPRoleOrderSubmissionSimulated + _
                                SPRoleHistoricalDataInput + _
                                SPRoleContractDataPrimary + _
                                SPRoleRealtimeData, _
                Optional ByVal pScopeName As String, _
                Optional ByVal pDefaultStateChangeListener As IStateChangeListener, _
                Optional ByVal pTwsServer As String, _
                Optional ByVal pTwsPort As Long = 7496, _
                Optional ByVal pTwsClientId As Long = -1, _
                Optional ByVal pTwsConnectionRetryIntervalSecs As Long = 60, _
                Optional ByVal pTwsLogLevel As String = "ERROR", _
                Optional ByVal pTickfilesPath As String, _
                Optional ByVal pOrderPersistenceFilePath As String) As TradeBuildAPI
Const ProcName As String = "CreateTradeBuildAPIWithDefaultProviders"
On Error GoTo Err

Set CreateTradeBuildAPIWithDefaultProviders = New TradeBuildAPI
CreateTradeBuildAPIWithDefaultProviders.InitialiseWithDefaultProviders _
                pScopeName, _
                pPermittedServiceProviderRoles, _
                pDefaultStateChangeListener, _
                pTwsServer, _
                pTwsPort, _
                pTwsClientId, _
                pTwsConnectionRetryIntervalSecs, _
                pTwsLogLevel, _
                pTickfilesPath, _
                pOrderPersistenceFilePath

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


Public Sub SetDefaultTradeBuildConfiguration( _
                ByVal pConfig As ConfigurationSection, _
                Optional ByVal pPermittedServiceProviderRoles As ServiceProviderRoles = _
                                SPRoleOrderSubmissionSimulated + _
                                SPRoleHistoricalDataInput + _
                                SPRoleContractDataPrimary + _
                                SPRoleRealtimeData, _
                Optional ByVal pTwsServer As String, _
                Optional ByVal pTwsPort As Long = 7496, _
                Optional ByVal pTwsClientId As Long = -1, _
                Optional ByVal pTwsConnectionRetryIntervalSecs As Long = 60, _
                Optional ByVal pTwsLogLevel As String = "ERROR", _
                Optional ByVal pMarketDataSourceOptions As MarketDataSourceOptions = MarketDataSourceOptions.MarketDataSourceOptUseExchangeTimeZone, _
                Optional ByVal pNumberOfMarketDepthRows As Long = 20, _
                Optional ByVal pTickfilesPath As String, _
                Optional ByVal pOrderPersistenceFilePath As String)
Const ProcName As String = "SetDefaultTradeBuildConfiguration"
On Error GoTo Err

setDefaultSPConfiguration pConfig, pPermittedServiceProviderRoles, pTwsServer, pTwsPort, pTwsClientId, pTwsConnectionRetryIntervalSecs, pTwsLogLevel, pTickfilesPath, pOrderPersistenceFilePath
setDefaultMarketDataSourceSettingsConfig pConfig, pMarketDataSourceOptions, pNumberOfMarketDepthRows

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createSPConfigSection( _
                ByVal pSPsConfig As ConfigurationSection, _
                ByVal pSPName As String, _
                ByVal pSPProgId As String, _
                Optional ByVal pEnabled As Boolean = True) As ConfigurationSection
Const ProcName As String = "createSPConfigSection"
On Error GoTo Err

Dim lSPConfig As ConfigurationSection
Set lSPConfig = pSPsConfig.AddConfigurationSection(ConfigSectionServiceProvider & _
                                                "(" & pSPName & ")")
lSPConfig.SetAttribute AttributeNameServiceProviderEnabled, CStr(pEnabled)
lSPConfig.SetAttribute AttributeNameServiceProviderProgId, pSPProgId

Set createSPConfigSection = lSPConfig

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setDefaultMarketDataSourceSettingsConfig( _
                ByVal pConfig As ConfigurationSection, _
                ByVal pMarketDataSourceOptions As MarketDataSourceOptions, _
                ByVal pNumberOfMarketDepthRows As Long)
Const ProcName As String = "setDefaultMarketDataSourceSettingsConfig"
On Error GoTo Err

Dim lDataSourceConfig As ConfigurationSection
Set lDataSourceConfig = pConfig.AddConfigurationSection(ConfigSectionMarketDataSources)

lDataSourceConfig.SetAttribute ConfigSettingNoImpliedTrades, CStr((pMarketDataSourceOptions And MarketDataSourceOptNoImpliedTrades) = MarketDataSourceOptNoImpliedTrades)
lDataSourceConfig.SetAttribute ConfigSettingNoVolumeAdjustments, CStr((pMarketDataSourceOptions And MarketDataSourceOptNoVolumeAdjustments) = MarketDataSourceOptNoVolumeAdjustments)
lDataSourceConfig.SetAttribute ConfigSettingUseExchangeTimezone, CStr((pMarketDataSourceOptions And MarketDataSourceOptUseExchangeTimeZone) = MarketDataSourceOptUseExchangeTimeZone)
lDataSourceConfig.SetAttribute ConfigSettingNumberOfMarketDepthRows, CStr(pNumberOfMarketDepthRows)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setDefaultSPConfiguration( _
                ByVal pConfig As ConfigurationSection, _
                ByVal pPermittedServiceProviderRoles As ServiceProviderRoles, _
                ByVal pTwsServer As String, _
                ByVal pTwsPort As Long, _
                ByVal pTwsClientId As Long, _
                ByVal pTwsConnectionRetryIntervalSecs As Long, _
                ByVal pTwsLogLevel As String, _
                ByVal pTickfilesPath As String, _
                ByVal pOrderPersistenceFilePath As String)
Const ProcName As String = "setDefaultSPConfiguration"
On Error GoTo Err

Dim lSPConfig As ConfigurationSection
Dim lPropsConfig As ConfigurationSection

Dim lSPsConfig As ConfigurationSection
On Error Resume Next
Set lSPsConfig = pConfig.GetConfigurationSection(ConfigSectionServiceProviders)
On Error GoTo Err

If Not lSPsConfig Is Nothing Then lSPsConfig.Remove

If pTwsClientId < 0 Then pTwsClientId = Rnd * &H6FFFFFFF + &H10000000

Set lSPsConfig = pConfig.AddConfigurationSection(ConfigSectionServiceProviders, , ServiceProvidersRenderer)

If pPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleRealtimeData Then
    Set lSPConfig = createSPConfigSection(lSPsConfig, SrvcProviderNameRealtimeData, ProgIdTwsRealtimeData)

    Set lPropsConfig = lSPConfig.AddConfigurationSection(ConfigSectionProperties)
    setProperty lPropsConfig, PropertyNameTwsServer, pTwsServer
    setProperty lPropsConfig, PropertyNameTwsPort, pTwsPort
    setProperty lPropsConfig, PropertyNameTwsClientId, pTwsClientId
    setProperty lPropsConfig, PropertyNameTwsKeepConnection, "True"
    setProperty lPropsConfig, PropertyNameTwsConnectionRetryInterval, pTwsConnectionRetryIntervalSecs
    setProperty lPropsConfig, PropertyNameTwsLogLevel, pTwsLogLevel
End If
If pPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleContractDataPrimary Then
    Set lSPConfig = createSPConfigSection(lSPsConfig, SrvcProviderNamePrimaryContractData, ProgIdTwsContractData)

    Set lPropsConfig = lSPConfig.AddConfigurationSection(ConfigSectionProperties)
    setProperty lPropsConfig, PropertyNameRole, "Primary"
    setProperty lPropsConfig, PropertyNameTwsServer, pTwsServer
    setProperty lPropsConfig, PropertyNameTwsPort, pTwsPort
    setProperty lPropsConfig, PropertyNameTwsClientId, pTwsClientId
    setProperty lPropsConfig, PropertyNameTwsKeepConnection, "True"
    setProperty lPropsConfig, PropertyNameTwsConnectionRetryInterval, pTwsConnectionRetryIntervalSecs
    setProperty lPropsConfig, PropertyNameTwsLogLevel, pTwsLogLevel
End If
If pPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleContractDataSecondary Then
    Set lSPConfig = createSPConfigSection(lSPsConfig, SrvcProviderNameSecondaryContractData, ProgIdTwsContractData, False)

    Set lPropsConfig = lSPConfig.AddConfigurationSection(ConfigSectionProperties)
    setProperty lPropsConfig, PropertyNameRole, "Secondary"
    setProperty lPropsConfig, PropertyNameTwsServer, pTwsServer
    setProperty lPropsConfig, PropertyNameTwsPort, pTwsPort
    setProperty lPropsConfig, PropertyNameTwsClientId, pTwsClientId
    setProperty lPropsConfig, PropertyNameTwsKeepConnection, "True"
    setProperty lPropsConfig, PropertyNameTwsConnectionRetryInterval, pTwsConnectionRetryIntervalSecs
    setProperty lPropsConfig, PropertyNameTwsLogLevel, pTwsLogLevel
End If
If pPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleHistoricalDataInput Then
    Set lSPConfig = createSPConfigSection(lSPsConfig, SrvcProviderNameHistoricalDataInput, ProgIdTwsBarData)

    Set lPropsConfig = lSPConfig.AddConfigurationSection(ConfigSectionProperties)
    setProperty lPropsConfig, PropertyNameTwsServer, pTwsServer
    setProperty lPropsConfig, PropertyNameTwsPort, pTwsPort
    setProperty lPropsConfig, PropertyNameTwsClientId, pTwsClientId
    setProperty lPropsConfig, PropertyNameTwsKeepConnection, "True"
    setProperty lPropsConfig, PropertyNameTwsConnectionRetryInterval, pTwsConnectionRetryIntervalSecs
    setProperty lPropsConfig, PropertyNameTwsLogLevel, pTwsLogLevel
End If
If pPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleHistoricalDataOutput Then
    Set lSPConfig = createSPConfigSection(lSPsConfig, SrvcProviderNameHistoricalDataOutput, ProgIdTbBarData)

    Set lPropsConfig = lSPConfig.AddConfigurationSection(ConfigSectionProperties)
    setProperty lPropsConfig, PropertyNameDatabaseType, "Sql Server"
    setProperty lPropsConfig, PropertyNameDatabaseName, "TradingDB"
    setProperty lPropsConfig, PropertyNameDatabasePassword, ""
    setProperty lPropsConfig, PropertyNameDatabaseServer, ""
    setProperty lPropsConfig, PropertyNameDatabaseUserName, ""
    setProperty lPropsConfig, PropertyNameDatabaseUseSynchronousWrites, "False"
    setProperty lPropsConfig, PropertyNameDatabaseUseSynchronousReads, "False"
End If
If pPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleOrderPersistence Then
    Set lSPConfig = createSPConfigSection(lSPsConfig, SrvcProviderNameOrderPersistence, ProgIdTbOrderPersistence)

    Set lPropsConfig = lSPConfig.AddConfigurationSection(ConfigSectionProperties)
    setProperty lPropsConfig, PropertyNameOrderPersistenceFilePath, pOrderPersistenceFilePath
End If
If pPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleOrderSubmissionLive Then
    Set lSPConfig = createSPConfigSection(lSPsConfig, SrvcProviderNameBrokerLive, ProgIdTwsOrders)

    Set lPropsConfig = lSPConfig.AddConfigurationSection(ConfigSectionProperties)
    setProperty lPropsConfig, PropertyNameTwsServer, pTwsServer
    setProperty lPropsConfig, PropertyNameTwsPort, pTwsPort
    setProperty lPropsConfig, PropertyNameTwsClientId, pTwsClientId
    setProperty lPropsConfig, PropertyNameTwsKeepConnection, "True"
    setProperty lPropsConfig, PropertyNameTwsConnectionRetryInterval, pTwsConnectionRetryIntervalSecs
    setProperty lPropsConfig, PropertyNameTwsLogLevel, pTwsLogLevel
End If
If pPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleOrderSubmissionSimulated Then
    Set lSPConfig = createSPConfigSection(lSPsConfig, SrvcProviderNameBrokerSimulated, ProgIdTbOrders)

    Set lPropsConfig = lSPConfig.AddConfigurationSection(ConfigSectionProperties)
End If
If pPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleTickfileInput Then
    Set lSPConfig = createSPConfigSection(lSPsConfig, SrvcProviderNameTickfileInput, ProgIdFileTickData)

    Set lPropsConfig = lSPConfig.AddConfigurationSection(ConfigSectionProperties)
    setProperty lPropsConfig, PropertyNameRole, "Input"
    setProperty lPropsConfig, PropertyNameTfTickfilePath, pTickfilesPath
End If
If pPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleTickfileOutput Then
    Set lSPConfig = createSPConfigSection(lSPsConfig, SrvcProviderNameTickfileOutput, ProgIdFileTickData)

    Set lPropsConfig = lSPConfig.AddConfigurationSection(ConfigSectionProperties)
    setProperty lPropsConfig, PropertyNameRole, "Output"
    setProperty lPropsConfig, PropertyNameTfTickfilePath, pTickfilesPath
    setProperty lPropsConfig, PropertyNameTfTickfileGranularity, "FILE PER WEEK"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setProperty( _
                ByVal props As ConfigurationSection, _
                ByVal Name As String, _
                ByVal Value As String)
Const ProcName As String = "setProperty"
On Error GoTo Err

props.SetSetting "." & ConfigSectonProperty & "(" & Name & ")", Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



