VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigUtils"
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

Public Enum ConfigFlags
    ConfigFlagSetAsDefault = 1
    ConfigFlagIncludeDefaultStudyLibrary = 2
    ConfigFlagIncludeDefaultBarFormatterLibrary = 4
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ConfigUtils"

Private Const AttributeNameAppInstanceConfigDefault As String = "Default"

Private Const AttributeValueTrue                    As String = "True"
Private Const AttributeValueFalse                   As String = "False"

Private Const ConfigSectionAppConfig                As String = "AppConfig"
Private Const ConfigSectionAppConfigs               As String = "AppConfigs"
Private Const ConfigNameTradeBuild                  As String = "TradeBuild"

Private Const ConfigSectionPathSeparator            As String = "/"

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

Public Function AddAppInstanceConfig( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal pNewAppConfigName As String, _
                ByVal pFlags As ConfigFlags, _
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
                Optional ByVal pOrderPersistenceFilePath As String) As ConfigurationSection
Const ProcName As String = "AddAppInstanceConfig"
On Error GoTo Err

checkValidConfigurationFile pConfigStore

AssertArgument GetAppInstanceConfig(pConfigStore, pNewAppConfigName) Is Nothing, "App instance config already exists"

Dim newAppConfigSection As ConfigurationSection
Set newAppConfigSection = pConfigStore.AddConfigurationSection(generateAppInstanceSectionPath(pNewAppConfigName))
If (pFlags And ConfigFlagSetAsDefault) Then setDefault pConfigStore, newAppConfigSection

Dim newAppTradeBuildSection As ConfigurationSection
Set newAppTradeBuildSection = newAppConfigSection.AddConfigurationSection(ConfigNameTradeBuild)
SetDefaultTradeBuildConfiguration newAppTradeBuildSection, _
                                    pPermittedServiceProviderRoles, _
                                    pTwsServer, _
                                    pTwsPort, _
                                    pTwsClientId, _
                                    pTwsConnectionRetryIntervalSecs, _
                                    pTwsLogLevel, _
                                    pMarketDataSourceOptions, _
                                    pNumberOfMarketDepthRows, _
                                    pTickfilesPath, _
                                    pOrderPersistenceFilePath

If (pFlags And ConfigFlagIncludeDefaultStudyLibrary) Then
    SetDefaultStudyLibraryConfig newAppTradeBuildSection
End If

If (pFlags And ConfigFlagIncludeDefaultBarFormatterLibrary) Then
    SetDefaultBarFormatterLibraryConfig newAppTradeBuildSection
End If

Set AddAppInstanceConfig = newAppConfigSection

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetAppInstanceConfig( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal pName As String) As ConfigurationSection
Const ProcName As String = "GetAppInstanceConfig"
On Error GoTo Err

Set GetAppInstanceConfig = pConfigStore.GetConfigurationSection(generateAppInstanceSectionPath(pName))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetDefaultAppInstanceConfig( _
                ByVal pConfigStore As ConfigurationStore) As ConfigurationSection
Const ProcName As String = "GetDefaultAppInstanceConfig"
On Error GoTo Err

Dim sections As ConfigurationSection
Dim section As ConfigurationSection

Set sections = getAppInstanceConfigsSection(pConfigStore)
For Each section In sections
    If CBool(section.GetAttribute(AttributeNameAppInstanceConfigDefault, AttributeValueFalse)) Then
        Set GetDefaultAppInstanceConfig = section
        Exit Function
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function InitialiseConfigFile( _
                ByVal pConfigStore As ConfigurationStore) As ConfigurationStore
Const ProcName As String = "InitialiseConfigFile"
On Error GoTo Err

Dim sections As ConfigurationSection

Set sections = getAppInstanceConfigsSection(pConfigStore)
If Not sections Is Nothing Then
    log "Removing existing app instance configs section from config file"
    pConfigStore.RemoveConfigurationSection getAppInstanceConfigsPath
End If

log "Creating app instance configs section in config file"
pConfigStore.AddConfigurationSection ConfigSectionPathSeparator & ConfigSectionAppConfigs

Set InitialiseConfigFile = pConfigStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidConfigurationFile( _
                ByVal pConfigStore As ConfigurationStore) As Boolean
Const ProcName As String = "IsValidConfigurationFile"
On Error GoTo Err

If getAppInstanceConfigsSection(pConfigStore) Is Nothing Then
    IsValidConfigurationFile = False
Else
    IsValidConfigurationFile = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub RemoveAppInstanceConfig( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal name As String)
Const ProcName As String = "RemoveAppInstanceConfig"
On Error GoTo Err

pConfigStore.RemoveConfigurationSection generateAppInstanceSectionPath(name)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function SetDefaultAppInstanceConfig( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal name As String) As ConfigurationSection
Const ProcName As String = "SetDefaultAppInstanceConfig"
On Error GoTo Err

checkValidConfigurationFile pConfigStore

Set SetDefaultAppInstanceConfig = GetAppInstanceConfig(pConfigStore, _
                                                    name)
AssertArgument Not SetDefaultAppInstanceConfig Is Nothing, "Specified app instance config does not exist"
                                            
setDefault pConfigStore, SetDefaultAppInstanceConfig

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub UnsetDefaultAppInstanceConfig( _
                ByVal pConfigStore As ConfigurationStore)
Const ProcName As String = "UnsetDefaultAppInstanceConfig"
On Error GoTo Err

Dim defaultConfig As ConfigurationSection
Set defaultConfig = GetDefaultAppInstanceConfig(pConfigStore)
If Not defaultConfig Is Nothing Then defaultConfig.SetAttribute AttributeNameAppInstanceConfigDefault, AttributeValueFalse

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkValidConfigurationFile( _
                ByVal pConfigStore As ConfigurationStore)
Const ProcName As String = "checkValidConfigurationFile"
On Error GoTo Err

AssertArgument IsValidConfigurationFile(pConfigStore), "Configuration file has not been correctly intialised"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function generateAppInstanceSectionPath( _
                ByVal name As String) As String
Const ProcName As String = "generateAppInstanceSectionPath"
On Error GoTo Err

generateAppInstanceSectionPath = ConfigSectionPathSeparator & ConfigSectionAppConfigs & ConfigSectionPathSeparator & ConfigSectionAppConfig & "(" & name & ")"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getAppInstanceConfigsPath() As String
Const ProcName As String = "getAppInstanceConfigsPath"
On Error GoTo Err

getAppInstanceConfigsPath = ConfigSectionPathSeparator & ConfigSectionAppConfigs

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getAppInstanceConfigsSection( _
                ByVal pConfigStore As ConfigurationStore) As ConfigurationSection
Const ProcName As String = "getAppInstanceConfigsSection"
On Error GoTo Err

Set getAppInstanceConfigsSection = pConfigStore.GetConfigurationSection(getAppInstanceConfigsPath)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub log( _
                ByVal message As String, _
                Optional ByVal level As LogLevels = LogLevelDetail)
Const ProcName As String = "log"
On Error GoTo Err

Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("configutils.log")

lLogger.log level, message

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setDefault( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal targetSection As ConfigurationSection)
Const ProcName As String = "setDefault"
On Error GoTo Err

Dim sections As ConfigurationSection
Dim section As ConfigurationSection

Set sections = getAppInstanceConfigsSection(pConfigStore)
For Each section In sections
    If section Is targetSection Then
        If Not CBool(section.GetAttribute(AttributeNameAppInstanceConfigDefault, AttributeValueFalse)) Then
            section.SetAttribute AttributeNameAppInstanceConfigDefault, AttributeValueTrue
        End If
    Else
        If CBool(section.GetAttribute(AttributeNameAppInstanceConfigDefault, AttributeValueFalse)) Then
            section.SetAttribute AttributeNameAppInstanceConfigDefault, AttributeValueFalse
        End If
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

