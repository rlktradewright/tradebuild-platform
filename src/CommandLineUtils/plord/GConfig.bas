Attribute VB_Name = "gConfig"
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

Private Const ModuleName                            As String = "gConfig"

Private Const ConfigFileVersion                     As String = "1.3"

Private Const ConfigSectionPathSeparator            As String = "/"

Private Const ConfigSectionAppConfig                As String = "AppConfig"
Private Const ConfigSectionAppConfigs               As String = "AppConfigs"

Private Const ConfigSectionMarketDataSources        As String = "MarketDataSources"

Private Const ConfigSettingNoImpliedTrades          As String = "&NoImpliedTrades"
Private Const ConfigSettingNoVolumeAdjustments      As String = "&NoVolumeAdjustments"
Private Const ConfigSettingUseExchangeTimezone      As String = "&UseExchangeTimezone"

Private Const DefaultAppInstanceConfigName          As String = "Default Config"

Private Const AttributeNameAppInstanceConfigDefault As String = "Default"

Private Const AttributeValueTrue                    As String = "True"
Private Const AttributeValueFalse                   As String = "False"

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

Public Function gGetAppInstanceConfig( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal pName As String) As ConfigurationSection
Const ProcName As String = "gGetAppInstanceConfig"
On Error GoTo Err

Set gGetAppInstanceConfig = pConfigStore.GetConfigurationSection(generateAppInstanceSectionPath(pName))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetConfigStore() As ConfigurationStore
Const ProcName As String = "gGetConfigStore"
On Error GoTo Err

Dim lConfigStore As ConfigurationStore
Set lConfigStore = GetDefaultConfigurationStore(command, ConfigFileVersion, False, ConfigFileOptionSettingsSwitch)

If lConfigStore Is Nothing Then
    LogMessage "The configuration file does not exist: creating a new one"
    Set lConfigStore = createNewConfigStore
Else
    Assert isValidConfigurationFile(lConfigStore), "The configuration file is invalid."
End If

Set gGetConfigStore = lConfigStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetDefaultAppInstanceConfig( _
                ByVal pConfigStore As ConfigurationStore) As ConfigurationSection
Const ProcName As String = "gGetDefaultAppInstanceConfig"
On Error GoTo Err

Dim sections As ConfigurationSection
Dim section As ConfigurationSection

Set sections = getAppInstanceConfigsSection(pConfigStore)
For Each section In sections
    If CBool(section.GetAttribute(AttributeNameAppInstanceConfigDefault, AttributeValueFalse)) Then
        Set gGetDefaultAppInstanceConfig = section
        Exit Function
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetMarketDataSourcesConfig(ByVal pConfigStore As ConfigurationStore) As ConfigurationSection
Const ProcName As String = "gGetMarketDataSourcesConfig"
On Error GoTo Err

Set gGetMarketDataSourcesConfig = gGetDefaultAppInstanceConfig(pConfigStore).AddConfigurationSection(ConfigSectionMarketDataSources)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addAppInstanceConfig( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal pNewAppConfigName As String, _
                ByVal pSetAsDefault As Boolean, _
                Optional ByVal pMarketDataSourceOptions As MarketDataSourceOptions = MarketDataSourceOptions.MarketDataSourceOptUseExchangeTimeZone) As ConfigurationSection
Const ProcName As String = "addAppInstanceConfig"
On Error GoTo Err

Assert isValidConfigurationFile(pConfigStore), "Configuration file has not been correctly intialised"
AssertArgument gGetAppInstanceConfig(pConfigStore, pNewAppConfigName) Is Nothing, "App instance config already exists"

Dim newAppConfigSection As ConfigurationSection
Set newAppConfigSection = pConfigStore.AddConfigurationSection(generateAppInstanceSectionPath(pNewAppConfigName))

If pSetAsDefault Then setDefault pConfigStore, newAppConfigSection

Dim lDataSourceConfig As ConfigurationSection
Set lDataSourceConfig = newAppConfigSection.AddConfigurationSection(ConfigSectionMarketDataSources)

lDataSourceConfig.SetAttribute ConfigSettingNoImpliedTrades, CStr((pMarketDataSourceOptions And MarketDataSourceOptNoImpliedTrades) = MarketDataSourceOptNoImpliedTrades)
lDataSourceConfig.SetAttribute ConfigSettingNoVolumeAdjustments, CStr((pMarketDataSourceOptions And MarketDataSourceOptNoVolumeAdjustments) = MarketDataSourceOptNoVolumeAdjustments)
lDataSourceConfig.SetAttribute ConfigSettingUseExchangeTimezone, CStr((pMarketDataSourceOptions And MarketDataSourceOptUseExchangeTimeZone) = MarketDataSourceOptUseExchangeTimeZone)

Set addAppInstanceConfig = newAppConfigSection

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createNewConfigStore() As ConfigurationStore
Const ProcName As String = "createNewConfigStore"
On Error GoTo Err

LogMessage "Creating a new default configuration file"

Dim lConfigStore As ConfigurationStore
Set lConfigStore = GetDefaultConfigurationStore(command, ConfigFileVersion, True, ConfigFileOptionConfigSwitch)
initialiseConfigFile lConfigStore
addAppInstanceConfig lConfigStore, _
                    DefaultAppInstanceConfigName, _
                    True
lConfigStore.Save

Set createNewConfigStore = lConfigStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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

Private Function initialiseConfigFile( _
                ByVal pConfigStore As ConfigurationStore) As ConfigurationStore
Const ProcName As String = "initialiseConfigFile"
On Error GoTo Err

Dim sections As ConfigurationSection
Set sections = getAppInstanceConfigsSection(pConfigStore)

If Not sections Is Nothing Then
    LogMessage "Removing existing app instance configs section from config file"
    pConfigStore.RemoveConfigurationSection getAppInstanceConfigsPath
End If

LogMessage "Creating app instance configs section in config file"
pConfigStore.AddConfigurationSection ConfigSectionPathSeparator & ConfigSectionAppConfigs

Set initialiseConfigFile = pConfigStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isValidConfigurationFile( _
                ByVal pConfigStore As ConfigurationStore) As Boolean
Const ProcName As String = "isValidConfigurationFile"
On Error GoTo Err

If getAppInstanceConfigsSection(pConfigStore) Is Nothing Then
    isValidConfigurationFile = False
Else
    isValidConfigurationFile = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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




