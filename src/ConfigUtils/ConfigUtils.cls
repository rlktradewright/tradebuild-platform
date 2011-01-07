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

Private Const ConfigNameAppConfig                   As String = "AppConfig"
Private Const ConfigNameAppConfigs                  As String = "AppConfigs"
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
                ByVal pName As String, _
                ByVal pFlags As ConfigFlags, _
                Optional ByVal twsServer As String, _
                Optional ByVal twsPort As Long = 7496, _
                Optional ByVal clientId As Long = -1, _
                Optional ByVal connectionRetryIntervalSecs As Long = 10, _
                Optional ByVal twsLogLevel As String = "ERROR", _
                Optional ByVal tickfilesPath As String) As ConfigurationSection
Const ProcName As String = "AddAppInstanceConfig"
Dim failpoint As String
On Error GoTo Err

Set AddAppInstanceConfig = addNewAppConfig(pConfigStore, _
                                            pName, _
                                            pFlags, _
                                            twsServer, _
                                            twsPort, _
                                            clientId, _
                                            connectionRetryIntervalSecs, _
                                            twsLogLevel, _
                                            tickfilesPath)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function ConfigureTradeBuild( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal appInstanceConfigName As String) As Boolean
Dim config As ConfigurationSection

Const ProcName As String = "ConfigureTradeBuild"
Dim failpoint As String
On Error GoTo Err

Set config = GetAppInstanceConfig(pConfigStore, appInstanceConfigName)
If config Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "ConfigureTradeBuild", _
            "App instance config does not exist"
End If

removeServiceProviders

log "Configuring TradeBuild using config: " & config.InstanceQualifier, LogLevelNormal

ConfigureTradeBuild = True
If Not setupServiceProviders(config) Then
    log "Failed loading service providers from configuration: " & config.InstanceQualifier, LogLevelNormal
    ConfigureTradeBuild = False
End If
If Not setupStudyLibraries(config) Then
    log "Failed loading study libraries from configuration: " & config.InstanceQualifier, LogLevelNormal
    ConfigureTradeBuild = False
End If

If ConfigureTradeBuild Then
    log "Loaded configuration: " & config.InstanceQualifier, LogLevelNormal
Else
    log "Failed to load configuration: " & config.InstanceQualifier, LogLevelNormal
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function GetAppInstanceConfig( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal name As String) As ConfigurationSection
Const ProcName As String = "GetAppInstanceConfig"
Dim failpoint As String
On Error GoTo Err

Set GetAppInstanceConfig = pConfigStore.GetConfigurationSection(generateSectionPath(name))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function GetDefaultAppInstanceConfig( _
                ByVal pConfigStore As ConfigurationStore) As ConfigurationSection
Dim sections As ConfigurationSection
Dim section As ConfigurationSection

Const ProcName As String = "GetDefaultAppInstanceConfig"
Dim failpoint As String
On Error GoTo Err

Set sections = getAppInstanceConfigsSection(pConfigStore)
For Each section In sections
    If CBool(section.GetAttribute(AttributeNameAppInstanceConfigDefault, AttributeValueFalse)) Then
        Set GetDefaultAppInstanceConfig = section
        Exit Function
    End If
Next

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function

Public Function InitialiseConfigFile( _
                ByVal pConfigStore As ConfigurationStore) As ConfigurationStore
Dim sections As ConfigurationSection

Const ProcName As String = "InitialiseConfigFile"
Dim failpoint As String
On Error GoTo Err

Set sections = getAppInstanceConfigsSection(pConfigStore)
If Not sections Is Nothing Then
    log "Removing existing app instance configs section from config file"
    pConfigStore.RemoveConfigurationSection getAppInstanceConfigsPath
End If

log "Creating app instance configs section in config file"
pConfigStore.AddConfigurationSection ConfigSectionPathSeparator & ConfigNameAppConfigs

Set InitialiseConfigFile = pConfigStore

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function IsValidConfigurationFile( _
                ByVal pConfigStore As ConfigurationStore) As Boolean
Const ProcName As String = "IsValidConfigurationFile"
Dim failpoint As String
On Error GoTo Err

If getAppInstanceConfigsSection(pConfigStore) Is Nothing Then
    IsValidConfigurationFile = False
Else
    IsValidConfigurationFile = True
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub RemoveAppInstanceConfig( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal name As String)
Const ProcName As String = "RemoveAppInstanceConfig"
Dim failpoint As String
On Error GoTo Err

pConfigStore.RemoveConfigurationSection generateSectionPath(name)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function SetDefaultAppInstanceConfig( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal name As String) As ConfigurationSection
Const ProcName As String = "SetDefaultAppInstanceConfig"
Dim failpoint As String
On Error GoTo Err

checkValidConfigurationFile pConfigStore

Set SetDefaultAppInstanceConfig = GetAppInstanceConfig(pConfigStore, _
                                                    name)
If SetDefaultAppInstanceConfig Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "SetDefaultAppInstanceConfig", _
            "Specified app instance config does not exist"
End If
                                            
setDefault pConfigStore, SetDefaultAppInstanceConfig

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub UnsetDefaultAppInstanceConfig( _
                ByVal pConfigStore As ConfigurationStore)
Dim defaultConfig As ConfigurationSection
Const ProcName As String = "UnsetDefaultAppInstanceConfig"
Dim failpoint As String
On Error GoTo Err

Set defaultConfig = GetDefaultAppInstanceConfig(pConfigStore)
If Not defaultConfig Is Nothing Then defaultConfig.SetAttribute AttributeNameAppInstanceConfigDefault, AttributeValueFalse

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addNewAppConfig( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal newAppConfigName As String, _
                ByVal pFlags As ConfigFlags, _
                Optional ByVal twsServer As String, _
                Optional ByVal twsPort As Long = 7496, _
                Optional ByVal clientId As Long = -1, _
                Optional ByVal connectionRetryIntervalSecs As Long = 60, _
                Optional ByVal twsLogLevel As String = "ERROR", _
                Optional ByVal tickfilesPath As String) As ConfigurationSection
Dim newAppConfigSection As ConfigurationSection
Dim newAppTradeBuildSection As ConfigurationSection

Const ProcName As String = "addNewAppConfig"
Dim failpoint As String
On Error GoTo Err

checkValidConfigurationFile pConfigStore

If Not GetAppInstanceConfig(pConfigStore, newAppConfigName) Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "addNewAppConfig", _
            "App instance config already exists"
End If

Set newAppConfigSection = pConfigStore.AddConfigurationSection(generateSectionPath(newAppConfigName))
If (pFlags And ConfigFlagSetAsDefault) Then setDefault pConfigStore, newAppConfigSection

Set newAppTradeBuildSection = newAppConfigSection.AddConfigurationSection(ConfigNameTradeBuild)
TradeBuildAPI.ServiceProviders.SetDefaultSPConfig newAppTradeBuildSection, _
                                                    twsServer, _
                                                    twsPort, _
                                                    clientId, _
                                                    connectionRetryIntervalSecs, _
                                                    twsLogLevel, _
                                                    tickfilesPath

If (pFlags And ConfigFlagIncludeDefaultStudyLibrary) Then
    SetDefaultStudyLibraryConfig newAppTradeBuildSection
End If

If (pFlags And ConfigFlagIncludeDefaultBarFormatterLibrary) Then
    SetDefaultBarFormatterLibraryConfig newAppTradeBuildSection
End If

Set addNewAppConfig = newAppConfigSection

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function

Private Sub checkValidConfigurationFile( _
                ByVal pConfigStore As ConfigurationStore)
Const ProcName As String = "checkValidConfigurationFile"
Dim failpoint As String
On Error GoTo Err

If Not IsValidConfigurationFile(pConfigStore) Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "checkValidConfigurationFile", _
            "Configuration file has not been correctly intialised"
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function generateSectionPath( _
                ByVal name As String) As String
Const ProcName As String = "generateSectionPath"
Dim failpoint As String
On Error GoTo Err

generateSectionPath = ConfigSectionPathSeparator & ConfigNameAppConfigs & ConfigSectionPathSeparator & ConfigNameAppConfig & "(" & name & ")"

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getAppInstanceConfigsPath() As String
Const ProcName As String = "getAppInstanceConfigsPath"
Dim failpoint As String
On Error GoTo Err

getAppInstanceConfigsPath = ConfigSectionPathSeparator & ConfigNameAppConfigs

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getAppInstanceConfigsSection( _
                ByVal pConfigStore As ConfigurationStore) As ConfigurationSection
Const ProcName As String = "getAppInstanceConfigsSection"
Dim failpoint As String
On Error GoTo Err

Set getAppInstanceConfigsSection = pConfigStore.GetConfigurationSection(getAppInstanceConfigsPath)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub log( _
                ByVal message As String, _
                Optional ByVal level As LogLevels = LogLevelDetail)
Static lLogger As Logger
Const ProcName As String = "log"
Dim failpoint As String
On Error GoTo Err

If lLogger Is Nothing Then Set lLogger = GetLogger("configutils.log")

lLogger.log level, message

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub removeServiceProviders()
Const ProcName As String = "removeServiceProviders"
Dim failpoint As String
On Error GoTo Err

If Not TradeBuildAPI Is Nothing Then TradeBuildAPI.ServiceProviders.RemoveAll
RemoveAllStudyLibraries

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function setupServiceProviders( _
                ByVal config As ConfigurationSection) As Boolean
Dim tradebuildEntry As ConfigurationSection

Const ProcName As String = "setupServiceProviders"
Dim failpoint As String
On Error GoTo Err

Set tradebuildEntry = config.GetConfigurationSection(ConfigNameTradeBuild)

If tradebuildEntry Is Nothing Then
    log "No service providers defined in this configuration", LogLevelNormal
    setupServiceProviders = True
Else
    On Error GoTo ConfigErr
    TradeBuildAPI.ServiceProviders.LoadServiceProviderConfiguration tradebuildEntry
    setupServiceProviders = True
End If

Exit Function

ConfigErr:
log "Service provider configuration failed: " & Err.Description & vbCrLf & _
    Err.Source, _
    LogLevelSevere
setupServiceProviders = False

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function setupStudyLibraries( _
                ByVal config As ConfigurationSection) As Boolean
Dim tradebuildEntry As ConfigurationSection

Const ProcName As String = "setupStudyLibraries"
Dim failpoint As String
On Error GoTo Err

Set tradebuildEntry = config.GetConfigurationSection(ConfigNameTradeBuild)

If tradebuildEntry Is Nothing Then
    log "No study libraries defined in this configuration", LogLevelNormal
    setupStudyLibraries = True
Else
    On Error GoTo ConfigErr
    LoadStudyLibraryConfiguration tradebuildEntry
    setupStudyLibraries = True
End If

Exit Function

ConfigErr:
log "Study library configuration failed: " & Err.Description & vbCrLf & _
    Err.Source, _
    LogLevelSevere
setupStudyLibraries = False

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub setDefault( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal targetSection As ConfigurationSection)
Dim sections As ConfigurationSection
Dim section As ConfigurationSection

Const ProcName As String = "setDefault"
Dim failpoint As String
On Error GoTo Err

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

