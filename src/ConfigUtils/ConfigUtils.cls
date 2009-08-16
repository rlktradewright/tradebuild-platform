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

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                           As String = "ConfigUtils26"
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
                ByVal pConfigFile As ConfigurationFile, _
                ByVal name As String, _
                ByVal includeDefaultStudyLibrary As Boolean, _
                Optional ByVal setAsDefault As Boolean = False, _
                Optional ByVal twsServer As String, _
                Optional ByVal twsPort As Long = 7496, _
                Optional ByVal clientId As Long = -1, _
                Optional ByVal connectionRetryIntervalSecs As Long = 10, _
                Optional ByVal twsLogLevel As String = "ERROR", _
                Optional ByVal tickfilesPath As String) As ConfigurationSection
Dim failpoint As Long
On Error GoTo Err

Set AddAppInstanceConfig = addNewAppConfig(pConfigFile, _
                                            name, _
                                            includeDefaultStudyLibrary, _
                                            setAsDefault, _
                                            twsServer, _
                                            twsPort, _
                                            clientId, _
                                            connectionRetryIntervalSecs, _
                                            twsLogLevel, _
                                            tickfilesPath)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "AddAppInstanceConfig" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
log "Error " & errNumber & ": " & errDescription & vbCrLf & errSource, LogLevelSevere
Err.Raise errNumber, errSource, errDescription
End Function

Public Function ConfigureTradeBuild( _
                ByVal pConfigFile As ConfigurationFile, _
                ByVal appInstanceConfigName As String) As Boolean
Dim config As ConfigurationSection

Set config = GetAppInstanceConfig(pConfigFile, appInstanceConfigName)
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
End Function

Public Function GetAppInstanceConfig( _
                ByVal pConfigFile As ConfigurationFile, _
                ByVal name As String) As ConfigurationSection
Dim failpoint As Long
On Error GoTo Err

Set GetAppInstanceConfig = pConfigFile.GetConfigurationSection(generateSectionPath(name))

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GetAppInstanceConfig" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
log "Error " & errNumber & ": " & errDescription & vbCrLf & errSource, LogLevelSevere
Err.Raise errNumber, errSource, errDescription
End Function

Public Function GetDefaultAppInstanceConfig( _
                ByVal pConfigFile As ConfigurationFile) As ConfigurationSection
Dim sections As ConfigurationSection
Dim section As ConfigurationSection

Set sections = getAppInstanceConfigsSection(pConfigFile)
For Each section In sections
    If CBool(section.getAttribute(AttributeNameAppInstanceConfigDefault, AttributeValueFalse)) Then
        Set GetDefaultAppInstanceConfig = section
        Exit Function
    End If
Next

End Function

Public Function InitialiseConfigFile( _
                ByVal pConfigFile As ConfigurationFile) As ConfigurationFile
Dim sections As ConfigurationSection
Dim failpoint As Long
On Error GoTo Err

Set sections = getAppInstanceConfigsSection(pConfigFile)
If Not sections Is Nothing Then
    log "Removing existing app instance configs section from config file"
    pConfigFile.RemoveConfigurationSection getAppInstanceConfigsPath
End If

log "Creating app instance configs section in config file"
pConfigFile.AddConfigurationSection ConfigSectionPathSeparator & ConfigNameAppConfigs

Set InitialiseConfigFile = pConfigFile

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "CreateConfigurationFile" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
log "Error " & errNumber & ": " & errDescription & vbCrLf & errSource, LogLevelSevere
Err.Raise errNumber, errSource, errDescription
End Function

Public Function IsValidConfigurationFile( _
                ByVal pConfigFile As ConfigurationFile) As Boolean
If getAppInstanceConfigsSection(pConfigFile) Is Nothing Then
    IsValidConfigurationFile = False
Else
    IsValidConfigurationFile = True
End If
End Function

Public Sub RemoveAppInstanceConfig( _
                ByVal pConfigFile As ConfigurationFile, _
                ByVal name As String)
Dim failpoint As Long
On Error GoTo Err

pConfigFile.RemoveConfigurationSection generateSectionPath(name)

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RemoveAppInstanceConfig" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
log "Error " & errNumber & ": " & errDescription & vbCrLf & errSource, LogLevelSevere
Err.Raise errNumber, errSource, errDescription
End Sub

Public Function SetDefaultAppInstanceConfig( _
                ByVal pConfigFile As ConfigurationFile, _
                ByVal name As String) As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err

checkValidConfigurationFile pConfigFile

Set SetDefaultAppInstanceConfig = GetAppInstanceConfig(pConfigFile, _
                                                    name)
If SetDefaultAppInstanceConfig Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "SetDefaultAppInstanceConfig", _
            "Specified app instance config does not exist"
End If
                                            
setDefault pConfigFile, SetDefaultAppInstanceConfig

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetDefaultAppInstanceConfig" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
log "Error " & errNumber & ": " & errDescription & vbCrLf & errSource, LogLevelSevere
Err.Raise errNumber, errSource, errDescription
End Function

Public Sub UnsetDefaultAppInstanceConfig( _
                ByVal pConfigFile As ConfigurationFile)
Dim defaultConfig As ConfigurationSection
Set defaultConfig = GetDefaultAppInstanceConfig(pConfigFile)
If Not defaultConfig Is Nothing Then defaultConfig.setAttribute AttributeNameAppInstanceConfigDefault, AttributeValueFalse
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addNewAppConfig( _
                ByVal pConfigFile As ConfigurationFile, _
                ByVal newAppConfigName As String, _
                ByVal includeDefaultStudyLibrary As Boolean, _
                ByVal setAsDefault As Boolean, _
                Optional ByVal twsServer As String, _
                Optional ByVal twsPort As Long = 7496, _
                Optional ByVal clientId As Long = -1, _
                Optional ByVal connectionRetryIntervalSecs As Long = 60, _
                Optional ByVal twsLogLevel As String = "ERROR", _
                Optional ByVal tickfilesPath As String) As ConfigurationSection
Dim newAppConfigSection As ConfigurationSection
Dim newAppTradeBuildSection As ConfigurationSection

checkValidConfigurationFile pConfigFile

If Not GetAppInstanceConfig(pConfigFile, newAppConfigName) Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "addNewAppConfig", _
            "App instance config already exists"
End If

Set newAppConfigSection = pConfigFile.AddConfigurationSection(generateSectionPath(newAppConfigName))
If setAsDefault Then setDefault pConfigFile, newAppConfigSection

Set newAppTradeBuildSection = newAppConfigSection.AddConfigurationSection(ConfigNameTradeBuild)
TradeBuildAPI.ServiceProviders.SetDefaultSPConfig newAppTradeBuildSection, _
                                                    twsServer, _
                                                    twsPort, _
                                                    clientId, _
                                                    connectionRetryIntervalSecs, _
                                                    twsLogLevel, _
                                                    tickfilesPath

If includeDefaultStudyLibrary Then
    SetDefaultStudyLibraryConfig newAppTradeBuildSection
End If

Set addNewAppConfig = newAppConfigSection

End Function

Private Sub checkValidConfigurationFile( _
                ByVal pConfigFile As ConfigurationFile)
If Not IsValidConfigurationFile(pConfigFile) Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "checkValidConfigurationFile", _
            "Configuration file has not been correctly intialised"
End If
End Sub

Private Function generateSectionPath( _
                ByVal name As String) As String
generateSectionPath = ConfigSectionPathSeparator & ConfigNameAppConfigs & ConfigSectionPathSeparator & ConfigNameAppConfig & "(" & name & ")"
End Function

Private Function getAppInstanceConfigsPath() As String
getAppInstanceConfigsPath = ConfigSectionPathSeparator & ConfigNameAppConfigs
End Function

Private Function getAppInstanceConfigsSection( _
                ByVal pConfigFile As ConfigurationFile) As ConfigurationSection
Set getAppInstanceConfigsSection = pConfigFile.GetConfigurationSection(getAppInstanceConfigsPath)
End Function

Private Sub log( _
                ByVal message As String, _
                Optional ByVal level As LogLevels = LogLevelDetail)
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("configutils.log")

lLogger.log level, message

End Sub

Private Sub removeServiceProviders()
If Not TradeBuildAPI Is Nothing Then TradeBuildAPI.ServiceProviders.RemoveAll
RemoveAllStudyLibraries
End Sub

Private Function setupServiceProviders( _
                ByVal config As ConfigurationSection) As Boolean
Dim tradebuildEntry As ConfigurationSection

Set tradebuildEntry = config.GetConfigurationSection(ConfigNameTradeBuild)

If tradebuildEntry Is Nothing Then
    log "No service providers defined in this configuration", LogLevelNormal
    setupServiceProviders = True
Else
    On Error GoTo Err
    TradeBuildAPI.ServiceProviders.LoadServiceProviderConfiguration tradebuildEntry
    setupServiceProviders = True
End If

Exit Function

Err:
log "Service provider configuration failed: " & Err.Description, LogLevelNormal
setupServiceProviders = False
End Function

Private Function setupStudyLibraries( _
                ByVal config As ConfigurationSection) As Boolean
Dim tradebuildEntry As ConfigurationSection

Set tradebuildEntry = config.GetConfigurationSection(ConfigNameTradeBuild)

If tradebuildEntry Is Nothing Then
    log "No study libraries defined in this configuration", LogLevelNormal
    setupStudyLibraries = True
Else
    On Error GoTo Err
    LoadStudyLibraryConfiguration tradebuildEntry
    setupStudyLibraries = True
End If

Exit Function

Err:
log "Study library configuration failed: " & Err.Description, LogLevelNormal
setupStudyLibraries = False
End Function

Private Sub setDefault( _
                ByVal pConfigFile As ConfigurationFile, _
                ByVal targetSection As ConfigurationSection)
Dim sections As ConfigurationSection
Dim section As ConfigurationSection

Set sections = getAppInstanceConfigsSection(pConfigFile)
For Each section In sections
    If section Is targetSection Then
        If Not CBool(section.getAttribute(AttributeNameAppInstanceConfigDefault, AttributeValueFalse)) Then
            section.setAttribute AttributeNameAppInstanceConfigDefault, AttributeValueTrue
        End If
    Else
        If CBool(section.getAttribute(AttributeNameAppInstanceConfigDefault, AttributeValueFalse)) Then
            section.setAttribute AttributeNameAppInstanceConfigDefault, AttributeValueFalse
        End If
    End If
Next

End Sub

