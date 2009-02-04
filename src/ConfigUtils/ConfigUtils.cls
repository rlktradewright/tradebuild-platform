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

Private Const AttributeNameAppConfigDefault          As String = "Default"

Private Const ConfigNameAppConfig                    As String = "AppConfig"
Private Const ConfigNameAppConfigs                   As String = "AppConfigs"
Private Const ConfigNameTradeBuild                   As String = "TradeBuild"

Private Const ConfigSectionPathSeparator             As String = "/"

Private Const DefaultAppConfigName                  As String = "Default config"

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

Public Function CreateNewTBAppConfig( _
                ByVal pConfigFile As ConfigurationFile, _
                ByVal name As String, _
                ByVal permittedServiceProviders As ServiceProviderRoles, _
                ByVal includeDefaultStudyLibrary As Boolean) As ConfigurationSection
Set CreateNewTBAppConfig = createNewAppConfig(pConfigFile, _
                                            name, _
                                            permittedServiceProviders, _
                                            includeDefaultStudyLibrary)
End Function

Public Function CreateTradeBuildConfig( _
                ByVal pConfigFile As ConfigFile, _
                ByVal filename As String, _
                ByVal permittedServiceProviders As ServiceProviderRoles, _
                ByVal includeDefaultStudyLibrary As Boolean) As ConfigurationFile
Dim failpoint As Long
On Error GoTo Err

Set CreateTradeBuildConfig = CreateConfigurationFile(pConfigFile, _
                                    filename)

GetLogger("configutils.log").Log LogLevelDetail, "Creating app configs section in config file"
CreateTradeBuildConfig.AddConfigurationSection (ConfigSectionPathSeparator & ConfigNameAppConfigs)

GetLogger("configutils.log").Log LogLevelDetail, "Creating default app config in config file"
createNewAppConfig CreateTradeBuildConfig, _
                DefaultAppConfigName, _
                permittedServiceProviders, _
                includeDefaultStudyLibrary

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "CreateConfigurationFile" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
GetLogger("configutils.log").Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createNewAppConfig( _
                ByVal pConfigFile As ConfigurationFile, _
                ByVal newAppConfigName As String, _
                ByVal permittedServiceProviders As ServiceProviderRoles, _
                ByVal includeDefaultStudyLibrary As Boolean) As ConfigurationSection
Dim newAppConfigSection As ConfigurationSection
Dim newAppTradeBuildSection As ConfigurationSection

Set newAppConfigSection = pConfigFile.AddConfigurationSection(ConfigSectionPathSeparator & ConfigNameAppConfigs & ConfigSectionPathSeparator & ConfigNameAppConfig & "(" & newAppConfigName & ")")
newAppConfigSection.SetAttribute AttributeNameAppConfigDefault, "False"

Set newAppTradeBuildSection = newAppConfigSection.AddConfigurationSection(ConfigNameTradeBuild)
TradeBuildAPI.ServiceProviders.SetDefaultSPConfig newAppTradeBuildSection, _
                                                permittedServiceProviders

If includeDefaultStudyLibrary Then
    SetDefaultStudyLibraryConfig newAppTradeBuildSection
End If

Set createNewAppConfig = newAppConfigSection

End Function


