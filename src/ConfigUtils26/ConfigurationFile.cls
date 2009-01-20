VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigurationFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "ConfigurationFile"

'@================================================================================
' Member variables
'@================================================================================

Private mConfigFile                                 As ConfigFile

Private mConfigSections                             As New Collection

Private mIsValid                                    As Boolean

Private mFilename                                   As String

Private mPermittedServiceProviders                  As ServiceProviderRoles

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mIsValid = False
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

Public Property Get isValid() As Boolean
isValid = mIsValid
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AddConfigurationSection( _
                ByVal sectionPath As String, _
                Optional ByVal nameAttribute As String) As ConfigurationSection
Dim sections() As String
Dim section As String
Dim i As Long
Dim parentpath As String

If Not IsValidSectionPath Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "AddConfigurationSection", _
            "Invalid section path"
End If

sections = ParsePath(sectionPath)

If sections(0) <> RootSectionName Then parentpath = SectionPathSeparator & RootSectionName

For i = 0 To UBound(sections) - 1
    parentpath = parentpath & SectionPathSeparator & section(i)
Next

GetConfigurationSection(parentpath).AddConfigurationSection sections(UBound(sections)), nameAttribute

End Function

Public Sub CreateNewApplicationConfig( _
                ByVal newAppConfigName As String)
createNewAppConfig newAppConfigName, mPermittedServiceProviders
End Sub

Public Function GetConfigurationSection( _
                ByVal sectionPath As String) As ConfigurationSection
Set gGetConfigSection = mConfigSections.Item(sectionPath)

If gGetConfigSection Is Nothing Then
    Dim parentpath As String
    Dim parentSection As ConfigurationSection
    parentpath = getParentSectionPath(sectionPath)
    Set parentSection = GetConfigurationSection(parentpath)
    If Not parentSection Is Nothing Then mConfigSections.Add gGetConfigSection, parentpath
End If
End Function

Public Function GetSetting( _
                ByVal settingPath As String) As String

End Function

Public Sub Initialise( _
                ByVal filename As String, _
                ByVal permittedServiceProviders As ServiceProviderRoles)

mFilename = filename
mPermittedServiceProviders = permittedServiceProviders

gLogger.Log LogLevelDetail, "Creating new default config file"
Set mConfigFile = CreateXMLConfigurationFile(applicationName, ConfigFileVersion)

gLogger.Log LogLevelDetail, "Creating config definitions root in config file"
mConfigFile.rootItem.childItems.AddItem ConfigNameAppConfigs

gLogger.Log LogLevelDetail, "Creating default app config in config file"
createNewAppConfig lConfigFile, DefaultConfigName

mIsValid = True
End Sub

Public Function LoadFromFile( _
                ByVal filename As String) As Boolean
Dim failpoint As Long
On Error GoTo Err

mFilename = filename
Set mConfigFile = LoadXMLConfigurationFile(mFilename)

If mConfigFile.applicationName = applicationName And _
    mConfigFile.applicationVersion <> ConfigFileVersion _
Then
    mIsValid = True
Else
    mIsValid = False
End If

LoadFromFile = True

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "LoadFromFile" & "." & failpoint & IIf(Err.source <> "", vbCrLf & Err.source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
mIsValid = False
End Function

Public Sub RemoveConfigurationSection( _
                ByVal sectionPath As String, _
                Optional ByVal nameAttribute As String)

End Sub

Public Sub RemoveSetting( _
                ByVal settingPath As String)
                
End Sub

Public Function SetSetting( _
                ByVal settingPath As String, _
                ByVal value As String, _
                Optional ByVal defaultValue) As String

End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub createNewAppConfig( _
                ByVal newAppConfigName As String, _
                ByVal permittedSPs As ServiceProviderRoles)
Dim appConfigs As ConfigItem
Dim newAppConfig As ConfigItem

Set appConfigs = mConfigFile.rootItem.childItems.Item(ConfigNameAppConfigs)

Set newAppConfig = appConfigs.childItems.AddItem(ConfigNameAppConfig)
newAppConfig.SetAttribute AttributeNameAppConfigName, newAppConfigName
newAppConfig.SetAttribute AttributeNameAppConfigDefault, "False"
newAppConfig.childItems.AddItem ConfigNameTradeBuild

TradeBuildAPI.ServiceProviders.SetDefaultSPConfig newAppConfig.childItems.Item(ConfigNameTradeBuild), _
                                        mPermittedServiceProviders
    
SetDefaultStudyLibraryConfig newAppConfig.childItems.Item(ConfigNameTradeBuild)

End Sub

Private Function getSettingName( _
                ByRef path As String) As String
Dim matches As MatchCollection
Dim lMatch As Match
gRegExp.Pattern = "[a-zA-Z0-9]+$"
matches = gRegExp.Execute(path)
lMatch = matches(0)
getSettingName = lMatch.value
End Function


