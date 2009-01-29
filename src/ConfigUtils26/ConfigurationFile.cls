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

Private mConfigFile                                 As configFile

Private mFilename                                   As String

Private mPermittedServiceProviders                  As ServiceProviderRoles

Private mRootConfigSection                          As ConfigurationSection
Private mAppConfigsSection                          As ConfigurationSection

Private mCache                                      As New ConfigSectionCache

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

Public Property Get ApplicationName() As String
ApplicationName = mConfigFile.ApplicationName
End Property

Public Property Get Dirty() As Boolean
Dirty = mConfigFile.Dirty
End Property

Public Property Get FileVersion() As String
FileVersion = mConfigFile.applicationVersion
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AddConfigurationSection( _
                ByVal configSectionPath As String, _
                Optional ByVal Comment As String, _
                Optional ByVal RendererProgId As String) As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err

Set AddConfigurationSection = AddConfigSection(configSectionPath, False, Comment, RendererProgId)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "AddConfigurationSection" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Public Function AddPrivateConfigurationSection( _
                ByVal configSectionPath As String, _
                Optional ByVal Comment As String, _
                Optional ByVal RendererProgId As String) As ConfigurationSection
Dim failpoint As Long
On Error GoTo Err

Set AddPrivateConfigurationSection = AddConfigSection(configSectionPath, True, Comment, RendererProgId)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "AddPrivateConfigurationSection" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function CreateNewApplicationConfig( _
                ByVal newAppConfigName As String, _
                ByVal includeDefaultStudyLibrary As Boolean) As ConfigurationSection
Set CreateNewApplicationConfig = createNewAppConfig(newAppConfigName, includeDefaultStudyLibrary)
End Function

Public Function GetConfigurationSection( _
                ByVal configSectionPath As String) As ConfigurationSection
Dim sectionPath As ConfigurationPath

Dim failpoint As Long
On Error GoTo Err

If configSectionPath = "" Then
    Set GetConfigurationSection = mRootConfigSection
    Exit Function
End If

Set GetConfigurationSection = mCache.Item(configSectionPath)

If GetConfigurationSection Is Nothing Then
    Set sectionPath = gGetConfigPath(configSectionPath)
    If Not sectionPath.IsValidSectionPath Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & "GetConfigurationSection", _
                "Not a valid configuration section path"
    End If
    
    Dim parentSection As ConfigurationSection
    Set parentSection = GetConfigurationSection(sectionPath.ParentPath)
    If Not parentSection Is Nothing Then
        Set GetConfigurationSection = parentSection.GetConfigurationSection(sectionPath.TailPath)
    Else
        Set GetConfigurationSection = mRootConfigSection.GetConfigurationSection(sectionPath.TailPath)
    End If
End If

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GetConfigurationSection" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function GetSetting( _
                ByVal settingPath As String, _
                Optional ByVal defaultValue As String) As String
Dim cp As ConfigurationPath
Dim section As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err

Set cp = gGetConfigPath(settingPath)
If cp.IsValidAttributePath Then
ElseIf cp.IsValidValuePath Then
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "GetSetting", _
            "Not a valid attribute or value path"
End If

If cp.IsRelative Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "GetSetting", _
            "Not relative attribute or value path not permitted"
End If

Set section = GetConfigurationSection(cp.ParentPath)

If section Is Nothing Then
    GetSetting = defaultValue
ElseIf cp.IsValidAttributePath Then
    GetSetting = section.GetAttribute(cp.TailName, defaultValue)
Else
    GetSetting = section.GetConfigurationSection(Right$(cp.TailPath, Len(cp.TailPath) - 1)).value
End If

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GetSetting" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Friend Sub Initialise( _
                ByVal pConfigFile As configFile, _
                ByVal filename As String, _
                ByVal permittedServiceProviders As ServiceProviderRoles, _
                ByVal includeDefaultStudyLibrary As Boolean)

mFilename = filename
mPermittedServiceProviders = permittedServiceProviders

Set mConfigFile = pConfigFile
Set mRootConfigSection = New ConfigurationSection
mRootConfigSection.Initialise Me, "", mConfigFile.rootItem, mCache, False

gLogger.Log LogLevelDetail, "Creating app configs section in config file"
Set mAppConfigsSection = AddConfigurationSection(SectionPathSeparator & ConfigNameAppConfigs)

gLogger.Log LogLevelDetail, "Creating default app config in config file"
createNewAppConfig DefaultAppConfigName, includeDefaultStudyLibrary

End Sub

Public Sub RemoveConfigurationSection( _
                ByVal configSectionPath As String)
Dim sectionPath As ConfigurationPath

Dim failpoint As Long
On Error GoTo Err

Set sectionPath = gGetConfigPath(configSectionPath)
If Not sectionPath.IsValidSectionPath Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "RemoveConfigurationSection", _
            "Not a valid configuration section path"
End If

Dim parentSection As ConfigurationSection
Set parentSection = GetConfigurationSection(sectionPath.ParentPath)
If Not parentSection Is Nothing Then
    parentSection.RemoveConfigurationSection sectionPath.TailPath
Else
    mRootConfigSection.RemoveConfigurationSection sectionPath.TailPath
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RemoveConfigurationSection" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub RemoveSetting( _
                ByVal settingPath As String)
Dim cp As ConfigurationPath
Dim section As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err

Set cp = gGetConfigPath(settingPath)
If cp.IsValidAttributePath Then
ElseIf cp.IsValidValuePath Then
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "RemoveSetting", _
            "Not a valid attribute or value path"
End If

If cp.IsRelative Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "RemoveSetting", _
            "Not relative attribute or value path not permitted"
End If

Set section = GetConfigurationSection(cp.ParentPath)
If section Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "RemoveSetting", _
            "Invalid path - config section does not exist"
End If

section.RemoveSetting cp.TailPath

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RemoveSetting" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Sub Save( _
                Optional ByVal filename As String)
Dim failpoint As Long
On Error GoTo Err

If filename <> "" Then
    mConfigFile.Save filename
Else
    mConfigFile.Save mFilename
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Save" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub SetBooleanSetting( _
                ByVal settingPath As String, _
                ByVal value As Boolean)
Dim section As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err

Set section = SetSettingEx(settingPath, value)
section.SetAttribute AttributeNameType, AttributeValueTypeBoolean

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetBooleanSetting" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub
                
Public Sub SetSelectionSetting( _
                ByVal settingPath As String, _
                ByVal value As String, _
                ByRef permittedValues() As String)
Dim section As ConfigurationSection
Dim permittedValue As Variant
Dim selsItem As ConfigItem

Dim failpoint As Long

On Error GoTo Err

Set section = SetSettingEx(settingPath, CStr(value))
section.SetAttribute AttributeNameType, AttributeValueTypeSelection
Set selsItem = section.Config.childItems.AddItem(ConfigNameSelections)

For Each permittedValue In permittedValues
    selsItem.childItems.AddItem ConfigNameSelection, gXMLEncode(permittedValue)
Next
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetSelectionSetting" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub SetSetting( _
                ByVal settingPath As String, _
                ByVal value As String)
Dim failpoint As Long
On Error GoTo Err

SetSettingEx settingPath, value

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetSetting" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function AddConfigSection( _
                ByVal configSectionPath As String, _
                ByVal isPrivate As Boolean, _
                ByVal Comment As String, _
                ByVal RendererProgId As String) As ConfigurationSection
Dim sectionPath As ConfigurationPath
Dim parentSection As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err

Set sectionPath = gGetConfigPath(configSectionPath)

If Not sectionPath.IsValidSectionPath Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "AddConfigSection", _
            "Invalid section path"
End If

If sectionPath.IsRelative Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "AddConfigSection", _
            "Relative section path not permitted"
End If

Set AddConfigSection = GetConfigurationSection(configSectionPath)

If AddConfigSection Is Nothing Then
    If sectionPath.ParentPath = "" Then
        If isPrivate Then
            Set AddConfigSection = mRootConfigSection.AddPrivateConfigurationSection(sectionPath.TailPath, Comment, RendererProgId)
        Else
            Set AddConfigSection = mRootConfigSection.AddConfigurationSection(sectionPath.TailPath, Comment, RendererProgId)
        End If
    Else
        Set parentSection = GetConfigurationSection(sectionPath.ParentPath)
        If parentSection Is Nothing Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    ProjectName & "." & ModuleName & ":" & "AddConfigSection", _
                    "Invalid path - config section does not exist"
        End If
        
        If isPrivate Then
            Set AddConfigSection = parentSection.AddPrivateConfigurationSection(sectionPath.TailPath, Comment, RendererProgId)
        Else
            Set AddConfigSection = parentSection.AddConfigurationSection(sectionPath.TailPath, Comment, RendererProgId)
        End If
        
    End If
End If

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "AddConfigSection" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
Err.Raise errNumber, errSource, errDescription

End Function

Private Function createNewAppConfig( _
                ByVal newAppConfigName As String, _
                ByVal includeDefaultStudyLibrary As Boolean) As ConfigurationSection
Dim newAppConfigSection As ConfigurationSection
Dim newAppTradeBuildSection As ConfigurationSection

Set newAppConfigSection = mAppConfigsSection.AddConfigurationSection(ConfigNameAppConfig & "(" & newAppConfigName & ")")
newAppConfigSection.SetAttribute AttributeNameAppConfigDefault, "False"

Set newAppTradeBuildSection = newAppConfigSection.AddConfigurationSection(ConfigNameTradeBuild)
TradeBuildAPI.ServiceProviders.SetDefaultSPConfig newAppTradeBuildSection, _
                                                mPermittedServiceProviders

If includeDefaultStudyLibrary Then
    SetDefaultStudyLibraryConfig newAppTradeBuildSection.Config
End If

Set createNewAppConfig = newAppConfigSection

End Function

Private Function SetSettingEx( _
                ByVal settingPath As String, _
                ByVal value As String) As ConfigurationSection
Dim cp As ConfigurationPath
Dim section As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err

Set cp = gGetConfigPath(settingPath)
If cp.IsValidAttributePath Then
ElseIf cp.IsValidValuePath Then
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "SetSetting", _
            "Not a valid attribute or value path"
End If

If cp.IsRelative Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "SetSetting", _
            "Not relative attribute or value path not permitted"
End If

Set section = GetConfigurationSection(cp.ParentPath)
If section Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "SetSetting", _
            "Invalid path - config section does not exist"
End If

Set SetSettingEx = section.SetSettingEx(cp.TailPath, value)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetSetting" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
Err.Raise errNumber, errSource, errDescription

End Function



