VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigManager"
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

Implements ChangeListener

'@================================================================================
' Events
'@================================================================================

Event Clean()
Event Dirty()

Event SelectedItemChanged()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "TradeSkilDemo26"
Private Const ModuleName                    As String = "ConfigManager"

Private Const ConfigNameAppConfig           As String = "AppConfig"
Private Const ConfigNameAppConfigs          As String = "AppConfigs"

Private Const DefaultAppInstanceConfigName  As String = "Default Config"

Private Const NewConfigNameStub             As String = "New config"

Private Const TWSClientId                   As Long = 414552987
Private Const TWSConnectRetryInterval       As Long = 10
Private Const TWSPort                       As Long = 7496
Private Const TWSServer                     As String = ""

Private Const TickfilesPath                 As String = "C:\Data\Tickfiles"

'@================================================================================
' Member variables
'@================================================================================

Private mConfigFilename                     As String
Private mConfigFile                         As ConfigurationFile
Private mAppConfigs                         As ConfigurationSection

Private mCurrAppConfig                      As ConfigurationSection

Private mDefaultAppConfig                   As ConfigurationSection

Private mConfigNames                        As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mConfigNames = New Collection
End Sub

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change( _
                ev As TWUtilities30.ChangeEvent)
If ev.Source Is mConfigFile Then
    Select Case ev.changeType
    Case ConfigChangeTypes.ConfigClean
        RaiseEvent Clean
    Case ConfigChangeTypes.ConfigDirty
        RaiseEvent Dirty
    End Select
End If
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get appConfig( _
                ByVal name As String) As ConfigurationSection
Set appConfig = findConfig(name)
End Property

Public Property Get ConfigurationFile() As ConfigurationFile
Set ConfigurationFile = mConfigFile
End Property

Public Property Get Dirty() As Boolean
If Not mConfigFile Is Nothing Then Dirty = mConfigFile.Dirty
End Property

Public Property Get currentAppConfig() As ConfigurationSection
Set currentAppConfig = mCurrAppConfig
End Property

Public Property Get defaultAppConfig() As ConfigurationSection
Set defaultAppConfig = mDefaultAppConfig
End Property

Public Property Get firstAppConfig() As ConfigurationSection
Dim appConfig As ConfigurationSection

For Each appConfig In mAppConfigs
    Exit For
Next

Set firstAppConfig = appConfig

End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub clearCurrent()
Set mCurrAppConfig = Nothing
End Sub

Public Function addNew( _
                Optional ByVal configName As String) As ConfigurationSection
Dim i As Long

If configName = "" Then
    configName = NewConfigNameStub
    Do While nameAlreadyInUse(configName)
        i = i + 1
        configName = NewConfigNameStub & i
    Loop
End If

Set mCurrAppConfig = AddAppInstanceConfig(mConfigFile, _
                                    configName, _
                                    includeDefaultStudyLibrary:=False, _
                                    setAsDefault:=False, _
                                    TWSServer:=TWSServer, _
                                    TWSPort:=TWSPort, _
                                    clientId:=TWSClientId, _
                                    connectionRetryIntervalSecs:=TWSConnectRetryInterval, _
                                    TickfilesPath:=TickfilesPath)
mCurrAppConfig.addConfigurationSection ConfigSectionCollectionControl
mCurrAppConfig.addConfigurationSection ConfigSectionContracts

mConfigNames.Add configName, configName

Set addNew = mCurrAppConfig
End Function

Public Sub deleteCurrent()
RemoveAppInstanceConfig mConfigFile, mCurrAppConfig.InstanceQualifier
If mCurrAppConfig Is mDefaultAppConfig Then Set mDefaultAppConfig = Nothing
mConfigNames.Remove mCurrAppConfig.InstanceQualifier
Set mCurrAppConfig = Nothing
End Sub

Public Function initialise( _
                ByVal configFilename As String) As Boolean
Dim baseConfigFile As TWUtilities30.ConfigFile
Dim appConfig As ConfigurationSection

mConfigFilename = configFilename

On Error Resume Next
Set baseConfigFile = LoadXMLConfigurationFile(mConfigFilename)
On Error GoTo 0
If baseConfigFile Is Nothing Then
    gLogger.Log LogLevelNormal, "No configuration exists - creating skeleton configuration file"
    Set baseConfigFile = CreateXMLConfigurationFile(App.ProductName, ConfigFileVersion)
    Set mConfigFile = CreateConfigurationFile(baseConfigFile, mConfigFilename)
    InitialiseConfigFile mConfigFile
    Set appConfig = AddAppInstanceConfig(mConfigFile, _
                        DefaultAppInstanceConfigName, _
                        includeDefaultStudyLibrary:=False, _
                        setAsDefault:=True, _
                        TWSServer:=TWSServer, _
                        TWSPort:=TWSPort, _
                        clientId:=TWSClientId, _
                        connectionRetryIntervalSecs:=TWSConnectRetryInterval, _
                        TickfilesPath:=TickfilesPath)
    appConfig.addConfigurationSection ConfigSectionCollectionControl
    appConfig.addConfigurationSection ConfigSectionContracts

    mConfigNames.Add DefaultAppInstanceConfigName, DefaultAppInstanceConfigName

Else
    Set mConfigFile = CreateConfigurationFile(baseConfigFile, _
                                            mConfigFilename)
    If mConfigFile.applicationName <> App.ProductName Or _
        mConfigFile.fileVersion <> ConfigFileVersion Or _
        Not IsValidConfigurationFile(mConfigFile) _
    Then
        gLogger.Log LogLevelNormal, "The configuration file is not the correct format for this program"
        Exit Function
    End If
End If

mConfigFile.AddChangeListener Me

Set mAppConfigs = mConfigFile.GetConfigurationSection("/" & ConfigNameAppConfigs)

Set mDefaultAppConfig = GetDefaultAppInstanceConfig(mConfigFile)

For Each appConfig In mAppConfigs
    mConfigNames.Add appConfig.InstanceQualifier, appConfig.InstanceQualifier
Next

initialise = True
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mAppConfigs.NewEnum
End Function

Public Function renameCurrent( _
                ByVal newName As String) As Boolean
If newName = "" Then Exit Function

If newName = mCurrAppConfig.InstanceQualifier Then Exit Function

If nameAlreadyInUse(newName) Then Exit Function

mConfigNames.Remove mCurrAppConfig.InstanceQualifier
mCurrAppConfig.InstanceQualifier = newName
mConfigNames.Add newName, newName
renameCurrent = True
End Function

Public Sub saveConfigFile( _
                Optional ByVal filename As String)
If filename <> "" Then
    mConfigFilename = filename
End If
mConfigFile.save mConfigFilename
End Sub

Public Sub setCurrent( _
                ByVal cs As ConfigurationSection)
Set mCurrAppConfig = cs
End Sub

Public Sub toggleDefaultConfig()
If mCurrAppConfig Is mDefaultAppConfig Then
    UnsetDefaultAppInstanceConfig mConfigFile
    Set mDefaultAppConfig = Nothing
Else
    SetDefaultAppInstanceConfig mConfigFile, mCurrAppConfig.InstanceQualifier
    
    Set mDefaultAppConfig = mCurrAppConfig
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function findConfig( _
                ByVal name As String) As ConfigurationSection
Set findConfig = mAppConfigs.GetConfigurationSection(ConfigNameAppConfig & "(" & name & ")")
End Function

Private Function nameAlreadyInUse( _
                ByVal name As String) As Boolean
Dim s As String
On Error Resume Next
s = mConfigNames(name)
If s <> "" Then nameAlreadyInUse = True
End Function



