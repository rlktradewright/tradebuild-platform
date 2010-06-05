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

Private mConfigStorename                     As String
Private mConfigStore                         As ConfigurationStore
Private mAppConfigs                         As ConfigurationSection

Private mCurrAppConfig                      As ConfigurationSection

Private mDefaultAppConfig                   As ConfigurationSection

Private mConfigNames                        As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mConfigNames = New Collection

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change( _
                ev As TWUtilities30.ChangeEvent)
Const ProcName As String = "ChangeListener_Change"
On Error GoTo Err

If ev.Source Is mConfigStore Then
    Select Case ev.changeType
    Case ConfigChangeTypes.ConfigClean
        RaiseEvent Clean
    Case ConfigChangeTypes.ConfigDirty
        RaiseEvent Dirty
    End Select
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get appConfig( _
                ByVal name As String) As ConfigurationSection
Const ProcName As String = "appConfig"
On Error GoTo Err

Set appConfig = findConfig(name)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ConfigurationFile() As ConfigurationStore
Set ConfigurationFile = mConfigStore
End Property

Public Property Get Dirty() As Boolean
Const ProcName As String = "Dirty"
On Error GoTo Err

If Not mConfigStore Is Nothing Then Dirty = mConfigStore.Dirty

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get currentAppConfig() As ConfigurationSection
Const ProcName As String = "currentAppConfig"
On Error GoTo Err

Set currentAppConfig = mCurrAppConfig

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get defaultAppConfig() As ConfigurationSection
Const ProcName As String = "defaultAppConfig"
On Error GoTo Err

Set defaultAppConfig = mDefaultAppConfig

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get firstAppConfig() As ConfigurationSection
Dim appConfig As ConfigurationSection

Const ProcName As String = "firstAppConfig"
On Error GoTo Err

For Each appConfig In mAppConfigs
    Exit For
Next

Set firstAppConfig = appConfig

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub clearCurrent()
Const ProcName As String = "clearCurrent"
On Error GoTo Err

Set mCurrAppConfig = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function addNew( _
                Optional ByVal configName As String) As ConfigurationSection
Dim i As Long

Const ProcName As String = "addNew"
On Error GoTo Err

If configName = "" Then
    configName = NewConfigNameStub
    Do While nameAlreadyInUse(configName)
        i = i + 1
        configName = NewConfigNameStub & i
    Loop
End If

Set mCurrAppConfig = AddAppInstanceConfig(mConfigStore, _
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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub deleteCurrent()
Const ProcName As String = "deleteCurrent"
On Error GoTo Err

RemoveAppInstanceConfig mConfigStore, mCurrAppConfig.InstanceQualifier
If mCurrAppConfig Is mDefaultAppConfig Then Set mDefaultAppConfig = Nothing
mConfigNames.Remove mCurrAppConfig.InstanceQualifier
Set mCurrAppConfig = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function initialise( _
                ByVal configFilename As String) As Boolean
Dim baseConfigFile As ConfigStoreProvider
Dim appConfig As ConfigurationSection

Const ProcName As String = "initialise"
On Error GoTo Err

mConfigStorename = configFilename

On Error Resume Next
Set baseConfigFile = LoadConfigProviderFromXMLFile(mConfigStorename)
On Error GoTo Err
If baseConfigFile Is Nothing Then
    LogMessage "No configuration exists - creating skeleton configuration file"
    Set baseConfigFile = CreateXMLConfigurationProvider(App.ProductName, ConfigFileVersion)
    Set mConfigStore = CreateConfigurationStore(baseConfigFile, mConfigStorename)
    InitialiseConfigFile mConfigStore
    Set appConfig = AddAppInstanceConfig(mConfigStore, _
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
    Set mConfigStore = CreateConfigurationStore(baseConfigFile, _
                                            mConfigStorename)
    If mConfigStore.ApplicationName <> App.ProductName Or _
        mConfigStore.fileVersion <> ConfigFileVersion Or _
        Not IsValidConfigurationFile(mConfigStore) _
    Then
        LogMessage "The configuration file is not the correct format for this program"
        Exit Function
    End If
End If

mConfigStore.AddChangeListener Me

Set mAppConfigs = mConfigStore.GetConfigurationSection("/" & ConfigNameAppConfigs)

Set mDefaultAppConfig = GetDefaultAppInstanceConfig(mConfigStore)

For Each appConfig In mAppConfigs
    mConfigNames.Add appConfig.InstanceQualifier, appConfig.InstanceQualifier
Next

initialise = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
On Error GoTo Err

Set NewEnum = mAppConfigs.NewEnum

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function renameCurrent( _
                ByVal newName As String) As Boolean
Const ProcName As String = "renameCurrent"
On Error GoTo Err

If newName = "" Then Exit Function

If newName = mCurrAppConfig.InstanceQualifier Then Exit Function

If nameAlreadyInUse(newName) Then Exit Function

mConfigNames.Remove mCurrAppConfig.InstanceQualifier
mCurrAppConfig.InstanceQualifier = newName
mConfigNames.Add newName, newName
renameCurrent = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub saveConfigFile( _
                Optional ByVal filename As String)
Const ProcName As String = "saveConfigFile"
On Error GoTo Err

If filename <> "" Then
    mConfigStorename = filename
End If
mConfigStore.Save mConfigStorename

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub setCurrent( _
                ByVal cs As ConfigurationSection)
Const ProcName As String = "setCurrent"
On Error GoTo Err

Set mCurrAppConfig = cs

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub toggleDefaultConfig()
Const ProcName As String = "toggleDefaultConfig"
On Error GoTo Err

If mCurrAppConfig Is mDefaultAppConfig Then
    UnsetDefaultAppInstanceConfig mConfigStore
    Set mDefaultAppConfig = Nothing
Else
    SetDefaultAppInstanceConfig mConfigStore, mCurrAppConfig.InstanceQualifier
    
    Set mDefaultAppConfig = mCurrAppConfig
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function findConfig( _
                ByVal name As String) As ConfigurationSection
Const ProcName As String = "findConfig"
On Error GoTo Err

Set findConfig = mAppConfigs.GetConfigurationSection(ConfigNameAppConfig & "(" & name & ")")

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function nameAlreadyInUse( _
                ByVal name As String) As Boolean
Dim s As String
Const ProcName As String = "nameAlreadyInUse"
On Error GoTo Err

On Error Resume Next
s = mConfigNames(name)
If s <> "" Then nameAlreadyInUse = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function



