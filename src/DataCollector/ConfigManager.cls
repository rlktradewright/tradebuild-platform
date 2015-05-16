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

Private Const ModuleName                    As String = "ConfigManager"

Private Const ConfigNameAppConfig           As String = "AppConfig"
Private Const ConfigNameAppConfigs          As String = "AppConfigs"

Private Const NewConfigNameStub             As String = "New config"

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
                ev As ChangeEventData)
Const ProcName As String = "ChangeListener_Change"
On Error GoTo Err

If ev.Source Is mConfigStore Then
    Select Case ev.ChangeType
    Case ConfigChangeTypes.ConfigClean
        RaiseEvent Clean
    Case ConfigChangeTypes.ConfigDirty
        RaiseEvent Dirty
    End Select
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get AppConfig( _
                ByVal pName As String) As ConfigurationSection
Const ProcName As String = "AppConfig"
On Error GoTo Err

Set AppConfig = findConfig(pName)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentAppConfig() As ConfigurationSection
Const ProcName As String = "CurrentAppConfig"
On Error GoTo Err

Set CurrentAppConfig = mCurrAppConfig

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultAppConfig() As ConfigurationSection
Const ProcName As String = "DefaultAppConfig"
On Error GoTo Err

Set DefaultAppConfig = mDefaultAppConfig

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FirstAppConfig() As ConfigurationSection
Const ProcName As String = "FirstAppConfig"
On Error GoTo Err

Dim lAppConfig As ConfigurationSection
For Each lAppConfig In mAppConfigs
    Exit For
Next

Set FirstAppConfig = lAppConfig

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ClearCurrent()
Const ProcName As String = "ClearCurrent"
On Error GoTo Err

Set mCurrAppConfig = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function AddNew( _
                Optional ByVal configName As String) As ConfigurationSection
Const ProcName As String = "AddNew"
On Error GoTo Err

If configName = "" Then
    configName = NewConfigNameStub
    Dim i As Long
    Do While nameAlreadyInUse(configName)
        i = i + 1
        configName = NewConfigNameStub & i
    Loop
End If

Set mCurrAppConfig = AddAppInstanceConfig(mConfigStore, _
                                    configName, _
                                    0, _
                                    pTWSServer:=TWSServer, _
                                    pTWSPort:=TWSPort, _
                                    pTwsClientId:=TWSClientId, _
                                    pTwsConnectionRetryIntervalSecs:=TWSConnectRetryInterval, _
                                    pTickfilesPath:=TickfilesPath)
mCurrAppConfig.addConfigurationSection ConfigSectionCollectionControl
mCurrAppConfig.addConfigurationSection ConfigSectionContracts

mConfigNames.Add configName, configName

Set AddNew = mCurrAppConfig

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub DeleteCurrent()
Const ProcName As String = "DeleteCurrent"
On Error GoTo Err

RemoveAppInstanceConfig mConfigStore, mCurrAppConfig.InstanceQualifier
If mCurrAppConfig Is mDefaultAppConfig Then Set mDefaultAppConfig = Nothing
mConfigNames.Remove mCurrAppConfig.InstanceQualifier
Set mCurrAppConfig = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Initialise( _
                ByVal configFilename As String) As Boolean
Const ProcName As String = "Initialise"
On Error GoTo Err

mConfigStorename = configFilename

On Error Resume Next
Dim lBaseConfigFile As ConfigStoreProvider
Set lBaseConfigFile = LoadConfigProviderFromXMLFile(mConfigStorename)
On Error GoTo Err

If lBaseConfigFile Is Nothing Then Exit Function
    
Set mConfigStore = CreateConfigurationStore(lBaseConfigFile, _
                                        mConfigStorename)
If mConfigStore.ApplicationName <> App.ProductName Or _
    mConfigStore.fileVersion <> ConfigFileVersion Or _
    Not IsValidConfigurationFile(mConfigStore) _
Then
    LogMessage "The configuration file is not the correct format for this program"
    Exit Function
End If

mConfigStore.AddChangeListener Me

Set mAppConfigs = mConfigStore.GetConfigurationSection("/" & ConfigNameAppConfigs)

Set mDefaultAppConfig = GetDefaultAppInstanceConfig(mConfigStore)

Dim lAppConfig As ConfigurationSection
For Each lAppConfig In mAppConfigs
    mConfigNames.Add lAppConfig.InstanceQualifier, lAppConfig.InstanceQualifier
Next

Initialise = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
On Error GoTo Err

Set NewEnum = mAppConfigs.NewEnum

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function RenameCurrent( _
                ByVal newName As String) As Boolean
Const ProcName As String = "RenameCurrent"
On Error GoTo Err

If newName = "" Then Exit Function

If newName = mCurrAppConfig.InstanceQualifier Then Exit Function

If nameAlreadyInUse(newName) Then Exit Function

mConfigNames.Remove mCurrAppConfig.InstanceQualifier
mCurrAppConfig.InstanceQualifier = newName
mConfigNames.Add newName, newName
RenameCurrent = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub SaveConfigFile( _
                Optional ByVal filename As String)
Const ProcName As String = "SaveConfigFile"
On Error GoTo Err

If filename <> "" Then
    mConfigStorename = filename
End If
mConfigStore.Save mConfigStorename

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetCurrent( _
                ByVal cs As ConfigurationSection)
Const ProcName As String = "SetCurrent"
On Error GoTo Err

Set mCurrAppConfig = cs

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ToggleDefaultConfig()
Const ProcName As String = "ToggleDefaultConfig"
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function findConfig( _
                ByVal Name As String) As ConfigurationSection
Const ProcName As String = "findConfig"
On Error GoTo Err

Set findConfig = mAppConfigs.GetConfigurationSection(ConfigNameAppConfig & "(" & Name & ")")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function nameAlreadyInUse( _
                ByVal Name As String) As Boolean
Dim s As String
Const ProcName As String = "nameAlreadyInUse"
On Error GoTo Err

On Error Resume Next
s = mConfigNames(Name)
If s <> "" Then nameAlreadyInUse = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



