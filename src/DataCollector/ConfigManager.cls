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

Private Const NewConfigNameStub             As String = "New config"

'@================================================================================
' Member variables
'@================================================================================

Private mConfigFilename                     As String
Private mConfigFile                         As ConfigFile
Private mAppConfigs                         As ConfigItem

Private mCurrAppConfig                      As ConfigItem

Private mDefaultAppConfig                   As ConfigItem
Private mDefaultConfigNode                  As Node

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
                ev As twutilities30.ChangeEvent)
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
                ByVal name As String) As ConfigItem
Set appConfig = findConfig(name)
End Property

Public Property Get Dirty() As Boolean
If Not mConfigFile Is Nothing Then Dirty = mConfigFile.Dirty
End Property

Public Property Get currentAppConfig() As ConfigItem
Set currentAppConfig = mCurrAppConfig
End Property

Public Property Get defaultAppConfig() As ConfigItem
Set defaultAppConfig = mDefaultAppConfig
End Property

Public Property Get firstAppConfig() As ConfigItem
Dim appConfig As ConfigItem

For Each appConfig In mAppConfigs.childItems
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
                Optional ByVal configName As String) As ConfigItem
Dim i As Long

If configName = "" Then
    configName = NewConfigNameStub
    Do While nameAlreadyInUse(configName)
        i = i + 1
        configName = NewConfigNameStub & i
    Loop
End If

Set mCurrAppConfig = mAppConfigs.childItems.AddItem(ConfigNameAppConfig)
mCurrAppConfig.setAttribute AttributeNameAppConfigName, configName
mCurrAppConfig.setAttribute AttributeNameAppConfigDefault, "False"
mCurrAppConfig.childItems.AddItem ConfigNameTradeBuild
mCurrAppConfig.childItems.AddItem ConfigNameCollectionControl
mCurrAppConfig.childItems.AddItem ConfigNameContracts

Set addNew = mCurrAppConfig
End Function

Public Sub deleteCurrent()
mAppConfigs.childItems.Remove mCurrAppConfig
If mCurrAppConfig Is mDefaultAppConfig Then Set mDefaultAppConfig = Nothing
mConfigNames.Remove mCurrAppConfig.getAttribute(AttributeNameAppConfigName)
Set mCurrAppConfig = Nothing
End Sub

Public Function initialise( _
                ByVal configFilename As String) As Boolean
Dim appConfig As ConfigItem
Dim isDefault As Boolean
Dim index As Long
Dim newnode As Node

mConfigFilename = configFilename

On Error Resume Next
Set mConfigFile = LoadXMLConfigurationFile(mConfigFilename)
On Error GoTo 0
If mConfigFile Is Nothing Then
    gLogger.Log LogLevelNormal, "No configuration exists - creating skeleton configuration file"
    Set mConfigFile = CreateXMLConfigurationFile(AppName, ConfigFileVersion)
Else
    If mConfigFile.applicationName <> AppName Or _
        mConfigFile.applicationVersion <> ConfigFileVersion _
    Then
        gLogger.Log LogLevelNormal, "The configuration file is not the correct format for this program"
        Exit Function
    End If
End If

mConfigFile.addChangeListener Me

On Error Resume Next
Set mAppConfigs = mConfigFile.rootItem.childItems.Item(ConfigNameAppConfigs)
On Error GoTo 0

If mAppConfigs Is Nothing Then
    Set mAppConfigs = mConfigFile.rootItem.childItems.AddItem(ConfigNameAppConfigs)
End If

For Each appConfig In mAppConfigs.childItems
    mConfigNames.Add appConfig.getAttribute(AttributeNameAppConfigName), appConfig.getAttribute(AttributeNameAppConfigName)
    isDefault = False
    On Error Resume Next
    isDefault = (UCase$(appConfig.getAttribute(AttributeNameAppConfigDefault)) = "TRUE")
    On Error GoTo 0
    If isDefault Then Set mDefaultAppConfig = appConfig
    index = index + 1
Next

initialise = True
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mAppConfigs.childItems.NewEnum
End Function

Public Function renameCurrent( _
                ByVal newName As String) As Boolean
If newName = "" Then Exit Function

If newName = mCurrAppConfig.getAttribute(AttributeNameAppConfigName) Then Exit Function

If nameAlreadyInUse(newName) Then Exit Function

mConfigNames.Remove mCurrAppConfig.getAttribute(AttributeNameAppConfigName)
mCurrAppConfig.setAttribute AttributeNameAppConfigName, newName
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
                ByVal ci As ConfigItem)
Set mCurrAppConfig = ci
End Sub

Public Sub toggleDefaultConfig()
If mCurrAppConfig Is mDefaultAppConfig Then
    mCurrAppConfig.setAttribute AttributeNameAppConfigDefault, "False"
    Set mDefaultAppConfig = Nothing
Else
    If Not mDefaultAppConfig Is Nothing Then
        mDefaultAppConfig.setAttribute AttributeNameAppConfigDefault, "False"
    End If
    
    mCurrAppConfig.setAttribute AttributeNameAppConfigDefault, "True"
    Set mDefaultAppConfig = mCurrAppConfig
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function findConfig( _
                ByVal name As String) As ConfigItem
Dim appConfig As ConfigItem

For Each appConfig In mAppConfigs.childItems
    If UCase$(appConfig.getAttribute(AttributeNameAppConfigName)) = UCase$(name) Then
        Set findConfig = appConfig
        Exit Function
    End If
Next

End Function

Private Function nameAlreadyInUse( _
                ByVal name As String) As Boolean
Dim s As String
On Error Resume Next
s = mConfigNames(name)
If s <> "" Then nameAlreadyInUse = True
End Function



