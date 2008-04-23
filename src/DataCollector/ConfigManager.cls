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

Private Const AttributeNameAppConfigDefault As String = "Default"
Private Const AttributeNameAppConfigName    As String = "Name"
Private Const AttributeNameWriteTickData    As String = "WriteTickData"
Private Const AttributeNameWriteBarData     As String = "WriteBarData"


Private Const ConfigFileVersion             As String = "1.0"

Private Const ConfigNameAppConfig           As String = "AppConfig"
Private Const ConfigNameAppConfigs          As String = "AppConfigs"
Private Const ConfigNameTradeBuild          As String = "TradeBuild"
Private Const ConfigNameCollectionControl   As String = "CollectionControl"
Private Const ConfigNameContract            As String = "Contract"
Private Const ConfigNameContracts           As String = "Contracts"
Private Const ConfigNameTickfilePath        As String = "TickfilePath"

Private Const ConfigNodeContractSpecs       As String = "Contract Specifications"
Private Const ConfigNodeServiceProviders    As String = "Service Providers"
Private Const ConfigNodeStudyLibraries      As String = "Study Libraries"
Private Const ConfigNodeParameters          As String = "Parameters"

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

Public Property Get Dirty() As Boolean
If Not mConfigFile Is Nothing Then Dirty = mConfigFile.Dirty
End Property

Public Property Get appConfig( _
                ByVal name As String) As ConfigItem
Set appConfig = findConfig(name)
End Property

Public Property Get firstAppConfig() As ConfigItem
Dim appConfig As ConfigItem

For Each appConfig In mAppConfigs.childItems
    Exit For
Next

Set firstAppConfig = appConfig

End Property

Public Property Get selectedAppConfig() As ConfigItem
Set selectedAppConfig = mSelectedAppConfig
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub appyPendingChanges()
If StudyLibConfigurer1.Dirty Then
    StudyLibConfigurer1.ApplyChanges
End If
If SPConfigurer1.Dirty Then
    SPConfigurer1.ApplyChanges
End If

End Sub

Public Sub clearCurrent()
Set mCurrAppConfig = Nothing
End Sub

Public Function addNew( _
                Optional ByVal configName As String, _
                Optional ByVal includeDefaultServiceProviders As Boolean, _
                Optional ByVal includeDefaultStudyLibrary As Boolean) As ConfigItem
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

If includeDefaultServiceProviders Then
    SPConfigurer1.setDefaultServiceProviders mCurrAppConfig.childItems.Item(ConfigNameTradeBuild), _
                                            PermittedServiceProviders.SPHistoricalDataInput Or _
                                            PermittedServiceProviders.SPHistoricalDataOutput Or _
                                            PermittedServiceProviders.SPPrimaryContractData Or _
                                            PermittedServiceProviders.SPRealtimeData Or _
                                            PermittedServiceProviders.SPTickfileOutput
End If
If includeDefaultStudyLibrary Then
    StudyLibConfigurer1.setDefaultStudyLibrary mCurrAppConfig.childItems.Item(ConfigNameTradeBuild)
End If


Set addNew = mCurrAppConfig
End Function

Private Sub deleteCurrent()
mAppConfigs.childItems.Remove mCurrAppConfig
If mCurrAppConfig Is mDefaultAppConfig Then Set mDefaultAppConfig = Nothing
Set mCurrAppConfig = Nothing
End Sub

Public Function initialise( _
                ByVal configFilename As String, _
                ByVal applicationName As String) As Boolean
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
    Set mConfigFile = CreateXMLConfigurationFile(applicationName, ConfigFileVersion)
Else
    If mConfigFile.applicationName <> applicationName Or _
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
    isDefault = False
    On Error Resume Next
    isDefault = (UCase$(appConfig.getAttribute(AttributeNameAppConfigDefault)) = "TRUE")
    On Error GoTo 0
    Set newnode = addConfigNode(appConfig)
    If isDefault Then
        newnode.Bold = True
        Set mDefaultAppConfig = appConfig
        Set mDefaultConfigNode = newnode
    End If
    index = index + 1
Next

If Not mDefaultConfigNode Is Nothing Then
    ConfigsTV.SelectedItem = mDefaultConfigNode
ElseIf ConfigsTV.Nodes.Count > 0 Then
    ConfigsTV.SelectedItem = ConfigsTV.Nodes(1)
End If
If Not ConfigsTV.SelectedItem Is Nothing Then ConfigsTV_NodeClick ConfigsTV.SelectedItem
initialise = True
End Function

Public Function renameCurrent( _
                ByVal newName As String) As Boolean
If newName = "" Then Exit Sub

If newName = mCurrAppConfig.getAttribute(AttributeNameAppConfigName) Then Exit Sub

If nameAlreadyInUse(newName) Then Exit Sub

mConfigNames.Remove mCurrAppConfig.getAttribute(AttributeNameAppConfigName)
mCurrAppConfig.setAttribute AttributeNameAppConfigName, newName
mConfigNames.Add newName, newName
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

Private Function addConfigNode( _
                ByVal appConfig As ConfigItem) As Node
Dim name As String
name = appConfig.getAttribute(AttributeNameAppConfigName)
Set addConfigNode = ConfigsTV.Nodes.Add(, , name, name)
Set addConfigNode.Tag = appConfig
mConfigNames.Add name, name
ConfigsTV.Nodes.Add addConfigNode, tvwChild, , ConfigNodeServiceProviders
ConfigsTV.Nodes.Add addConfigNode, tvwChild, , ConfigNodeStudyLibraries
ConfigsTV.Nodes.Add addConfigNode, tvwChild, , ConfigNodeParameters
ConfigsTV.Nodes.Add addConfigNode, tvwChild, , ConfigNodeContractSpecs
End Function

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

Private Sub hideConfigControls()
SPConfigurer1.Visible = False
StudyLibConfigurer1.Visible = False
ParametersPicture.Visible = False
ContractsConfigurer1.Visible = False
End Sub

Private Function nameAlreadyInUse( _
                ByVal name As String) As Boolean
Dim s As String
On Error Resume Next
s = mConfigNames(name)
If s <> "" Then nameAlreadyInUse = True
End Function

Private Sub showContractSpecsConfigDetails()

hideConfigControls

ContractsConfigurer1.initialise mCurrAppConfig.childItems.Item(ConfigNameContracts)

ContractsConfigurer1.Left = Box1.Left
ContractsConfigurer1.Top = Box1.Top
ContractsConfigurer1.Visible = True
End Sub

Private Sub showParametersConfigDetails()
Dim ci As ConfigItem

hideConfigControls

Set ci = mCurrAppConfig.childItems.Item(ConfigNameCollectionControl)

WriteBarDataCheck.value = IIf(ci.getDefaultableAttribute(AttributeNameWriteBarData, "False") = "True", vbChecked, vbUnchecked)
WriteTickDataCheck.value = IIf(ci.getDefaultableAttribute(AttributeNameWriteTickData, "False") = "True", vbChecked, vbUnchecked)

On Error Resume Next
Set ci = mCurrAppConfig.childItems.Item(ConfigNameTickfilePath)
On Error GoTo 0
If Not ci Is Nothing Then
    TickfilePathText = ci.value
Else
    TickfilePathText = ""
End If

ParametersPicture.Left = Box1.Left
ParametersPicture.Top = Box1.Top
ParametersPicture.Visible = True
End Sub

Private Sub showServiceProviderConfigDetails()
hideConfigControls
SPConfigurer1.Left = Box1.Left
SPConfigurer1.Top = Box1.Top
SPConfigurer1.initialise mCurrAppConfig.childItems.Item(ConfigNameTradeBuild), _
                                        PermittedServiceProviders.SPRealtimeData Or _
                                        PermittedServiceProviders.SPPrimaryContractData Or _
                                        PermittedServiceProviders.SPHistoricalDataInput Or _
                                        PermittedServiceProviders.SPHistoricalDataOutput Or _
                                        PermittedServiceProviders.SPTickfileOutput
SPConfigurer1.Visible = True
End Sub

Private Sub showStudyLibraryConfigDetails()
hideConfigControls
StudyLibConfigurer1.Left = Box1.Left
StudyLibConfigurer1.Top = Box1.Top
StudyLibConfigurer1.initialise mCurrAppConfig.childItems.Item(ConfigNameTradeBuild)
StudyLibConfigurer1.Visible = True
End Sub





