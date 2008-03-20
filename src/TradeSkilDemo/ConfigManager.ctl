VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#49.0#0"; "TradeBuildUI2-6.ocx"
Object = "{6F9EA9CF-F55B-4AFA-8431-9ECC5BED8D43}#23.3#0"; "StudiesUI2-6.ocx"
Begin VB.UserControl ConfigManager 
   BackStyle       =   0  'Transparent
   ClientHeight    =   8325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13890
   DefaultCancel   =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   13890
   Begin VB.CommandButton DeleteConfigButton 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton NewConfigButton 
      Caption         =   "New"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton SaveConfigButton 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3600
      Width           =   735
   End
   Begin StudiesUI26.StudyLibConfigurer StudyLibConfigurer1 
      Height          =   4005
      Left            =   8520
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   7064
   End
   Begin MSComctlLib.TreeView ConfigsTV 
      Height          =   3495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   6165
      _Version        =   393217
      HideSelection   =   0   'False
      LineStyle       =   1
      Style           =   7
      Appearance      =   0
   End
   Begin TradeBuildUI26.SPConfigurer SPConfigurer1 
      Height          =   3975
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   7064
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   11640
      X2              =   12360
      Y1              =   3240
      Y2              =   4920
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   11520
      X2              =   6960
      Y1              =   3240
      Y2              =   4560
   End
   Begin VB.Label Label3 
      Caption         =   "The appropriate control is moved into Box A when editing  service providers or study libraries"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   10560
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   6840
      X2              =   9960
      Y1              =   1560
      Y2              =   2040
   End
   Begin VB.Label Label2 
      Caption         =   "Thix box is the area within which controls for editing config items must fit (Box A)"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   5520
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   10920
      X2              =   10080
      Y1              =   600
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Thix box represents the outline of the control when it is run"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10320
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Shape BoundingRect 
      Height          =   4095
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10095
   End
   Begin VB.Shape Box1 
      BorderColor     =   &H00E7D395&
      Height          =   3975
      Left            =   2520
      Top             =   0
      Width           =   7455
   End
   Begin VB.Menu ConfigTVMenu 
      Caption         =   "Config"
      Visible         =   0   'False
      Begin VB.Menu SetDefaultConfigMenu 
         Caption         =   "Set as default"
         Enabled         =   0   'False
      End
      Begin VB.Menu ConfigSep1Menu 
         Caption         =   "-"
      End
      Begin VB.Menu NewConfigMenu 
         Caption         =   "New"
      End
      Begin VB.Menu RenameConfigMenu 
         Caption         =   "Rename"
         Enabled         =   0   'False
      End
      Begin VB.Menu DeleteConfigMenu 
         Caption         =   "Delete"
         Enabled         =   0   'False
      End
      Begin VB.Menu ConfigSep2Menu 
         Caption         =   "-"
      End
      Begin VB.Menu SaveConfigMenu 
         Caption         =   "Save changes"
         Enabled         =   0   'False
      End
   End
End
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

Event SelectedItemChanged()
Event ConfigFileInvalid()

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

Private Const AttributeNameAppConfigName    As String = "Name"
Private Const AttributeNameAppConfigDefault As String = "Default"

Private Const ConfigFileVersion             As String = "1.0"

Private Const ConfigNameAppConfig           As String = "AppConfig"
Private Const ConfigNameAppConfigs          As String = "AppConfigs"
Private Const ConfigNameTradeBuild          As String = "TradeBuild"

Private Const ConfigNodeServiceProviders    As String = "Service Providers"
Private Const ConfigNodeStudyLibraries      As String = "Study Libraries"

Private Const NewConfigNameStub             As String = "New config"

'@================================================================================
' Member variables
'@================================================================================

Private mConfigFilename                     As String
Private mConfigFile                         As ConfigFile
Private mAppConfigs                         As ConfigItem

Private mCurrAppConfig                      As ConfigItem
Private mCurrConfigNode                     As Node

Private mSelectedAppConfig                  As ConfigItem

Private mDefaultAppConfig                   As ConfigItem
Private mDefaultConfigNode                  As Node

Private mConfigNames                        As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Set mConfigNames = New Collection
End Sub

Private Sub UserControl_Resize()
UserControl.Width = BoundingRect.Width
UserControl.Height = BoundingRect.Height
End Sub

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change( _
                ev As TWUtilities30.ChangeEvent)
If ev.source Is mConfigFile Then
    Select Case ev.changeType
    Case ConfigChangeTypes.ConfigClean
        SaveConfigButton.Enabled = False
        SaveConfigMenu.Enabled = False
    Case ConfigChangeTypes.ConfigDirty
        SaveConfigButton.Enabled = True
        SaveConfigMenu.Enabled = True
    End Select
End If
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ConfigsTV_AfterLabelEdit( _
                cancel As Integer, _
                NewString As String)

If NewString = "" Then
    cancel = True
    Exit Sub
End If

If NewString = ConfigsTV.SelectedItem.Text Then Exit Sub

If nameAlreadyInUse(NewString) Then
    MsgBox "Configuration name '" & NewString & "' is already in use", vbExclamation, "Error"
    cancel = True
    Exit Sub
End If

renameCurrentConfig NewString

End Sub

Private Sub ConfigsTV_MouseUp( _
                Button As Integer, _
                Shift As Integer, _
                x As Single, _
                y As Single)
                
Dim lNode As Node
If Button = vbRightButton Then
    Set lNode = ConfigsTV.HitTest(x, y)
    If Not lNode Is Nothing Then
        DeleteConfigMenu.Enabled = True
        NewConfigMenu.Enabled = True
        RenameConfigMenu.Enabled = True
        SetDefaultConfigMenu.Enabled = True
        If IsObject(lNode.Tag) Then
            If lNode Is mDefaultConfigNode Then
                SetDefaultConfigMenu.Checked = True
            Else
                SetDefaultConfigMenu.Checked = False
            End If
            PopupMenu ConfigTVMenu, , , , RenameConfigMenu
        End If
    Else
        DeleteConfigMenu.Enabled = False
        NewConfigMenu.Enabled = True
        RenameConfigMenu.Enabled = False
        SetDefaultConfigMenu.Enabled = False
        SetDefaultConfigMenu.Checked = False
        PopupMenu ConfigTVMenu, , , , RenameConfigMenu
    End If
End If
End Sub

Private Sub ConfigsTV_NodeClick( _
                ByVal Node As MSComctlLib.Node)

If IsObject(Node.Tag) Then
    setCurrentConfig Node.Tag, Node
    Set mSelectedAppConfig = Node.Tag
Else
    If Not Node.Parent.Tag Is mCurrAppConfig Then setCurrentConfig Node.Parent.Tag, Node.Parent
    
    If Node.Text = ConfigNodeServiceProviders Then
        SPConfigurer1.Left = Box1.Left
        SPConfigurer1.Top = Box1.Top
        SPConfigurer1.Visible = True
        StudyLibConfigurer1.Visible = False
        showServiceProviderConfigDetails
    Else
        StudyLibConfigurer1.Left = Box1.Left
        StudyLibConfigurer1.Top = Box1.Top
        StudyLibConfigurer1.Visible = True
        SPConfigurer1.Visible = False
        showStudyLibraryConfigDetails
    End If
    DeleteConfigButton.Enabled = False
    
    Set mSelectedAppConfig = Nothing
End If
RaiseEvent SelectedItemChanged
End Sub

Private Sub DeleteConfigButton_Click()
deleteAppConfig
End Sub

Private Sub DeleteConfigMenu_Click()
deleteAppConfig
End Sub

Private Sub NewConfigButton_Click()
newAppConfig
End Sub

Private Sub NewConfigMenu_Click()
newAppConfig
End Sub

Private Sub RenameConfigMenu_Click()
ConfigsTV.StartLabelEdit
End Sub

Private Sub SaveConfigButton_Click()
saveConfigFile
End Sub

Private Sub SaveConfigMenu_Click()
saveConfigFile
End Sub

Private Sub SetDefaultConfigMenu_Click()
toggleDefaultConfig
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get changesPending() As Boolean
If StudyLibConfigurer1.dirty Or SPConfigurer1.dirty Then
    changesPending = True
End If
End Property

Public Property Get dirty() As Boolean
If Not mConfigFile Is Nothing Then dirty = mConfigFile.dirty
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
If StudyLibConfigurer1.dirty Then
    StudyLibConfigurer1.ApplyChanges
End If
If SPConfigurer1.dirty Then
    SPConfigurer1.ApplyChanges
End If

End Sub

Public Sub createNewAppConfig( _
                ByVal configName As String, _
                ByVal includeDefaultServiceProviders As Boolean, _
                ByVal includeDefaultStudyLibrary As Boolean)
Set mCurrAppConfig = mAppConfigs.childItems.AddItem(ConfigNameAppConfig)
mCurrAppConfig.setAttribute AttributeNameAppConfigName, configName
mCurrAppConfig.setAttribute AttributeNameAppConfigDefault, "False"
mCurrAppConfig.childItems.AddItem ConfigNameTradeBuild

If includeDefaultServiceProviders Then
    SPConfigurer1.setDefaultServiceProviders mCurrAppConfig.childItems.Item(ConfigNameTradeBuild), _
                                            PermittedServiceProviders.SPBroker Or _
                                            PermittedServiceProviders.SPHistoricalDataInput Or _
                                            PermittedServiceProviders.SPPrimaryContractData Or _
                                            PermittedServiceProviders.SPRealtimeData Or _
                                            PermittedServiceProviders.SPTickfileInput
End If
If includeDefaultStudyLibrary Then
    StudyLibConfigurer1.setDefaultStudyLibrary mCurrAppConfig.childItems.Item(ConfigNameTradeBuild)
End If

Set mCurrConfigNode = addConfigNode(mCurrAppConfig)
mCurrConfigNode.Expanded = True
ConfigsTV.SelectedItem = mCurrConfigNode
ConfigsTV_NodeClick ConfigsTV.SelectedItem
End Sub

Public Function initialise( _
                ByVal configFilename As String, _
                ByVal AppName As String) As Boolean
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
        RaiseEvent ConfigFileInvalid
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

Public Sub saveConfigFile( _
                Optional ByVal filename As String)
If filename <> "" Then
    mConfigFilename = filename
End If
mConfigFile.save mConfigFilename
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
End Function

Private Sub deleteAppConfig()
If MsgBox("Do you want to delete this configuration?" & vbCrLf & _
        "If you click Yes, all data for this configuration will be removed from the configuration file", _
        vbYesNo Or vbQuestion, _
        "Attention!") = vbYes Then
    mAppConfigs.childItems.Remove mCurrAppConfig
    ConfigsTV.Nodes.Remove ConfigsTV.SelectedItem.index
    If mCurrAppConfig Is mDefaultAppConfig Then Set mDefaultAppConfig = Nothing
    Set mCurrAppConfig = Nothing
    If mCurrConfigNode Is mDefaultConfigNode Then Set mDefaultConfigNode = Nothing
    Set mCurrConfigNode = Nothing
    
End If
End Sub

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

Private Sub newAppConfig()
Dim name As String
Dim i As Long
name = NewConfigNameStub
Do While nameAlreadyInUse(name)
    i = i + 1
    name = NewConfigNameStub & i
Loop
createNewAppConfig name, False, False
End Sub

Private Sub renameCurrentConfig( _
                ByVal newName As String)
mConfigNames.Remove mCurrAppConfig.getAttribute(AttributeNameAppConfigName)
mCurrAppConfig.setAttribute AttributeNameAppConfigName, newName
mConfigNames.Add newName, newName
End Sub

Private Sub setCurrentConfig( _
                ByVal ci As ConfigItem, _
                ByVal lNode As Node)
Set mCurrAppConfig = ci
Set mCurrConfigNode = lNode

SPConfigurer1.Visible = False
StudyLibConfigurer1.Visible = False
DeleteConfigButton.Enabled = True
End Sub

Private Sub showServiceProviderConfigDetails()
SPConfigurer1.initialise mCurrAppConfig.childItems.Item(ConfigNameTradeBuild), _
                                        PermittedServiceProviders.SPRealtimeData Or _
                                        PermittedServiceProviders.SPPrimaryContractData Or _
                                        PermittedServiceProviders.SPSecondaryContractData Or _
                                        PermittedServiceProviders.SPBroker Or _
                                        PermittedServiceProviders.SPHistoricalDataInput Or _
                                        PermittedServiceProviders.SPTickfileInput
End Sub

Private Sub showStudyLibraryConfigDetails()
StudyLibConfigurer1.initialise mCurrAppConfig.childItems.Item(ConfigNameTradeBuild)
End Sub

Private Sub toggleDefaultConfig()
If mCurrAppConfig Is mDefaultAppConfig Then
    mCurrAppConfig.setAttribute AttributeNameAppConfigDefault, "False"
    mDefaultConfigNode.Bold = False
    Set mDefaultAppConfig = Nothing
    Set mDefaultConfigNode = Nothing
Else
    If Not mDefaultAppConfig Is Nothing Then
        mDefaultAppConfig.setAttribute AttributeNameAppConfigDefault, "False"
        mDefaultConfigNode.Bold = False
    End If
    
    mCurrAppConfig.setAttribute AttributeNameAppConfigDefault, "True"
    Set mDefaultAppConfig = mCurrAppConfig
    Set mDefaultConfigNode = mCurrConfigNode
    mDefaultConfigNode.Bold = True
End If
End Sub



