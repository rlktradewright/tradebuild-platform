VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#207.0#0"; "TradeBuildUI2-6.ocx"
Begin VB.UserControl ConfigViewer 
   BackStyle       =   0  'Transparent
   ClientHeight    =   13740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16680
   DefaultCancel   =   -1  'True
   ScaleHeight     =   13740
   ScaleWidth      =   16680
   Begin DataCollector26.ContractsConfigurer ContractsConfigurer1 
      Height          =   4005
      Left            =   8520
      TabIndex        =   11
      Top             =   4320
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   7064
   End
   Begin VB.PictureBox ParametersPicture 
      BorderStyle     =   0  'None
      Height          =   4005
      Left            =   120
      ScaleHeight     =   4005
      ScaleWidth      =   7500
      TabIndex        =   10
      Top             =   8520
      Width           =   7500
      Begin VB.CheckBox WriteTickDataCheck 
         Caption         =   "Write tick data"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox WriteBarDataCheck 
         Caption         =   "Write bar data"
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.CommandButton DeleteConfigButton 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton NewConfigButton 
      Caption         =   "New"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton SaveConfigButton 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3600
      Width           =   735
   End
   Begin MSComctlLib.TreeView ConfigsTV 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
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
      TabIndex        =   6
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      Height          =   4005
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
Attribute VB_Name = "ConfigViewer"
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

Private Const ProjectName                   As String = "DataCollector26"
Private Const ModuleName                    As String = "ConfigViewer"

Private Const ConfigFileVersion             As String = "1.1"

Private Const ConfigNameTradeBuild          As String = "TradeBuild"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mConfigManager           As ConfigManager
Attribute mConfigManager.VB_VarHelpID = -1

Private mCurrConfigNode                     As Node

Private mSelectedAppConfig                  As ConfigurationSection

Private mReadOnly                           As Boolean

Private mDefaultConfigNode                  As Node

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Resize()
UserControl.Width = BoundingRect.Width
UserControl.Height = BoundingRect.Height
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ConfigsTV_AfterLabelEdit( _
                Cancel As Integer, _
                NewString As String)

If Not mConfigManager.renameCurrent(NewString) Then
    MsgBox "Configuration name '" & NewString & "' is already in use", vbExclamation, "Error"
    Cancel = True
    Exit Sub
End If

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
        If Not mReadOnly Then DeleteConfigMenu.enabled = True
        If Not mReadOnly Then NewConfigMenu.enabled = True
        If Not mReadOnly Then RenameConfigMenu.enabled = True
        If Not mReadOnly Then SetDefaultConfigMenu.enabled = True
        If IsObject(lNode.Tag) Then
            If lNode Is mDefaultConfigNode Then
                SetDefaultConfigMenu.Checked = True
            Else
                SetDefaultConfigMenu.Checked = False
            End If
            PopupMenu ConfigTVMenu, , , , RenameConfigMenu
        End If
    Else
        DeleteConfigMenu.enabled = False
        If Not mReadOnly Then NewConfigMenu.enabled = True
        RenameConfigMenu.enabled = False
        SetDefaultConfigMenu.enabled = False
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
    If Not Node.Parent.Tag Is mConfigManager.currentAppConfig Then setCurrentConfig Node.Parent.Tag, Node.Parent
    
    If Node.Text = ConfigNodeServiceProviders Then
        showServiceProviderConfigDetails
    ElseIf Node.Text = ConfigNodeParameters Then
        showParametersConfigDetails
    ElseIf Node.Text = ConfigNodeContractSpecs Then
        showContractSpecsConfigDetails
    End If
    DeleteConfigButton.enabled = False
    
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
addConfigNode mConfigManager.addNew
End Sub

Private Sub NewConfigMenu_Click()
addConfigNode (mConfigManager.addNew)
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

Private Sub WriteBarDataCheck_Click()
If mReadOnly Then Exit Sub
mConfigManager.currentAppConfig.SetSetting ConfigSettingWriteBarData, CStr(WriteBarDataCheck.value = vbChecked)
End Sub

Private Sub WriteTickDataCheck_Click()
If mReadOnly Then Exit Sub
mConfigManager.currentAppConfig.SetSetting ConfigSettingWriteTickData, CStr(WriteTickDataCheck.value = vbChecked)
End Sub

'@================================================================================
' mConfigManager Event Handlers
'@================================================================================

Private Sub mConfigManager_Clean()
SaveConfigButton.enabled = False
SaveConfigMenu.enabled = False
End Sub

Private Sub mConfigManager_Dirty()
If Not mReadOnly Then SaveConfigButton.enabled = True
If Not mReadOnly Then SaveConfigMenu.enabled = True
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get changesPending() As Boolean
If SPConfigurer1.Dirty Then
    changesPending = True
End If
End Property

Public Property Get Dirty() As Boolean
Dirty = mConfigManager.Dirty
End Property

Public Property Get appConfig( _
                ByVal name As String) As ConfigurationSection
Set appConfig = mConfigManager.appConfig(name)
End Property

Public Property Get firstAppConfig() As ConfigurationSection
Set firstAppConfig = mConfigManager.firstAppConfig
End Property

Public Property Get selectedAppConfig() As ConfigurationSection
Set selectedAppConfig = mConfigManager.currentAppConfig
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub applyPendingChanges()
If SPConfigurer1.Dirty Then
    SPConfigurer1.ApplyChanges
End If
End Sub

Public Sub createNewAppConfig( _
                ByVal configName As String, _
                ByVal includeDefaultServiceProviders As Boolean, _
                ByVal includeDefaultStudyLibrary As Boolean)
Set mCurrConfigNode = addConfigNode(mConfigManager.addNew)
mCurrConfigNode.Expanded = True
ConfigsTV.SelectedItem = mCurrConfigNode
ConfigsTV_NodeClick ConfigsTV.SelectedItem
End Sub

Public Function initialise( _
                ByVal pconfigManager As ConfigManager, _
                ByVal readonly As Boolean) As Boolean
Dim appConfig As ConfigurationSection
Dim index As Long
Dim newnode As Node

mReadOnly = readonly

Set mConfigManager = pconfigManager

For Each appConfig In mConfigManager
    Set newnode = addConfigNode(appConfig)
    If appConfig Is mConfigManager.defaultAppConfig Then
        newnode.Bold = True
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

If mReadOnly Then disableControls
initialise = True
End Function

Public Sub saveConfigFile( _
                Optional ByVal filename As String)
mConfigManager.saveConfigFile filename
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addConfigNode( _
                ByVal appConfig As ConfigurationSection) As Node
Dim name As String
name = appConfig.InstanceQualifier
Set addConfigNode = ConfigsTV.Nodes.Add(, , name, name)
Set addConfigNode.Tag = appConfig
ConfigsTV.Nodes.Add addConfigNode, tvwChild, , ConfigNodeServiceProviders
ConfigsTV.Nodes.Add addConfigNode, tvwChild, , ConfigNodeParameters
ConfigsTV.Nodes.Add addConfigNode, tvwChild, , ConfigNodeContractSpecs
End Function

Private Sub deleteAppConfig()
If MsgBox("Do you want to delete this configuration?" & vbCrLf & _
        "If you click Yes, all data for this configuration will be removed from the configuration file", _
        vbYesNo Or vbQuestion, _
        "Attention!") = vbYes Then
    mConfigManager.deleteCurrent
    ConfigsTV.Nodes.Remove ConfigsTV.SelectedItem.index
    If mCurrConfigNode Is mDefaultConfigNode Then Set mDefaultConfigNode = Nothing
    Set mCurrConfigNode = Nothing
End If
End Sub

Private Sub disableControls()
DeleteConfigButton.enabled = False
NewConfigButton.enabled = False
SaveConfigButton.enabled = False
End Sub

Private Sub hideConfigControls()
SPConfigurer1.Visible = False
ParametersPicture.Visible = False
ContractsConfigurer1.Visible = False
End Sub

Private Sub setCurrentConfig( _
                ByVal cs As ConfigurationSection, _
                ByVal lNode As Node)
mConfigManager.setCurrent cs
Set mCurrConfigNode = lNode

hideConfigControls
If Not mReadOnly Then DeleteConfigButton.enabled = True
End Sub

Private Sub showContractSpecsConfigDetails()

hideConfigControls

ContractsConfigurer1.initialise mConfigManager.currentAppConfig.GetConfigurationSection(ConfigSectionContracts), _
                                mReadOnly

ContractsConfigurer1.Left = Box1.Left
ContractsConfigurer1.Top = Box1.Top
ContractsConfigurer1.Visible = True
End Sub

Private Sub showParametersConfigDetails()
hideConfigControls

WriteBarDataCheck.value = IIf(mConfigManager.currentAppConfig.GetSetting(ConfigSettingWriteBarData, "False") = "True", vbChecked, vbUnchecked)
WriteTickDataCheck.value = IIf(mConfigManager.currentAppConfig.GetSetting(ConfigSettingWriteTickData, "False") = "True", vbChecked, vbUnchecked)

ParametersPicture.Left = Box1.Left
ParametersPicture.Top = Box1.Top
ParametersPicture.Visible = True
End Sub

Private Sub showServiceProviderConfigDetails()
hideConfigControls
SPConfigurer1.Left = Box1.Left
SPConfigurer1.Top = Box1.Top
SPConfigurer1.initialise mConfigManager.currentAppConfig.GetConfigurationSection(ConfigNameTradeBuild), _
                                        mReadOnly
SPConfigurer1.Visible = True
End Sub

Private Sub toggleDefaultConfig()
If mConfigManager.currentAppConfig Is mConfigManager.defaultAppConfig Then
    mDefaultConfigNode.Bold = False
    Set mDefaultConfigNode = Nothing
Else
    If Not mConfigManager.defaultAppConfig Is Nothing Then
        mDefaultConfigNode.Bold = False
    End If
    
    Set mDefaultConfigNode = mCurrConfigNode
    mDefaultConfigNode.Bold = True
End If
mConfigManager.toggleDefaultConfig
End Sub



