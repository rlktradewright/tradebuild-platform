VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{948AEB4D-03C6-4FAB-ACD2-E61F7B7A0EB3}#128.0#0"; "TradeBuildUI27.ocx"
Object = "{464F646E-C78A-4AAC-AC11-FBC7E41F58BB}#217.0#0"; "StudiesUI27.ocx"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#31.0#0"; "TWControls40.ocx"
Begin VB.UserControl ConfigManager 
   BackStyle       =   0  'Transparent
   ClientHeight    =   8325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16170
   ScaleHeight     =   8325
   ScaleWidth      =   16170
   Begin TWControls40.TWButton NewConfigButton 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   3600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "New"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TWControls40.TWButton DeleteConfigButton 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TradeBuildUI27.SPConfigurer SPConfigurer1 
      Height          =   4005
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   7064
   End
   Begin StudiesUI27.StudyLibConfigurer StudyLibConfigurer1 
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      Height          =   4020
      Left            =   2520
      Top             =   0
      Width           =   7515
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

Implements IThemeable

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

Private Const ModuleName                    As String = "ConfigManager"

Private Const ConfigNodeServiceProviders    As String = "Service Providers"
Private Const ConfigNodeStudyLibraries      As String = "Study Libraries"

Private Const NewConfigNameStub             As String = "New config"

'@================================================================================
' Member variables
'@================================================================================

Private mConfigStore                         As ConfigurationStore
Private mAppConfigs                         As ConfigurationSection

Private mCurrAppConfig                      As ConfigurationSection
Private mCurrConfigNode                     As Node

Private mSelectedAppConfig                  As ConfigurationSection

Private mDefaultAppConfig                   As ConfigurationSection
Private mDefaultConfigNode                  As Node

Private mConfigNames                        As Collection

Private mPermittedServiceProviderRoles      As ServiceProviderRoles
Private mFlags                              As ConfigFlags

Private mTheme                              As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

Set mConfigNames = New Collection

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

UserControl.Width = BoundingRect.Width
UserControl.Height = BoundingRect.Height

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal Value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ConfigsTV_AfterLabelEdit( _
                Cancel As Integer, _
                NewString As String)

Const ProcName As String = "ConfigsTV_AfterLabelEdit"
On Error GoTo Err

If NewString = "" Then
    Cancel = True
    Exit Sub
End If

If NewString = ConfigsTV.SelectedItem.Text Then Exit Sub

If nameAlreadyInUse(NewString) Then
    MsgBox "Configuration name '" & NewString & "' is already in use", vbExclamation, "Error"
    Cancel = True
    Exit Sub
End If

renameCurrentConfig NewString

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ConfigsTV_MouseUp( _
                Button As Integer, _
                Shift As Integer, _
                x As Single, _
                y As Single)
Const ProcName As String = "ConfigsTV_MouseUp"
On Error GoTo Err

If Button = vbRightButton Then
    Dim lNode As Node
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

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ConfigsTV_NodeClick( _
                ByVal Node As MSComctlLib.Node)
Const ProcName As String = "ConfigsTV_NodeClick"
On Error GoTo Err

If IsObject(Node.Tag) Then
    setCurrentConfig Node.Tag, Node
    Set mSelectedAppConfig = Node.Tag
Else
    If Not Node.Parent.Tag Is mCurrAppConfig Then setCurrentConfig Node.Parent.Tag, Node.Parent
    
    If Node.Text = ConfigNodeServiceProviders Then
        showServiceProviderConfigDetails
    Else
        showStudyLibraryConfigDetails
    End If
    DeleteConfigButton.Enabled = False
    
    Set mSelectedAppConfig = Nothing
End If
RaiseEvent SelectedItemChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub DeleteConfigButton_Click()
Const ProcName As String = "DeleteConfigButton_Click"
On Error GoTo Err

deleteAppConfig

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub DeleteConfigMenu_Click()
Const ProcName As String = "DeleteConfigMenu_Click"
On Error GoTo Err

deleteAppConfig

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub NewConfigButton_Click()
Const ProcName As String = "NewConfigButton_Click"
On Error GoTo Err

newAppConfig

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub NewConfigMenu_Click()
Const ProcName As String = "NewConfigMenu_Click"
On Error GoTo Err

newAppConfig

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub RenameConfigMenu_Click()
Const ProcName As String = "RenameConfigMenu_Click"
On Error GoTo Err

ConfigsTV.StartLabelEdit

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub SetDefaultConfigMenu_Click()
Const ProcName As String = "SetDefaultConfigMenu_Click"
On Error GoTo Err

toggleDefaultConfig

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ChangesPending() As Boolean
Const ProcName As String = "ChangesPending"
On Error GoTo Err

If StudyLibConfigurer1.Dirty Or SPConfigurer1.Dirty Then
    ChangesPending = True
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Dirty() As Boolean
Const ProcName As String = "Dirty"
On Error GoTo Err

If Not mConfigStore Is Nothing Then Dirty = mConfigStore.Dirty

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get AppConfig( _
                ByVal name As String) As ConfigurationSection
Const ProcName As String = "AppConfig"
On Error GoTo Err

Set AppConfig = findConfig(name)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FirstAppConfig() As ConfigurationSection
Const ProcName As String = "FirstAppConfig"
On Error GoTo Err

Dim AppConfig As ConfigurationSection
For Each AppConfig In mAppConfigs
    Exit For
Next

Set FirstAppConfig = AppConfig

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Get SelectedAppConfig() As ConfigurationSection
Set SelectedAppConfig = mSelectedAppConfig
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

If Value Is Nothing Then Exit Property

Set mTheme = Value
UserControl.BackColor = mTheme.BackColor
gApplyTheme mTheme, UserControl.Controls

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ApplyPendingChanges()
Const ProcName As String = "ApplyPendingChanges"
On Error GoTo Err

If StudyLibConfigurer1.Dirty Then
    StudyLibConfigurer1.ApplyChanges
End If
If SPConfigurer1.Dirty Then
    SPConfigurer1.ApplyChanges
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub CreateNewAppConfig( _
                ByVal configName As String)
Const ProcName As String = "CreateNewAppConfig"
On Error GoTo Err

Set mCurrAppConfig = AddAppInstanceConfig(mConfigStore, _
                                    configName, _
                                    mFlags, _
                                    mPermittedServiceProviderRoles)

Set mCurrConfigNode = addConfigNode(mCurrAppConfig)
mCurrConfigNode.Expanded = True
ConfigsTV.SelectedItem = mCurrConfigNode
ConfigsTV_NodeClick ConfigsTV.SelectedItem

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

SPConfigurer1.Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Initialise( _
                ByVal pConfigFile As ConfigurationStore, _
                ByVal pApplicationName As String, _
                ByVal pExpectedConfigFileVersion As String, _
                ByVal pPermittedServiceProviderRoles As ServiceProviderRoles, _
                ByVal pFlags As ConfigFlags) As Boolean
Const ProcName As String = "Initialise"
On Error GoTo Err

LogMessage "Initialising ConfigManager", LogLevelDetail

Set mConfigStore = pConfigFile
mPermittedServiceProviderRoles = pPermittedServiceProviderRoles
mFlags = pFlags
    
If mConfigStore.ApplicationName <> pApplicationName Or _
    mConfigStore.fileVersion <> pExpectedConfigFileVersion Or _
    Not IsValidConfigurationFile(mConfigStore) _
Then
    LogMessage "The configuration file is not the correct format for this program"
    Exit Function
End If
    
LogMessage "Locating config definitions in config file", LogLevelDetail

Set mAppConfigs = GetAppInstanceConfigs(mConfigStore)

LogMessage "Loading config definitions into ConfigManager control", LogLevelDetail

Set mDefaultAppConfig = GetDefaultAppInstanceConfig(mConfigStore)

Dim AppConfig As ConfigurationSection
For Each AppConfig In mAppConfigs
    Dim newnode As Node
    Set newnode = addConfigNode(AppConfig)
    If AppConfig Is mDefaultAppConfig Then
        newnode.Bold = True
        Set mDefaultConfigNode = newnode
    End If
Next

If Not mDefaultConfigNode Is Nothing Then
    ConfigsTV.SelectedItem = mDefaultConfigNode
ElseIf ConfigsTV.Nodes.Count > 0 Then
    ConfigsTV.SelectedItem = ConfigsTV.Nodes(1)
End If
If Not ConfigsTV.SelectedItem Is Nothing Then ConfigsTV_NodeClick ConfigsTV.SelectedItem
Initialise = True
LogMessage "ConfigManager initialised ok", LogLevelDetail

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addConfigNode( _
                ByVal AppConfig As ConfigurationSection) As Node
Const ProcName As String = "addConfigNode"
On Error GoTo Err

Dim name As String
name = AppConfig.InstanceQualifier
Set addConfigNode = ConfigsTV.Nodes.Add(, , name, name)
Set addConfigNode.Tag = AppConfig
mConfigNames.Add name, name
ConfigsTV.Nodes.Add addConfigNode, tvwChild, , ConfigNodeServiceProviders
ConfigsTV.Nodes.Add addConfigNode, tvwChild, , ConfigNodeStudyLibraries

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub deleteAppConfig()
Const ProcName As String = "deleteAppConfig"
On Error GoTo Err

If MsgBox("Do you want to delete this configuration?" & vbCrLf & _
        "If you click Yes, all data for this configuration will be removed from the configuration file", _
        vbYesNo Or vbQuestion, _
        "Attention!") = vbYes Then
    RemoveAppInstanceConfig mConfigStore, mCurrAppConfig.InstanceQualifier
    ConfigsTV.Nodes.Remove ConfigsTV.SelectedItem.Index
    If mCurrAppConfig Is mDefaultAppConfig Then Set mDefaultAppConfig = Nothing
    Set mCurrAppConfig = Nothing
    If mCurrConfigNode Is mDefaultConfigNode Then Set mDefaultConfigNode = Nothing
    Set mCurrConfigNode = Nothing
    
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function findConfig( _
                ByVal name As String) As ConfigurationSection
Const ProcName As String = "findConfig"
On Error GoTo Err

Set findConfig = getAppInstanceConfig(mConfigStore, name)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub hideConfigControls()
Const ProcName As String = "hideConfigControls"
On Error GoTo Err

SPConfigurer1.Visible = False
StudyLibConfigurer1.Visible = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function nameAlreadyInUse( _
                ByVal name As String) As Boolean
Const ProcName As String = "nameAlreadyInUse"
On Error GoTo Err

On Error Resume Next
Dim s As String
s = mConfigNames(name)
If s <> "" Then nameAlreadyInUse = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub newAppConfig()
Const ProcName As String = "newAppConfig"
On Error GoTo Err

Dim name As String
name = NewConfigNameStub
Dim i As Long
Do While nameAlreadyInUse(name)
    i = i + 1
    name = NewConfigNameStub & i
Loop
CreateNewAppConfig name

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub renameCurrentConfig( _
                ByVal newName As String)
Const ProcName As String = "renameCurrentConfig"
On Error GoTo Err

mConfigNames.Remove mCurrAppConfig.InstanceQualifier
mCurrAppConfig.InstanceQualifier = newName
mConfigNames.Add newName, newName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCurrentConfig( _
                ByVal cs As ConfigurationSection, _
                ByVal lNode As Node)
Const ProcName As String = "setCurrentConfig"
On Error GoTo Err

Set mCurrAppConfig = cs
Set mCurrConfigNode = lNode

SPConfigurer1.Visible = False
StudyLibConfigurer1.Visible = False
DeleteConfigButton.Enabled = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showServiceProviderConfigDetails()
Const ProcName As String = "showServiceProviderConfigDetails"
On Error GoTo Err

hideConfigControls
SPConfigurer1.Left = Box1.Left
SPConfigurer1.Top = Box1.Top
SPConfigurer1.Visible = True
SPConfigurer1.Initialise gPermittedServiceProviderRoles, _
                        GetTradeBuildConfig(mCurrAppConfig), _
                        False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showStudyLibraryConfigDetails()
Const ProcName As String = "showStudyLibraryConfigDetails"
On Error GoTo Err

hideConfigControls
StudyLibConfigurer1.Left = Box1.Left
StudyLibConfigurer1.Top = Box1.Top
StudyLibConfigurer1.Visible = True
StudyLibConfigurer1.Initialise GetTradeBuildConfig(mCurrAppConfig)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub toggleDefaultConfig()
Const ProcName As String = "toggleDefaultConfig"
On Error GoTo Err

If mCurrAppConfig Is mDefaultAppConfig Then
    UnsetDefaultAppInstanceConfig mConfigStore
    mDefaultConfigNode.Bold = False
    Set mDefaultAppConfig = Nothing
    Set mDefaultConfigNode = Nothing
Else
    If Not mDefaultAppConfig Is Nothing Then mDefaultConfigNode.Bold = False
    SetDefaultAppInstanceConfig mConfigStore, mCurrAppConfig.InstanceQualifier
    
    Set mDefaultAppConfig = mCurrAppConfig
    Set mDefaultConfigNode = mCurrConfigNode
    mDefaultConfigNode.Bold = True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



