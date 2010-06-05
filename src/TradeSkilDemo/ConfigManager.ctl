VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#240.0#0"; "TradeBuildUI2-6.ocx"
Object = "{6F9EA9CF-F55B-4AFA-8431-9ECC5BED8D43}#179.0#0"; "StudiesUI2-6.ocx"
Begin VB.UserControl ConfigManager 
   BackStyle       =   0  'Transparent
   ClientHeight    =   8325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16170
   ScaleHeight     =   8325
   ScaleWidth      =   16170
   Begin TradeBuildUI26.SPConfigurer SPConfigurer1 
      Height          =   4005
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   7064
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
   Begin StudiesUI26.StudyLibConfigurer StudyLibConfigurer1 
      Height          =   4005
      Left            =   8520
      TabIndex        =   5
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
Private Const ConfigNameTradeBuild          As String = "TradeBuild"

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

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

Set mConfigNames = New Collection

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

UserControl.Width = BoundingRect.Width
UserControl.Height = BoundingRect.Height

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
        SaveConfigButton.Enabled = False
        SaveConfigMenu.Enabled = False
    Case ConfigChangeTypes.ConfigDirty
        SaveConfigButton.Enabled = True
        SaveConfigMenu.Enabled = True
    End Select
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ConfigsTV_AfterLabelEdit( _
                cancel As Integer, _
                NewString As String)

Const ProcName As String = "ConfigsTV_AfterLabelEdit"

On Error GoTo Err

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

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub ConfigsTV_Collapse( _
                ByVal Node As MSComctlLib.Node)

End Sub

Private Sub ConfigsTV_MouseUp( _
                Button As Integer, _
                Shift As Integer, _
                x As Single, _
                y As Single)
                
Dim lNode As Node
Const ProcName As String = "ConfigsTV_MouseUp"

On Error GoTo Err

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

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
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
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub DeleteConfigButton_Click()
Const ProcName As String = "DeleteConfigButton_Click"

On Error GoTo Err

deleteAppConfig

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub DeleteConfigMenu_Click()
Const ProcName As String = "DeleteConfigMenu_Click"

On Error GoTo Err

deleteAppConfig

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub NewConfigButton_Click()
Const ProcName As String = "NewConfigButton_Click"

On Error GoTo Err

newAppConfig

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub NewConfigMenu_Click()
Const ProcName As String = "NewConfigMenu_Click"

On Error GoTo Err

newAppConfig

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub RenameConfigMenu_Click()
Const ProcName As String = "RenameConfigMenu_Click"

On Error GoTo Err

ConfigsTV.StartLabelEdit

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub SaveConfigButton_Click()
Const ProcName As String = "SaveConfigButton_Click"

On Error GoTo Err

saveConfigFile

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub SaveConfigMenu_Click()
Const ProcName As String = "SaveConfigMenu_Click"

On Error GoTo Err

saveConfigFile

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub SetDefaultConfigMenu_Click()
Const ProcName As String = "SetDefaultConfigMenu_Click"

On Error GoTo Err

toggleDefaultConfig

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get changesPending() As Boolean
Const ProcName As String = "changesPending"

On Error GoTo Err

If StudyLibConfigurer1.dirty Or SPConfigurer1.dirty Then
    changesPending = True
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get dirty() As Boolean
Const ProcName As String = "dirty"

On Error GoTo Err

If Not mConfigStore Is Nothing Then dirty = mConfigStore.dirty

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get appConfig( _
                ByVal name As String) As ConfigurationSection
Const ProcName As String = "appConfig"

On Error GoTo Err

Set appConfig = findConfig(name)

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

Public Property Get selectedAppConfig() As ConfigurationSection
Set selectedAppConfig = mSelectedAppConfig
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub applyPendingChanges()
Const ProcName As String = "applyPendingChanges"

On Error GoTo Err

If StudyLibConfigurer1.dirty Then
    StudyLibConfigurer1.ApplyChanges
End If
If SPConfigurer1.dirty Then
    SPConfigurer1.ApplyChanges
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Public Sub createNewAppConfig( _
                ByVal configName As String, _
                ByVal includeDefaultStudyLibrary As Boolean)
Const ProcName As String = "createNewAppConfig"

On Error GoTo Err

Set mCurrAppConfig = AddAppInstanceConfig(mConfigStore, _
                                    configName, _
                                    includeDefaultStudyLibrary)

Set mCurrConfigNode = addConfigNode(mCurrAppConfig)
mCurrConfigNode.Expanded = True
ConfigsTV.SelectedItem = mCurrConfigNode
ConfigsTV_NodeClick ConfigsTV.SelectedItem

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function initialise( _
                ByVal pConfigFile As ConfigurationStore, _
                ByVal pApplicationName As String, _
                ByVal pExpectedConfigFileVersion As String) As Boolean
Dim appConfig As ConfigurationSection
Dim newnode As Node

Const ProcName As String = "initialise"

On Error GoTo Err

LogMessage "Initialising ConfigManager", LogLevelDetail

Set mConfigStore = pConfigFile
    
If mConfigStore.ApplicationName <> pApplicationName Or _
    mConfigStore.fileVersion <> pExpectedConfigFileVersion Or _
    Not IsValidConfigurationFile(mConfigStore) _
Then
    LogMessage "The configuration file is not the correct format for this program"
    Exit Function
End If
    
mConfigStore.AddChangeListener Me

If mConfigStore.dirty Then
    SaveConfigButton.Enabled = True
    SaveConfigMenu.Enabled = True
End If

LogMessage "Locating config definitions in config file", LogLevelDetail

Set mAppConfigs = mConfigStore.GetConfigurationSection("/" & ConfigNameAppConfigs)

LogMessage "Loading config definitions into ConfigManager control", LogLevelDetail

Set mDefaultAppConfig = GetDefaultAppInstanceConfig(mConfigStore)

For Each appConfig In mAppConfigs
    Set newnode = addConfigNode(appConfig)
    If appConfig Is mDefaultAppConfig Then
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
initialise = True
LogMessage "ConfigManager initialised ok", LogLevelDetail

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub saveConfigFile( _
                Optional ByVal filename As String)
Const ProcName As String = "saveConfigFile"

On Error GoTo Err

mConfigStore.Save filename

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addConfigNode( _
                ByVal appConfig As ConfigurationSection) As Node
Dim name As String
Const ProcName As String = "addConfigNode"

On Error GoTo Err

name = appConfig.InstanceQualifier
Set addConfigNode = ConfigsTV.Nodes.Add(, , name, name)
Set addConfigNode.Tag = appConfig
mConfigNames.Add name, name
ConfigsTV.Nodes.Add addConfigNode, tvwChild, , ConfigNodeServiceProviders
ConfigsTV.Nodes.Add addConfigNode, tvwChild, , ConfigNodeStudyLibraries

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub deleteAppConfig()
Const ProcName As String = "deleteAppConfig"

On Error GoTo Err

If MsgBox("Do you want to delete this configuration?" & vbCrLf & _
        "If you click Yes, all data for this configuration will be removed from the configuration file", _
        vbYesNo Or vbQuestion, _
        "Attention!") = vbYes Then
    RemoveAppInstanceConfig mConfigStore, mCurrAppConfig.InstanceQualifier
    ConfigsTV.Nodes.Remove ConfigsTV.SelectedItem.index
    If mCurrAppConfig Is mDefaultAppConfig Then Set mDefaultAppConfig = Nothing
    Set mCurrAppConfig = Nothing
    If mCurrConfigNode Is mDefaultConfigNode Then Set mDefaultConfigNode = Nothing
    Set mCurrConfigNode = Nothing
    
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function findConfig( _
                ByVal name As String) As ConfigurationSection
Const ProcName As String = "findConfig"

On Error GoTo Err

Set findConfig = mAppConfigs.GetConfigurationSection(ConfigNameAppConfig & "(" & name & ")")

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub hideConfigControls()
Const ProcName As String = "hideConfigControls"

On Error GoTo Err

SPConfigurer1.Visible = False
StudyLibConfigurer1.Visible = False

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

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

Private Sub newAppConfig()
Dim name As String
Dim i As Long
Const ProcName As String = "newAppConfig"

On Error GoTo Err

name = NewConfigNameStub
Do While nameAlreadyInUse(name)
    i = i + 1
    name = NewConfigNameStub & i
Loop
createNewAppConfig name, False

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub showServiceProviderConfigDetails()
Const ProcName As String = "showServiceProviderConfigDetails"

On Error GoTo Err

hideConfigControls
SPConfigurer1.left = Box1.left
SPConfigurer1.Top = Box1.Top
SPConfigurer1.Visible = True
SPConfigurer1.initialise mCurrAppConfig.GetConfigurationSection(ConfigNameTradeBuild), _
                        False

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub showStudyLibraryConfigDetails()
Const ProcName As String = "showStudyLibraryConfigDetails"

On Error GoTo Err

hideConfigControls
StudyLibConfigurer1.left = Box1.left
StudyLibConfigurer1.Top = Box1.Top
StudyLibConfigurer1.Visible = True
StudyLibConfigurer1.initialise mCurrAppConfig.GetConfigurationSection(ConfigNameTradeBuild)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub



