VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#23.6#0"; "TWControls40.ocx"
Begin VB.Form fConfigEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuration editor"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TWControls40.TWButton CloseButton 
      Height          =   495
      Left            =   9120
      TabIndex        =   4
      Top             =   4560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Close"
      Object.Default         =   -1  'True
   End
   Begin TWControls40.TWButton ConfigureButton 
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Load Selected &Configuration"
      Object.Default         =   -1  'True
   End
   Begin TradeSkilDemo27.ConfigManager ConfigManager1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7223
   End
   Begin VB.TextBox CurrentConfigNameText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current configuration is:"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "fConfigEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "fConfigEditor"

'@================================================================================
' Member variables
'@================================================================================

Private mConfig                                     As ConfigurationSection

Private mSelectedAppConfig                          As ConfigurationSection

Private mOverridePositionSettings                   As Boolean

Private mTheme                                      As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Const ProcName As String = "Form_QueryUnload"
On Error GoTo Err

updateSettings

If UnloadMode = vbFormControlMenu Then
    Me.Hide
    Cancel = True

End If

If ConfigManager1.ChangesPending Then
    If MsgBox("Apply outstanding changes?" & vbCrLf & _
            "If you click No, your changes to this configuration item will be lost", _
            vbYesNo Or vbQuestion, _
            "Attention!") = vbYes Then
        ConfigManager1.ApplyPendingChanges
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const ProcName As String = "Form_Unload"
On Error GoTo Err

ConfigManager1.Finish

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

Private Sub CloseButton_Click()
Me.Hide
End Sub

Private Sub ConfigManager1_SelectedItemChanged()
Const ProcName As String = "ConfigManager1_SelectedItemChanged"
On Error GoTo Err

checkOkToLoadConfiguration

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ConfigureButton_Click()
Const ProcName As String = "ConfigureButton_Click"
On Error GoTo Err

updateSettings
Set mSelectedAppConfig = ConfigManager1.SelectedAppConfig
mOverridePositionSettings = False
Me.Hide

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

Friend Property Get SelectedAppConfig() As ConfigurationSection
Const ProcName As String = "SelectedAppConfig"
On Error GoTo Err

Set SelectedAppConfig = mSelectedAppConfig

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

If Value Is Nothing Then Exit Property

Set mTheme = Value
Me.BackColor = mTheme.BackColor
gApplyTheme mTheme, Me.Controls

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal pCurrAppInstanceConfig As ConfigurationSection, _
                ByVal pCentreWindow As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mConfig = pCurrAppInstanceConfig

If pCentreWindow Then
    mOverridePositionSettings = True
    Me.left = CLng((Screen.Width - Me.Width) / 2)
    Me.Top = CLng((Screen.Height - Me.Height) / 2)
Else
    Me.left = CLng(mConfig.GetSetting(ConfigSettingConfigEditorLeft, 0)) * Screen.TwipsPerPixelX
    Me.Top = CLng(mConfig.GetSetting(ConfigSettingConfigEditorTop, (Screen.Height - Me.Height) / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
End If

ConfigManager1.Initialise pConfigStore, App.ProductName, ConfigFileVersion, gPermittedServiceProviderRoles, ConfigFlags.ConfigFlagIncludeDefaultBarFormatterLibrary Or ConfigFlags.ConfigFlagIncludeDefaultStudyLibrary

CurrentConfigNameText = mConfig.InstanceQualifier

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkOkToLoadConfiguration()
Const ProcName As String = "checkOkToLoadConfiguration"
On Error GoTo Err

If Not ConfigManager1.SelectedAppConfig Is Nothing Then
    ConfigureButton.Enabled = True
    ConfigureButton.Default = True
Else
    ConfigureButton.Enabled = False
    CloseButton.Default = True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateSettings()
Const ProcName As String = "updateSettings"
On Error GoTo Err

If mOverridePositionSettings Then Exit Sub

If Not mConfig Is Nothing Then
    mConfig.AddPrivateConfigurationSection ConfigSectionConfigEditor
    mConfig.SetSetting ConfigSettingConfigEditorLeft, Me.left / Screen.TwipsPerPixelX
    mConfig.SetSetting ConfigSettingConfigEditorTop, Me.Top / Screen.TwipsPerPixelY
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


