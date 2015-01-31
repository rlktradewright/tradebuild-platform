VERSION 5.00
Begin VB.Form fFeaturesPanel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Features Panel"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleMode       =   0  'User
   ScaleWidth      =   4910.454
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TradeSkilDemo27.FeaturesPanel FeaturesPanel 
      Height          =   9675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   17066
   End
End
Attribute VB_Name = "fFeaturesPanel"
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

Event Hide()
Event Pin()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "fFeaturesPane"

'@================================================================================
' Member variables
'@================================================================================

Private mAppInstanceConfig                          As ConfigurationSection

'Private mMouseDown                                  As Boolean
'Private mLeftAtMousedown                            As Single
'Private mTopAtMouseDown                             As Single
'Private mMouseXAtMousedown                          As Single
'Private mMouseYAtMouseDown                          As Single

Private mTheme                                      As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

'Private Sub Form_Activate()
'Const ProcName As String = "Form_Activate"
'On Error GoTo Err
'
'Me.left = CLng(mAppInstanceConfig.GetSetting(ConfigSettingFloatingFeaturesPanelLeft, 0)) * Screen.TwipsPerPixelX
'Me.Top = CLng(mAppInstanceConfig.GetSetting(ConfigSettingFloatingFeaturesPanelTop, (Screen.Height - Me.Height) / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
'
'Exit Sub
'
'Err:
'gNotifyUnhandledError ProcName, ModuleName, ProjectName
'End Sub

Private Sub Form_Deactivate()
Const ProcName As String = "Form_Deactivate"
On Error GoTo Err

updateSettings

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Select Case UnloadMode
Case vbFormControlMenu
    Cancel = 1
    RaiseEvent Hide
Case vbFormCode

Case vbAppWindows

Case vbAppTaskManager

Case vbFormMDIForm

Case vbFormOwner

End Select
End Sub

Private Sub Form_Resize()
If Me.ScaleWidth < FeaturesPanel.MinimumWidth Then Me.Width = FeaturesPanel.MinimumWidth + Me.Width - Me.ScaleWidth
If Me.ScaleHeight < FeaturesPanel.MinimumHeight Then Me.Height = FeaturesPanel.MinimumHeight + Me.Height - Me.ScaleHeight
FeaturesPanel.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
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
' Controls Event Handlers
'@================================================================================

Private Sub FeaturesPanel_Hide()
Const ProcName As String = "FeaturesPanel_Hide"
On Error GoTo Err

RaiseEvent Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'Private Sub FeaturesPanel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Const ProcName As String = "FeaturesPanel_MouseDown"
'On Error GoTo Err
'
'mMouseDown = True
'mLeftAtMousedown = Me.Left
'mTopAtMouseDown = Me.Top
'mMouseXAtMousedown = x
'mMouseYAtMouseDown = y
'
'Exit Sub
'
'Err:
'gNotifyUnhandledError ProcName, ModuleName
'End Sub

'Private Sub FeaturesPanel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Const ProcName As String = "FeaturesPanel_MouseMove"
'On Error GoTo Err
'
''If mMouseDown Then moveMe x, y
'
'Exit Sub
'
'Err:
'gNotifyUnhandledError ProcName, ModuleName
'End Sub

'Private Sub FeaturesPanel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Const ProcName As String = "FeaturesPanel_MouseUp"
'On Error GoTo Err
'
'If mMouseDown Then moveMe x, y
'mMouseDown = False
'
'Exit Sub
'
'Err:
'gNotifyUnhandledError ProcName, ModuleName
'End Sub

Private Sub FeaturesPanel_Pin()
Const ProcName As String = "FeaturesPanel_Pin"
On Error GoTo Err

RaiseEvent Pin

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

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

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

FeaturesPanel.Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTradeBuildAPI As TradeBuildAPI, _
                ByVal pAppInstanceConfig As ConfigurationSection, _
                ByVal pTickerGrid As TickerGrid, _
                ByVal pInfoPanel As InfoPanel, _
                ByVal pInfoPanelFloating As InfoPanel, _
                ByVal pChartForms As ChartForms, _
                ByVal pOrderTicket As fOrderTicket)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mAppInstanceConfig = pAppInstanceConfig

Me.Move CLng(mAppInstanceConfig.GetSetting(ConfigSettingFloatingFeaturesPanelLeft, 0)) * Screen.TwipsPerPixelX, _
        CLng(mAppInstanceConfig.GetSetting(ConfigSettingFloatingFeaturesPanelTop, (Screen.Height - Me.Height) / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY, _
        CLng(mAppInstanceConfig.GetSetting(ConfigSettingFloatingFeaturesPanelWidth, 280)) * Screen.TwipsPerPixelX, _
        CLng(mAppInstanceConfig.GetSetting(ConfigSettingFloatingFeaturesPanelHeight, 650)) * Screen.TwipsPerPixelY

FeaturesPanel.Initialise False, pTradeBuildAPI, pAppInstanceConfig, pTickerGrid, pInfoPanel, pInfoPanelFloating, pChartForms, pOrderTicket

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

'Private Sub moveMe(ByVal x As Single, ByVal y As Single)
'Const ProcName As String = "moveMe"
'On Error GoTo Err
'
'Me.Move mLeftAtMousedown + x - mMouseXAtMousedown, mTopAtMouseDown + y - mMouseYAtMouseDown
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

Private Sub updateSettings()
Const ProcName As String = "updateSettings"
On Error GoTo Err

If Not mAppInstanceConfig Is Nothing Then
    mAppInstanceConfig.AddPrivateConfigurationSection ConfigSectionFloatingFeaturesPanel
    mAppInstanceConfig.SetSetting ConfigSettingFloatingFeaturesPanelLeft, Me.Left / Screen.TwipsPerPixelX
    mAppInstanceConfig.SetSetting ConfigSettingFloatingFeaturesPanelTop, Me.Top / Screen.TwipsPerPixelY
    mAppInstanceConfig.SetSetting ConfigSettingFloatingFeaturesPanelWidth, Me.Width / Screen.TwipsPerPixelX
    mAppInstanceConfig.SetSetting ConfigSettingFloatingFeaturesPanelHeight, Me.Height / Screen.TwipsPerPixelY
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




