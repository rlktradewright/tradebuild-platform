VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyHorizontalRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ConfigSettingY As String = "&Y"
Private Const ConfigSettingColor As String = "&Color"
Private Const ConfigSettingThickness As String = "&Thickness"
Private Const ConfigSettingStyle As String = "&Style"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mY As Double
Private mColor As Long
Private mThickness As Long
Private mStyle As LineStyles

Private mConfig                                         As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mColor = vbBlack
mThickness = 1
mStyle = LineSolid
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Color(ByVal value As Long)
mColor = value
storeSettings
End Property

Public Property Get Color() As Long
Color = mColor
End Property

Friend Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
If value Is mConfig Then Exit Property
Set mConfig = value
storeSettings
End Property

Public Property Let Style(ByVal value As LineStyles)
mStyle = value
storeSettings
End Property

Public Property Get Style() As LineStyles
Style = mStyle
End Property

Public Property Let Thickness(ByVal value As Long)
mThickness = value
storeSettings
End Property

Public Property Get Thickness() As Long
Thickness = mThickness
End Property

Public Property Let Y(ByVal value As Double)
mY = value
storeSettings
End Property

Public Property Get Y() As Double
Y = mY
storeSettings
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Set mConfig = config
mY = mConfig.GetSetting(ConfigSettingY)
mColor = mConfig.GetSetting(ConfigSettingColor)
mThickness = mConfig.GetSetting(ConfigSettingThickness)
mStyle = mConfig.GetSetting(ConfigSettingStyle)
End Sub

Friend Sub RemoveFromConfig()
If Not mConfig Is Nothing Then mConfig.Remove
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub storeSettings()
If mConfig Is Nothing Then Exit Sub

mConfig.SetSetting ConfigSettingY, mY
mConfig.SetSetting ConfigSettingColor, mColor
mConfig.SetSetting ConfigSettingThickness, mThickness
mConfig.SetSetting ConfigSettingStyle, mStyle
End Sub

