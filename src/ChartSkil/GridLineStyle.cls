VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridLineStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Event PropertyChanged( _
                ByRef ev As PropertyChangedEvent)
                
'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GridLineStyle"

Private Const ConfigSettingColor                    As String = "&Color"
Private Const ConfigSettingThickness                As String = "&Thickness"
Private Const ConfigSettingLineStyle                As String = "&LineStyle"

'@================================================================================
' Member variables
'@================================================================================

Private mConfig As ConfigurationSection

Private mStyle                                      As LineStyle

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStyle = New LineStyle
mStyle.Color = &HC0C0C0
mStyle.LineStyle = LineSolid
mStyle.Thickness = 1
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

Public Property Let Color( _
                ByVal value As OLE_COLOR)
If mStyle.Color = value Then Exit Property
mStyle.Color = value
firePropertyChanged "Color"
End Property

Public Property Get Color() As OLE_COLOR
Color = mStyle.Color
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
Set mConfig = value

mConfig.SetSetting ConfigSettingColor, mStyle.Color
mConfig.SetSetting ConfigSettingThickness, mStyle.Thickness
mConfig.SetSetting ConfigSettingLineStyle, mStyle.LineStyle
End Property

Public Property Let LineStyle( _
                ByVal value As LineStyles)
If mStyle.LineStyle = value Then Exit Property
mStyle.LineStyle = value
firePropertyChanged "LineStyle"
End Property

Public Property Get LineStyle() As LineStyles
LineStyle = mStyle.LineStyle
End Property

Public Property Let Thickness( _
                ByVal value As Long)
If mStyle.Thickness = value Then Exit Property
mStyle.Thickness = value
firePropertyChanged "Thickness"
End Property

Public Property Get Thickness() As Long
Thickness = mStyle.Thickness
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As LineStyle
Set clone = New GridLineStyle
clone.Color = mStyle.Color
clone.LineStyle = mStyle.LineStyle
clone.Thickness = mStyle.Thickness
End Function

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)

Set mConfig = config
mStyle.Color = mConfig.GetSetting(ConfigSettingColor)
mStyle.Thickness = mConfig.GetSetting(ConfigSettingThickness)
mStyle.LineStyle = mConfig.GetSetting(ConfigSettingLineStyle)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub firePropertyChanged( _
                ByVal name As String)
Dim ev As PropertyChangedEvent

Set ev.Source = Me
ev.PropertyName = name
RaiseEvent PropertyChanged(ev)
End Sub

