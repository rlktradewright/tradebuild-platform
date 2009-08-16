VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
' @see
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
' Constants
'@================================================================================

Private Const ConfigSettingTailThickness        As String = "&TailThickness"
Private Const ConfigSettingOutlineThickness     As String = "&OutlineThickness"
Private Const ConfigSettingColor                As String = "&BarColor"
Private Const ConfigSettingUpColor              As String = "&UpColor"
Private Const ConfigSettingDownColor            As String = "&DownColor"
Private Const ConfigSettingDisplayMode          As String = "&DisplayMode"
Private Const ConfigSettingSolidUpBody          As String = "&SolidUpBody"
Private Const ConfigSettingThickness            As String = "&BarThickness"
Private Const ConfigSettingWidth                As String = "&BarWidth"
Private Const ConfigSettingIncludeInAutoscale   As String = "&IncludeInAutoscale"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mConfig As ConfigurationSection

Private mTailThickness As Long
Private mOutlineThickness As Long

Private mColor As Long
Private mUpColor As Long
Private mDownColor As Long

Private mDisplayMode As BarDisplayModes

Private mSolidUpBody As Boolean
Private mThickness As Long
Private mWidth As Single
Private mIncludeInAutoscale As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mTailThickness = 1
mOutlineThickness = 1
mUpColor = vbBlack
mDownColor = vbBlack
mDisplayMode = BarDisplayModeCandlestick
mSolidUpBody = False
mThickness = 2
mWidth = 0.6
mColor = -1
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
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "BarStyle" & ":" & "Color", _
            "Value is not a valid color or -1"

mColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingColor, mColor
firePropertyChanged "Color"
End Property

Public Property Get Color() As Long
Color = mColor
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
If value Is Nothing Then
    RemoveFromConfig
    Set mConfig = Nothing
    Exit Property
End If

Set mConfig = value

mConfig.SetSetting ConfigSettingTailThickness, mTailThickness
mConfig.SetSetting ConfigSettingOutlineThickness, mOutlineThickness
mConfig.SetSetting ConfigSettingColor, mColor
mConfig.SetSetting ConfigSettingUpColor, mUpColor
mConfig.SetSetting ConfigSettingDownColor, mDownColor
mConfig.SetSetting ConfigSettingDisplayMode, mDisplayMode
mConfig.SetSetting ConfigSettingSolidUpBody, mSolidUpBody
mConfig.SetSetting ConfigSettingThickness, mThickness
mConfig.SetSetting ConfigSettingWidth, mWidth
mConfig.SetSetting ConfigSettingIncludeInAutoscale, mIncludeInAutoscale

End Property

Public Property Let DisplayMode(ByVal value As BarDisplayModes)
mDisplayMode = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingDisplayMode, mDisplayMode
firePropertyChanged "DisplayMode"
End Property

Public Property Get DisplayMode() As BarDisplayModes
DisplayMode = mDisplayMode
End Property

Public Property Let DownColor(ByVal value As Long)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "BarStyle" & ":" & "downColor", _
            "Value is not a valid color or -1"
mDownColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingDownColor, mDownColor
firePropertyChanged "DownColor"
End Property

Public Property Get DownColor() As Long
DownColor = mDownColor
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
mIncludeInAutoscale = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingIncludeInAutoscale, mIncludeInAutoscale
firePropertyChanged "IncludeInAutoscale"
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let OutlineThickness(ByVal value As Long)
mOutlineThickness = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingOutlineThickness, mOutlineThickness
firePropertyChanged "OutlineThickness"
End Property

Public Property Get OutlineThickness() As Long
OutlineThickness = mOutlineThickness
End Property

Public Property Let SolidUpBody(ByVal value As Boolean)
mSolidUpBody = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingSolidUpBody, mSolidUpBody
firePropertyChanged "SolidUpBody"
End Property

Public Property Get SolidUpBody() As Boolean
SolidUpBody = mSolidUpBody
End Property

Public Property Let TailThickness(ByVal value As Long)
mTailThickness = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingTailThickness, mTailThickness
firePropertyChanged "TailThickness"
End Property

Public Property Get TailThickness() As Long
TailThickness = mTailThickness
End Property

Public Property Let Thickness(ByVal value As Long)
mThickness = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingThickness, mThickness
firePropertyChanged "Thickness"
End Property

Public Property Get Thickness() As Long
Thickness = mThickness
End Property

Public Property Let UpColor(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "BarStyle" & ":" & "upColor", _
            "Value is not a valid color"

mUpColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingUpColor, mUpColor
firePropertyChanged "UpColor"
End Property

Public Property Get UpColor() As Long
UpColor = mUpColor
End Property

Public Property Let Width(ByVal value As Single)
mWidth = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingWidth, mWidth
firePropertyChanged "Width"
End Property

Public Property Get Width() As Single
Width = mWidth
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As BarStyle
Set clone = New BarStyle
clone.Color = mColor
clone.Thickness = mThickness
clone.Width = mWidth
clone.DisplayMode = mDisplayMode
clone.DownColor = mDownColor
clone.IncludeInAutoscale = mIncludeInAutoscale
clone.OutlineThickness = mOutlineThickness
clone.SolidUpBody = mSolidUpBody
clone.TailThickness = mTailThickness
clone.UpColor = mUpColor
End Function

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Set mConfig = config
mTailThickness = mConfig.GetSetting(ConfigSettingTailThickness)
mOutlineThickness = mConfig.GetSetting(ConfigSettingOutlineThickness)
mColor = mConfig.GetSetting(ConfigSettingColor)
mUpColor = mConfig.GetSetting(ConfigSettingUpColor)
mDownColor = mConfig.GetSetting(ConfigSettingDownColor)
mDisplayMode = mConfig.GetSetting(ConfigSettingDisplayMode)
mSolidUpBody = mConfig.GetSetting(ConfigSettingSolidUpBody)
mThickness = mConfig.GetSetting(ConfigSettingThickness)
mWidth = mConfig.GetSetting(ConfigSettingWidth)
mIncludeInAutoscale = mConfig.GetSetting(ConfigSettingIncludeInAutoscale)
                
End Sub

Public Sub RemoveFromConfig()
If Not mConfig Is Nothing Then mConfig.Remove
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub firePropertyChanged( _
                ByVal Name As String)
Dim ev As PropertyChangedEvent

Set ev.Source = Me
ev.PropertyName = Name
RaiseEvent PropertyChanged(ev)
End Sub




