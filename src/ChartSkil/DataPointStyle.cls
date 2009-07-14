VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataPointStyle"
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

Private Const ConfigSettingLineThickness        As String = "&LineThickness"
Private Const ConfigSettingColor                As String = "&Color"
Private Const ConfigSettingUpColor              As String = "&UpColor"
Private Const ConfigSettingDownColor            As String = "&DownColor"
Private Const ConfigSettingLineStyle            As String = "&LineStyle"
Private Const ConfigSettingPointStyle           As String = "&PointStyle"
Private Const ConfigSettingDisplayMode          As String = "&DisplayMode"
Private Const ConfigSettingHistBarWidth         As String = "&HistBarWidth"
'Private Const ConfigSettingIncludeInAutoscale   As String = "&IncludeInAutoscale"

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

Private mLineThickness As Long
Private mColor As Long
Private mUpColor As Long
Private mDownColor As Long
Private mLineStyle As LineStyles
Private mPointStyle As PointStyles
Private mDisplayMode As DataPointDisplayModes
Private mHistBarWidth As Single

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mLineThickness = 1
mColor = vbBlack
mLineStyle = LineStyles.LineSolid
mPointStyle = PointRound
mDisplayMode = DataPointDisplayModes.DataPointDisplayModeLine
mHistBarWidth = 0.6
mDownColor = -1
mUpColor = -1
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

Public Property Let Color(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointStyle" & ":" & "color", _
            "Value is not a valid color"
If mColor = value Then Exit Property
mColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingColor, mColor
firePropertyChanged "Color"
End Property

Public Property Get Color() As OLE_COLOR
Color = mColor
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
If mConfig Is Nothing Then
    RemoveFromConfig
    Set mConfig = Nothing
    Exit Property
End If

Set mConfig = value

mConfig.SetSetting ConfigSettingLineThickness, mLineThickness
mConfig.SetSetting ConfigSettingColor, mColor
mConfig.SetSetting ConfigSettingUpColor, mUpColor
mConfig.SetSetting ConfigSettingDownColor, mDownColor
mConfig.SetSetting ConfigSettingLineStyle, mLineStyle
mConfig.SetSetting ConfigSettingPointStyle, mPointStyle
mConfig.SetSetting ConfigSettingDisplayMode, mDisplayMode
mConfig.SetSetting ConfigSettingHistBarWidth, mHistBarWidth
'mConfig.SetSetting ConfigSettingIncludeInAutoscale, mIncludeInAutoscale

End Property

Public Property Let DisplayMode(ByVal value As DataPointDisplayModes)
If mDisplayMode = value Then Exit Property
mDisplayMode = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingDisplayMode, mDisplayMode
firePropertyChanged "DisplayMode"
End Property

Public Property Get DisplayMode() As DataPointDisplayModes
DisplayMode = mDisplayMode
End Property

Public Property Let DownColor(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointStyle" & ":" & "downColor", _
            "Value is not a valid color or -1"
If mDownColor = value Then Exit Property
mDownColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingDownColor, mDownColor
firePropertyChanged "DownColor"
End Property

Public Property Get DownColor() As OLE_COLOR
DownColor = mDownColor
End Property

Public Property Let HistBarWidth(ByVal value As Single)
If value <= 0 Or value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointStyle" & ":" & "histBarWidth", _
            "histBarWidth must be greater than zero but not greater than 1"

If mHistBarWidth = value Then Exit Property
mHistBarWidth = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingHistBarWidth, mHistBarWidth
firePropertyChanged "HistBarWidth"
End Property

Public Property Get HistBarWidth() As Single
HistBarWidth = mHistBarWidth
End Property

'Public Property Let IncludeInAutoscale(ByVal value As Boolean)
'If mIncludeInAutoscale = value Then Exit Property
'mIncludeInAutoscale = value
'if not mconfig is nothing then mConfig.SetSetting ConfigSettingIncludeInAutoscale, mIncludeInAutoscale
'firePropertyChanged "IncludeInAutoscale"
'End Property
'
'Public Property Get IncludeInAutoscale() As Boolean
'IncludeInAutoscale = mIncludeInAutoscale
'End Property

Public Property Let LineStyle(ByVal value As LineStyles)
If mLineStyle = value Then Exit Property
mLineStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingLineStyle, mLineStyle
firePropertyChanged "LineStyle"
End Property

Public Property Get LineStyle() As LineStyles
LineStyle = mLineStyle
End Property

Public Property Let LineThickness(ByVal value As Long)
If mLineThickness = value Then Exit Property
mLineThickness = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingLineThickness, mLineThickness
firePropertyChanged "LineThickness"
End Property

Public Property Get LineThickness() As Long
LineThickness = mLineThickness
End Property

Public Property Let PointStyle(ByVal value As PointStyles)
If mPointStyle = value Then Exit Property
mPointStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingPointStyle, mPointStyle
firePropertyChanged "PointStyle"
End Property

Public Property Get PointStyle() As PointStyles
PointStyle = mPointStyle
End Property

Public Property Let UpColor(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointStyle" & ":" & "downColor", _
            "Value is not a valid color or -1"
If mUpColor = value Then Exit Property
mUpColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingUpColor, mUpColor
firePropertyChanged "UpColor"
End Property

Public Property Get UpColor() As OLE_COLOR
UpColor = mUpColor
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As DataPointStyle
Set clone = New DataPointStyle
clone.Color = mColor
clone.DisplayMode = mDisplayMode
clone.DownColor = mDownColor
clone.HistBarWidth = mHistBarWidth
'clone.IncludeInAutoscale = mIncludeInAutoscale
clone.LineStyle = mLineStyle
clone.LineThickness = mLineThickness
clone.PointStyle = PointStyle
clone.UpColor = mUpColor
End Function

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)

Set mConfig = config
mLineThickness = mConfig.GetSetting(ConfigSettingLineThickness)
mColor = mConfig.GetSetting(ConfigSettingColor)
mUpColor = mConfig.GetSetting(ConfigSettingUpColor)
mDownColor = mConfig.GetSetting(ConfigSettingDownColor)
mLineStyle = mConfig.GetSetting(ConfigSettingLineStyle)
mPointStyle = mConfig.GetSetting(ConfigSettingPointStyle)
mDisplayMode = mConfig.GetSetting(ConfigSettingDisplayMode)
mHistBarWidth = mConfig.GetSetting(ConfigSettingHistBarWidth)
'mIncludeInAutoscale = mConfig.GetSetting(ConfigSettingIncludeInAutoscale)
                
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





