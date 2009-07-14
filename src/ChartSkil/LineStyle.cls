VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LineStyle"
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
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ConfigSettingColor As String = "&Color"
Private Const ConfigSettingThickness As String = "&Thickness"
Private Const ConfigSettingLineStyle  As String = "&LineStyle"
Private Const ConfigSettingExtendBefore As String = "&ExtendBefore"
Private Const ConfigSettingExtendAfter As String = "&ExtendAfter"
Private Const ConfigSettingArrowStartStyle  As String = "&ArrowStartStyle"
Private Const ConfigSettingArrowStartLength As String = "&ArrowStartLength"
Private Const ConfigSettingArrowStartWidth As String = "&ArrowStartWidth"
Private Const ConfigSettingArrowStartColor As String = "&ArrowStartColor"
Private Const ConfigSettingArrowStartFillColor As String = "&ArrowStartFillColor"
Private Const ConfigSettingArrowStartfillStyle  As String = "&ArrowStartFillStyle"
Private Const ConfigSettingArrowEndStyle  As String = "&ArrowEndStyle"
Private Const ConfigSettingArrowEndLength As String = "&ArrowEndLength"
Private Const ConfigSettingArrowEndWidth As String = "&ArrowEndWidth"
Private Const ConfigSettingArrowEndColor As String = "&ArrowEndColor"
Private Const ConfigSettingArrowEndFillColor As String = "&ArrowEndFillColor"
Private Const ConfigSettingArrowEndFillStyle  As String = "&ArrowEndFillStyle"

'@================================================================================
' Member variables
'@================================================================================

Private mConfig As ConfigurationSection

Private mColor As Long
Private mThickness As Long
Private mLineStyle As LineStyles
Private mExtendBefore As Boolean
Private mExtendAfter As Boolean
Private mArrowStartStyle As ArrowStyles
Private mArrowStartLength As Long
Private mArrowStartWidth As Long
Private mArrowStartColor As Long
Private mArrowStartFillColor As Long
Private mArrowStartfillStyle As FillStyles
Private mArrowEndStyle As ArrowStyles
Private mArrowEndLength As Long
Private mArrowEndWidth As Long
Private mArrowEndColor As Long
Private mArrowEndFillColor As Long
Private mArrowEndFillStyle As FillStyles

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mColor = vbBlack
mThickness = 1
mLineStyle = LineStyles.LineSolid
mExtendBefore = False
mExtendAfter = False
mArrowStartStyle = ArrowStyles.ArrowNone
mArrowStartLength = 10
mArrowStartWidth = 10
mArrowStartColor = vbBlack
mArrowStartFillColor = vbBlack
mArrowStartfillStyle = FillStyles.FillSolid
mArrowEndStyle = ArrowStyles.ArrowNone
mArrowEndLength = 10
mArrowEndWidth = 10
mArrowEndColor = vbBlack
mArrowEndFillColor = vbBlack
mArrowEndFillStyle = FillStyles.FillSolid
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

Public Property Let ArrowStartColor(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "LineStyle" & ":" & "ArrowStartColor", _
            "Value is not a valid color"
mArrowStartColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowStartColor, mArrowStartColor
firePropertyChanged "ArrowStartColor"
End Property

Public Property Get ArrowStartColor() As OLE_COLOR
ArrowStartColor = mArrowStartColor
End Property

Public Property Let ArrowStartFillColor(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "LineStyle" & ":" & "ArrowStartFillColor", _
            "Value is not a valid color"
mArrowStartFillColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowStartFillColor, mArrowStartFillColor
firePropertyChanged "ArrowStartFillColor"
End Property

Public Property Get ArrowStartFillColor() As OLE_COLOR
ArrowStartFillColor = mArrowStartFillColor
End Property

Public Property Let ArrowStartFillStyle(ByVal value As FillStyles)
mArrowStartfillStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowStartfillStyle, mArrowStartfillStyle
firePropertyChanged "ArrowStartFillStyle"
End Property

Public Property Get ArrowStartFillStyle() As FillStyles
ArrowStartFillStyle = mArrowStartfillStyle
End Property

Public Property Let ArrowStartLength(ByVal value As Long)
mArrowStartLength = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowStartLength, mArrowStartLength
firePropertyChanged "ArrowStartLength"
End Property

Public Property Get ArrowStartLength() As Long
ArrowStartLength = mArrowStartLength
End Property

Public Property Let ArrowStartStyle(ByVal value As ArrowStyles)
mArrowStartStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowStartStyle, mArrowStartStyle
firePropertyChanged "ArrowStartStyle"
End Property

Public Property Get ArrowStartStyle() As ArrowStyles
ArrowStartStyle = mArrowStartStyle
End Property

Public Property Let ArrowStartWidth(ByVal value As Long)
mArrowStartWidth = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowStartWidth, mArrowStartWidth
firePropertyChanged "ArrowStartWidth"
End Property

Public Property Get ArrowStartWidth() As Long
ArrowStartWidth = mArrowStartWidth
End Property

Public Property Let ArrowEndColor(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "LineStyle" & ":" & "ArrowEndColor", _
            "Value is not a valid color"
mArrowEndColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowEndColor, mArrowEndColor
firePropertyChanged "ArrowEndColor"
End Property

Public Property Get ArrowEndColor() As OLE_COLOR
ArrowEndColor = mArrowEndColor
End Property

Public Property Let ArrowEndFillColor(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "LineStyle" & ":" & "ArrowEndFillColor", _
            "Value is not a valid color"
mArrowEndFillColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowEndFillColor, mArrowEndFillColor
firePropertyChanged "ArrowEndFillColor"
End Property

Public Property Get ArrowEndFillColor() As OLE_COLOR
ArrowEndFillColor = mArrowEndFillColor
End Property

Public Property Let ArrowEndFillStyle(ByVal value As FillStyles)
mArrowEndFillStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowEndFillColor, mArrowEndFillColor
firePropertyChanged "ArrowEndFillStyle"
End Property

Public Property Get ArrowEndFillStyle() As FillStyles
ArrowEndFillStyle = mArrowEndFillStyle
End Property

Public Property Let ArrowEndLength(ByVal value As Long)
mArrowEndLength = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowEndLength, mArrowEndLength
firePropertyChanged "ArrowEndLength"
End Property

Public Property Get ArrowEndLength() As Long
ArrowEndLength = mArrowEndLength
End Property

Public Property Let ArrowEndStyle(ByVal value As ArrowStyles)
mArrowEndStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowEndStyle, mArrowEndStyle
firePropertyChanged "ArrowEndStyle"
End Property

Public Property Get ArrowEndStyle() As ArrowStyles
ArrowEndStyle = mArrowEndStyle
End Property

Public Property Let ArrowEndWidth(ByVal value As Long)
mArrowEndWidth = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowEndWidth, mArrowEndWidth
firePropertyChanged "ArrowEndWidth"
End Property

Public Property Get ArrowEndWidth() As Long
ArrowEndWidth = mArrowEndWidth
End Property

Public Property Let Color(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "LineStyle" & ":" & "color", _
            "Value is not a valid color"
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

mConfig.SetSetting ConfigSettingColor, mColor
mConfig.SetSetting ConfigSettingThickness, mThickness
mConfig.SetSetting ConfigSettingLineStyle, mLineStyle
mConfig.SetSetting ConfigSettingExtendBefore, mExtendBefore
mConfig.SetSetting ConfigSettingExtendAfter, mExtendAfter
mConfig.SetSetting ConfigSettingArrowStartStyle, mArrowStartStyle
mConfig.SetSetting ConfigSettingArrowStartLength, mArrowStartLength
mConfig.SetSetting ConfigSettingArrowStartWidth, mArrowStartWidth
mConfig.SetSetting ConfigSettingArrowStartColor, mArrowStartColor
mConfig.SetSetting ConfigSettingArrowStartFillColor, mArrowStartFillColor
mConfig.SetSetting ConfigSettingArrowStartfillStyle, mArrowStartfillStyle
mConfig.SetSetting ConfigSettingArrowEndStyle, mArrowEndStyle
mConfig.SetSetting ConfigSettingArrowEndLength, mArrowEndLength
mConfig.SetSetting ConfigSettingArrowEndWidth, mArrowEndWidth
mConfig.SetSetting ConfigSettingArrowEndColor, mArrowEndColor
mConfig.SetSetting ConfigSettingArrowEndFillColor, mArrowEndFillColor
mConfig.SetSetting ConfigSettingArrowEndFillStyle, mArrowEndFillStyle

End Property

Public Property Let ExtendAfter(ByVal value As Boolean)
mExtendAfter = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingExtendAfter, mExtendAfter
firePropertyChanged "ExtendAfter"
End Property

Public Property Get ExtendAfter() As Boolean
ExtendAfter = mExtendAfter
End Property

Public Property Let ExtendBefore(ByVal value As Boolean)
mExtendBefore = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingExtendBefore, mExtendBefore
firePropertyChanged "ExtendBefore"
End Property

Public Property Get ExtendBefore() As Boolean
ExtendBefore = mExtendBefore
End Property

Public Property Let LineStyle(ByVal value As LineStyles)
mLineStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingLineStyle, mLineStyle
firePropertyChanged "LineStyle"
End Property

Public Property Get LineStyle() As LineStyles
LineStyle = mLineStyle
End Property

Public Property Let Thickness(ByVal value As Long)
mThickness = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingThickness, mThickness
firePropertyChanged "Thickness"
End Property

Public Property Get Thickness() As Long
Thickness = mThickness
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As LineStyle
Set clone = New LineStyle
clone.ArrowEndColor = mArrowEndColor
clone.ArrowEndFillColor = mArrowEndFillColor
clone.ArrowEndFillStyle = mArrowEndFillStyle
clone.ArrowEndLength = mArrowEndLength
clone.ArrowEndStyle = mArrowEndStyle
clone.ArrowEndWidth = mArrowEndWidth
clone.ArrowStartColor = mArrowStartColor
clone.ArrowStartFillColor = mArrowStartFillColor
clone.ArrowStartFillStyle = mArrowStartfillStyle
clone.ArrowStartLength = mArrowStartLength
clone.ArrowStartStyle = mArrowStartStyle
clone.ArrowStartWidth = mArrowStartWidth
clone.Color = mColor
clone.ExtendAfter = mExtendAfter
clone.ExtendBefore = mExtendBefore
clone.LineStyle = mLineStyle
clone.Thickness = mThickness
End Function

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)

Set mConfig = config
mColor = mConfig.GetSetting(ConfigSettingColor)
mThickness = mConfig.GetSetting(ConfigSettingThickness)
mLineStyle = mConfig.GetSetting(ConfigSettingLineStyle)
mExtendBefore = mConfig.GetSetting(ConfigSettingExtendBefore)
mExtendAfter = mConfig.GetSetting(ConfigSettingExtendAfter)
mArrowStartStyle = mConfig.GetSetting(ConfigSettingArrowStartStyle)
mArrowStartLength = mConfig.GetSetting(ConfigSettingArrowStartLength)
mArrowStartWidth = mConfig.GetSetting(ConfigSettingArrowStartWidth)
mArrowStartColor = mConfig.GetSetting(ConfigSettingArrowStartColor)
mArrowStartFillColor = mConfig.GetSetting(ConfigSettingArrowStartFillColor)
mArrowStartfillStyle = mConfig.GetSetting(ConfigSettingArrowStartfillStyle)
mArrowEndStyle = mConfig.GetSetting(ConfigSettingArrowEndStyle)
mArrowEndLength = mConfig.GetSetting(ConfigSettingArrowEndLength)
mArrowEndWidth = mConfig.GetSetting(ConfigSettingArrowEndWidth)
mArrowEndColor = mConfig.GetSetting(ConfigSettingArrowEndColor)
mArrowEndFillColor = mConfig.GetSetting(ConfigSettingArrowEndFillColor)
mArrowEndFillStyle = mConfig.GetSetting(ConfigSettingArrowEndFillStyle)
                
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


