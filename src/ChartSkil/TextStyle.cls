VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextStyle"
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

Private Const ConfigSectionFont                     As String = "Font"

Private Const ConfigSettingColor                    As String = "&Color"
Private Const ConfigSettingBox                      As String = "&Box"
Private Const ConfigSettingBoxColor                 As String = "&BoxColor"
Private Const ConfigSettingBoxStyle                 As String = "&BoxStyle"
Private Const ConfigSettingBoxThickness             As String = "&BoxThickness"
Private Const ConfigSettingBoxFillColor             As String = "&BoxFillColor"
Private Const ConfigSettingBoxFillStyle             As String = "&BoxFillStyle"
Private Const ConfigSettingBoxFillWithBackgroundColor   As String = "&BoxFillWithBackgroundColor"
Private Const ConfigSettingAlign                    As String = "&Align"
Private Const ConfigSettingPaddingX                 As String = "&PaddingX"
Private Const ConfigSettingPaddingY                 As String = "&PaddingY"
Private Const ConfigSettingIncludeInAutoscale   As String = "&IncludeInAutoscale"
Private Const ConfigSettingFixedX               As String = "&FixedX"
Private Const ConfigSettingFixedY               As String = "&FixedY"
Private Const ConfigSettingExtended             As String = "&Extended"

'@================================================================================
' Member variables
'@================================================================================

Private mConfig As ConfigurationSection

Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1
Private mColor As Long
Private mBox As Boolean
Private mBoxColor As Long
Private mBoxStyle As LineStyles
Private mBoxThickness As Long
Private mBoxFillColor As Long
Private mBoxFillStyle As FillStyles
Private mBoxFillWithBackgroundColor As Boolean
Private mAlign As TextAlignModes
Private mPaddingX As Double
Private mPaddingY As Double
Private mIncludeInAutoscale As Boolean
Private mFixedX As Boolean
Private mFixedY As Boolean
Private mExtended As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

Set mFont = New StdFont
mFont.Bold = False
mFont.Italic = False
mFont.Name = "Arial"
mFont.size = 8
mFont.Strikethrough = False
mFont.Underline = False

mColor = vbBlack
mBox = False
mBoxColor = vbBlack
mBoxStyle = LineStyles.LineSolid
mBoxThickness = 1
mBoxFillColor = vbWhite
mBoxFillStyle = FillStyles.FillSolid
mBoxFillWithBackgroundColor = False
mAlign = TextAlignModes.AlignBottomRight
mPaddingX = 1
'mPaddingY = 0.5
mIncludeInAutoscale = False
mFixedX = False
mFixedY = False
mExtended = False

End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFont Event Handlers
'@================================================================================

Private Sub mFont_FontChanged(ByVal PropertyName As String)
firePropertyChanged "Font"
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let Align(ByVal value As TextAlignModes)
If mAlign = value Then Exit Property
mAlign = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingAlign, mAlign
firePropertyChanged "Align"
End Property

Public Property Get Align() As TextAlignModes
Align = mAlign
End Property

Public Property Let Box(ByVal value As Boolean)
If mBox = value Then Exit Property
mBox = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBox, mBox
firePropertyChanged "Box"
End Property

Public Property Get Box() As Boolean
Box = mBox
End Property

Public Property Let BoxColor(ByVal value As Long)
If mBoxColor = value Then Exit Property
mBoxColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBoxColor, mBoxColor
firePropertyChanged "BoxColor"
End Property

Public Property Get BoxColor() As Long
BoxColor = mBoxColor
End Property

Public Property Let BoxFillColor(ByVal value As Long)
If mBoxFillColor = value Then Exit Property
mBoxFillColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBoxFillColor, mBoxFillColor
firePropertyChanged "BoxFillColor"
End Property

Public Property Get BoxFillColor() As Long
BoxFillColor = mBoxFillColor
End Property

Public Property Let BoxFillStyle(ByVal value As FillStyles)
If mBoxFillStyle = value Then Exit Property
mBoxFillStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBoxFillStyle, mBoxFillStyle
firePropertyChanged "BoxFillStyle"
End Property

Public Property Get BoxFillStyle() As FillStyles
BoxFillStyle = mBoxFillStyle
End Property

Public Property Let BoxFillWithBackgroundColor(ByVal value As Boolean)
If mBoxFillWithBackgroundColor = value Then Exit Property
mBoxFillWithBackgroundColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBoxFillWithBackgroundColor, mBoxFillWithBackgroundColor
firePropertyChanged "BoxFillWithBackgroundColor"
End Property

Public Property Get BoxFillWithBackgroundColor() As Boolean
BoxFillWithBackgroundColor = mBoxFillWithBackgroundColor
End Property

Public Property Let BoxStyle(ByVal value As LineStyles)
If mBoxStyle = value Then Exit Property
mBoxStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBoxStyle, mBoxStyle
firePropertyChanged "BoxStyle"
End Property

Public Property Get BoxStyle() As LineStyles
BoxStyle = mBoxStyle
End Property

Public Property Let BoxThickness(ByVal value As Long)
If mBoxThickness = value Then Exit Property
mBoxThickness = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBoxThickness, mBoxThickness
firePropertyChanged "BoxThickness"
End Property

Public Property Get BoxThickness() As Long
BoxThickness = mBoxThickness
End Property

Public Property Let Color(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "TextStyle" & ":" & "color", _
            "Value is not a valid color"
If mColor = value Then Exit Property
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

SaveFontToConfig mFont, mConfig.AddConfigurationSection(ConfigSectionFont)
mConfig.SetSetting ConfigSettingColor, mColor
mConfig.SetSetting ConfigSettingBox, mBox
mConfig.SetSetting ConfigSettingBoxColor, mBoxColor
mConfig.SetSetting ConfigSettingBoxStyle, mBoxStyle
mConfig.SetSetting ConfigSettingBoxThickness, mBoxThickness
mConfig.SetSetting ConfigSettingBoxFillColor, mBoxFillColor
mConfig.SetSetting ConfigSettingBoxFillStyle, mBoxFillStyle
mConfig.SetSetting ConfigSettingBoxFillWithBackgroundColor, mBoxFillWithBackgroundColor
mConfig.SetSetting ConfigSettingAlign, mAlign
mConfig.SetSetting ConfigSettingPaddingX, mPaddingX
mConfig.SetSetting ConfigSettingPaddingY, mPaddingY
mConfig.SetSetting ConfigSettingIncludeInAutoscale, mIncludeInAutoscale
mConfig.SetSetting ConfigSettingExtended, mExtended
mConfig.SetSetting ConfigSettingFixedX, mFixedX
mConfig.SetSetting ConfigSettingFixedY, mFixedY

End Property

Public Property Let Extended(ByVal value As Boolean)
If mExtended = value Then Exit Property
mExtended = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingExtended, mExtended
firePropertyChanged "Extended"
End Property

Public Property Get Extended() As Boolean
Extended = mExtended
End Property

Public Property Let FixedX(ByVal value As Boolean)
If mFixedX = value Then Exit Property
mFixedX = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingFixedX, mFixedX
firePropertyChanged "FixedX"
End Property

Public Property Get FixedX() As Boolean
FixedX = mFixedX
End Property

Public Property Let FixedY(ByVal value As Boolean)
If mFixedY = value Then Exit Property
mFixedY = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingFixedY, mFixedY
firePropertyChanged "FixedY"
End Property

Public Property Get FixedY() As Boolean
FixedY = mFixedY
End Property

Public Property Let Font(ByVal value As StdFont)
If mFont Is value Then Exit Property
Set mFont = value
If Not mConfig Is Nothing Then SaveFontToConfig mFont, mConfig.AddConfigurationSection(ConfigSectionFont)
firePropertyChanged "Font"
End Property

Public Property Get Font() As StdFont
Set Font = mFont
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value Then Exit Property
mIncludeInAutoscale = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingIncludeInAutoscale, mIncludeInAutoscale
firePropertyChanged "IncludeInAutoscale"
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let PaddingX(ByVal value As Double)
If mPaddingX = value Then Exit Property
mPaddingX = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingPaddingX, mPaddingX
firePropertyChanged "PaddingX"
End Property

Public Property Get PaddingX() As Double
PaddingX = mPaddingX
End Property

Public Property Let PaddingY(ByVal value As Double)
If mPaddingY = value Then Exit Property
mPaddingY = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingPaddingY, mPaddingY
firePropertyChanged "PaddingY"
End Property

Public Property Get PaddingY() As Double
PaddingY = mPaddingY
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As TextStyle
Set clone = New TextStyle
clone.Align = mAlign
clone.Box = mBox
clone.BoxColor = mBoxColor
clone.BoxFillColor = mBoxFillColor
clone.BoxFillStyle = mBoxFillStyle
clone.BoxStyle = mBoxStyle
clone.BoxThickness = mBoxThickness
clone.Color = mColor
clone.Font = gCloneFont(mFont)
clone.PaddingX = mPaddingX
clone.PaddingY = mPaddingY
clone.IncludeInAutoscale = mIncludeInAutoscale
clone.Extended = mExtended
clone.FixedX = mFixedX
clone.FixedY = mFixedY

End Function

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)

Set mConfig = config

Set mFont = LoadFontFromConfig(mConfig.AddConfigurationSection(ConfigSectionFont))

mColor = mConfig.GetSetting(ConfigSettingColor)
mBox = mConfig.GetSetting(ConfigSettingBox)
mBoxColor = mConfig.GetSetting(ConfigSettingBoxColor)
mBoxStyle = mConfig.GetSetting(ConfigSettingBoxStyle)
mBoxThickness = mConfig.GetSetting(ConfigSettingBoxThickness)
mBoxFillColor = mConfig.GetSetting(ConfigSettingBoxFillColor)
mBoxFillStyle = mConfig.GetSetting(ConfigSettingBoxFillStyle)
mBoxFillWithBackgroundColor = mConfig.GetSetting(ConfigSettingBoxFillWithBackgroundColor)
mAlign = mConfig.GetSetting(ConfigSettingAlign)
mPaddingX = mConfig.GetSetting(ConfigSettingPaddingX)
mPaddingY = mConfig.GetSetting(ConfigSettingPaddingY)
mIncludeInAutoscale = mConfig.GetSetting(ConfigSettingIncludeInAutoscale)
mExtended = mConfig.GetSetting(ConfigSettingExtended)
mFixedX = mConfig.GetSetting(ConfigSettingFixedX)
mFixedY = mConfig.GetSetting(ConfigSettingFixedY)
                
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

