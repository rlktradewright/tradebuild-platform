VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartRegionStyle"
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


Private Const ModuleName                As String = "ChartRegionStyle"

Private Const ConfigSectionGridlineStyle                As String = "GridlineStyle"
Private Const ConfigSectionGridTextStyle                As String = "GridTextStyle"
Private Const ConfigSectionYAxisTextStyle               As String = "YAxisTextStyle"
Private Const ConfigSectionYCursorTextStyle             As String = "YCursorTextStyle"

Private Const ConfigSettingHasGrid                      As String = "&HasGrid"
Private Const ConfigSettingHasGridText                  As String = "&HasGridText"
Private Const ConfigSettingAutoscaling                  As String = "&Autoscale"
Private Const ConfigSettingIntegerYScale                As String = "&IntegerYScale"
Private Const ConfigSettingYScaleQuantum                As String = "&YScaleQuantum"
Private Const ConfigSettingGridlineSpacingY             As String = "&GridlineSpacingY"
Private Const ConfigSettingMinimumHeight                As String = "&MinimumHeight"
Private Const ConfigSettingCursorSnapsToTickBoundaries  As String = "&CursorSnapsToTickBoundaries"
Private Const ConfigSettingBackGradientFillColors       As String = "&BackGradientFillColors"

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

' indicates whether there is a grid, not whether it
' is currently visible. Visibility is controlled through
' the GridLineStyle's Thickness property.
Private mHasGrid As Boolean
Private mHasGridText As Boolean

Private mAutoscaling As Boolean
Private mIntegerYScale As Boolean
Private mYScaleQuantum As Double
Private mGridlineSpacingY As Double
Private mMinimumHeight As Double
Private mCursorSnapsToTickBoundaries As Boolean

Private mBackGradientFillColors() As Long

Private WithEvents mGridLineStyle As LineStyle
Attribute mGridLineStyle.VB_VarHelpID = -1
Private WithEvents mGridTextStyle As TextStyle
Attribute mGridTextStyle.VB_VarHelpID = -1
Private WithEvents mYAxisTextStyle As TextStyle
Attribute mYAxisTextStyle.VB_VarHelpID = -1
Private WithEvents mYCursorTextStyle As TextStyle
Attribute mYCursorTextStyle.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
setInitialGridLineStyle
setInitialGridTextStyle
setInitialYAxisTextStyle
setInitialYCursorTextStyle

mHasGrid = True
mHasGridText = False
mAutoscaling = True
mIntegerYScale = False
mYScaleQuantum = 0.01
mGridlineSpacingY = 1.8
mMinimumHeight = 0
mCursorSnapsToTickBoundaries = True
ReDim mBackGradientFillColors(0) As Long
mBackGradientFillColors(0) = vbWhite
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mGridLineStyle Event Handlers
'@================================================================================

Private Sub mGridLineStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
firePropertyChanged "GridLineStyle"
End Sub

'@================================================================================
' mGridTextStyle Event Handlers
'@================================================================================

Private Sub mGridTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
firePropertyChanged "GridTextStyle"
End Sub

'@================================================================================
' mYAxisTextStyle Event Handlers
'@================================================================================

Private Sub mYAxisTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
firePropertyChanged "YAxisTextStyle"
End Sub

'@================================================================================
' mYCursorTextStyle Event Handlers
'@================================================================================

Private Sub mYCursorTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
firePropertyChanged "YCursorTextStyle"
End Sub
'@================================================================================
' Properties
'@================================================================================

Public Property Get Autoscaling() As Boolean
Autoscaling = mAutoscaling
End Property

Public Property Let Autoscaling(ByVal value As Boolean)
mAutoscaling = value
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingAutoscaling, mAutoscaling
firePropertyChanged "Autoscaling"
End Property

Public Property Get BackColor() As Long
BackColor = mBackGradientFillColors(0)
End Property

Public Property Let BackColor(ByVal value As Long)
ReDim mBackGradientFillColors(0) As Long
mBackGradientFillColors(0) = value
If Not mConfig Is Nothing Then setBackGradientFillColorsConfig
firePropertyChanged "BackColor"
End Property

Public Property Get BackGradientFillColors() As Long()
BackGradientFillColors = mBackGradientFillColors
End Property

Public Property Let BackGradientFillColors(ByRef value() As Long)
Dim numColors As Long

On Error Resume Next
numColors = UBound(value) + 1
On Error GoTo 0

If numColors = 0 Then
    Erase mBackGradientFillColors
    Exit Property
End If

If numColors < 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "BackGradientFillColors", _
            "At least 1 background gradient fill color must be supplied"
End If

If numColors > 5 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "BackGradientFillColors", _
            "At most 5 background gradient fill colors may be supplied"
End If

mBackGradientFillColors = value

If Not mConfig Is Nothing Then setBackGradientFillColorsConfig

firePropertyChanged "BackGradientFillColors"
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)

If value Is Nothing Then
    RemoveFromConfig
    Set mConfig = Nothing
    Exit Property
End If

Set mConfig = value

mGridLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionGridlineStyle)
mGridTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionGridTextStyle)
mYAxisTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionYAxisTextStyle)
mYCursorTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionYCursorTextStyle)
mConfig.SetSetting ConfigSettingHasGrid, mHasGrid
mConfig.SetSetting ConfigSettingHasGridText, mHasGridText
mConfig.SetSetting ConfigSettingAutoscaling, mAutoscaling
mConfig.SetSetting ConfigSettingIntegerYScale, mIntegerYScale
mConfig.SetSetting ConfigSettingYScaleQuantum, mYScaleQuantum
mConfig.SetSetting ConfigSettingGridlineSpacingY, mGridlineSpacingY
mConfig.SetSetting ConfigSettingMinimumHeight, mMinimumHeight
mConfig.SetSetting ConfigSettingCursorSnapsToTickBoundaries, mCursorSnapsToTickBoundaries

setBackGradientFillColorsConfig
End Property

Public Property Get GridlineSpacingY() As Double
GridlineSpacingY = mGridlineSpacingY
End Property

Public Property Let GridlineSpacingY(ByVal value As Double)
mGridlineSpacingY = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingGridlineSpacingY, mGridlineSpacingY
firePropertyChanged "GridlineSpacingY"
End Property

Public Property Get GridLineStyle() As LineStyle
Set GridLineStyle = mGridLineStyle
End Property

Public Property Let GridLineStyle(ByVal value As LineStyle)
Set mGridLineStyle = value
If Not mConfig Is Nothing Then mGridLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionGridlineStyle)
firePropertyChanged "GridLineStyle"
End Property

Public Property Get GridTextStyle() As TextStyle
Set GridTextStyle = mGridTextStyle
End Property

Public Property Let GridTextStyle(ByVal value As TextStyle)
Set mGridTextStyle = value
If Not mConfig Is Nothing Then mGridTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionGridTextStyle)
firePropertyChanged "GridTextStyle"
End Property

Public Property Get HasGrid() As Boolean
HasGrid = mHasGrid
End Property

Public Property Let HasGrid(ByVal val As Boolean)
mHasGrid = val
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingHasGrid, mHasGrid
firePropertyChanged "HasGrid"
End Property

Public Property Get HasGridText() As Boolean
HasGridText = mHasGridText
End Property

Public Property Let HasGridText(ByVal val As Boolean)
mHasGridText = val
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingHasGridText, mHasGridText
firePropertyChanged "HasGridText"
End Property

Public Property Get IntegerYScale() As Boolean
IntegerYScale = mIntegerYScale
End Property

Public Property Let IntegerYScale(ByVal value As Boolean)
mIntegerYScale = value
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingIntegerYScale, mIntegerYScale
firePropertyChanged "IntegerYScale"
End Property

Public Property Get MinimumHeight() As Double
MinimumHeight = mMinimumHeight
End Property

Public Property Let MinimumHeight(ByVal value As Double)
mMinimumHeight = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingMinimumHeight, mMinimumHeight
firePropertyChanged "MinimumHeight"
End Property

Public Property Let CursorSnapsToTickBoundaries(ByVal value As Boolean)
mCursorSnapsToTickBoundaries = value
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingCursorSnapsToTickBoundaries, mCursorSnapsToTickBoundaries
firePropertyChanged "CursorSnapsToTickBoundaries"
End Property

Public Property Get CursorSnapsToTickBoundaries() As Boolean
CursorSnapsToTickBoundaries = mCursorSnapsToTickBoundaries
End Property

Public Property Get YAxisTextStyle() As TextStyle
Set YAxisTextStyle = mYAxisTextStyle
End Property

Public Property Let YAxisTextStyle(ByVal value As TextStyle)
Set mYAxisTextStyle = value
If Not mConfig Is Nothing Then mYAxisTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionYAxisTextStyle)
firePropertyChanged "YAxisTextStyle"
End Property

Public Property Get YCursorTextStyle() As TextStyle
Set YCursorTextStyle = mYCursorTextStyle
End Property

Public Property Let YCursorTextStyle(ByVal value As TextStyle)
Set mYCursorTextStyle = value
If Not mConfig Is Nothing Then mYCursorTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionYCursorTextStyle)
firePropertyChanged "YCursorTextStyle"
End Property

Public Property Get YScaleQuantum() As Double
YScaleQuantum = mYScaleQuantum
End Property

Public Property Let YScaleQuantum(ByVal value As Double)
mYScaleQuantum = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingYScaleQuantum, mYScaleQuantum
firePropertyChanged "YScaleQuantum"
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As ChartRegionStyle
Set clone = New ChartRegionStyle
clone.Autoscaling = mAutoscaling
clone.BackGradientFillColors = mBackGradientFillColors
clone.GridlineSpacingY = mGridlineSpacingY
clone.GridLineStyle = mGridLineStyle.clone
clone.GridTextStyle = mGridTextStyle
clone.IntegerYScale = mIntegerYScale
clone.MinimumHeight = mMinimumHeight
clone.HasGrid = mHasGrid
clone.HasGridText = mHasGridText
clone.CursorSnapsToTickBoundaries = mCursorSnapsToTickBoundaries
clone.YAxisTextStyle = mYAxisTextStyle
clone.YCursorTextStyle = mYCursorTextStyle
clone.YScaleQuantum = mYScaleQuantum
End Function

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Dim ar() As String
Dim i As Long

Set mConfig = config

getGridLineStyleFromConfig
getGridTextStyleFromConfig
getYAxisTextStyleFromConfig
getYCursorTextStyleFromConfig

mHasGrid = CBool(mConfig.GetSetting(ConfigSettingHasGrid, "True"))
mHasGridText = CBool(mConfig.GetSetting(ConfigSettingHasGridText, "False"))
mAutoscaling = CBool(mConfig.GetSetting(ConfigSettingAutoscaling, "True"))
mIntegerYScale = CBool(mConfig.GetSetting(ConfigSettingIntegerYScale, "False"))
mYScaleQuantum = CDbl(mConfig.GetSetting(ConfigSettingYScaleQuantum, "0.01"))
mGridlineSpacingY = CDbl(mConfig.GetSetting(ConfigSettingGridlineSpacingY, "1.8"))
mMinimumHeight = CDbl(mConfig.GetSetting(ConfigSettingMinimumHeight, "0"))
mCursorSnapsToTickBoundaries = CBool(mConfig.GetSetting(ConfigSettingCursorSnapsToTickBoundaries, "True"))

ar = Split(mConfig.GetSetting(ConfigSettingBackGradientFillColors), ",")
ReDim mBackGradientFillColors(UBound(ar())) As Long
For i = 0 To UBound(ar())
    mBackGradientFillColors(i) = CLng(ar(i))
Next
                
End Sub

Public Sub RemoveFromConfig()
If Not mConfig Is Nothing Then mConfig.Remove
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub getGridLineStyleFromConfig()
Dim cs As ConfigurationSection
Set cs = mConfig.GetConfigurationSection(ConfigSectionGridlineStyle)
If Not cs Is Nothing Then
    Set mGridLineStyle = New LineStyle
    mGridLineStyle.LoadFromConfig cs
End If
End Sub

Private Sub getGridTextStyleFromConfig()
Dim cs As ConfigurationSection
Set cs = mConfig.GetConfigurationSection(ConfigSectionGridTextStyle)
If Not cs Is Nothing Then
    Set mGridTextStyle = New TextStyle
    mGridTextStyle.LoadFromConfig cs
End If
End Sub

Private Sub getYAxisTextStyleFromConfig()
Dim cs As ConfigurationSection
Set cs = mConfig.GetConfigurationSection(ConfigSectionYAxisTextStyle)
If Not cs Is Nothing Then
    Set mYAxisTextStyle = New TextStyle
    mYAxisTextStyle.LoadFromConfig cs
End If
End Sub

Private Sub getYCursorTextStyleFromConfig()
Dim cs As ConfigurationSection
Set cs = mConfig.GetConfigurationSection(ConfigSectionYCursorTextStyle)
If Not cs Is Nothing Then
    Set mYCursorTextStyle = New TextStyle
    mYCursorTextStyle.LoadFromConfig cs
End If
End Sub

Private Sub firePropertyChanged( _
                ByVal Name As String)
Dim ev As PropertyChangedEvent

Set ev.Source = Me
ev.PropertyName = Name
RaiseEvent PropertyChanged(ev)
End Sub

Private Sub setBackGradientFillColorsConfig()
Dim i As Long
Dim s As String
For i = 0 To UBound(mBackGradientFillColors)
    If i <> 0 Then s = s & ","
    s = s & CStr(mBackGradientFillColors(i))
Next
mConfig.SetSetting ConfigSettingBackGradientFillColors, s
End Sub

Private Sub setInitialGridLineStyle()
Set mGridLineStyle = New LineStyle
mGridLineStyle.Color = &HC0C0C0
End Sub

Private Sub setInitialGridTextStyle()
Dim aFont As StdFont

Set mGridTextStyle = New TextStyle
mGridTextStyle.Box = False
mGridTextStyle.Color = vbBlack
    
Set aFont = New StdFont
aFont.Italic = False
aFont.Name = "Arial"
aFont.size = 8
aFont.Underline = False
aFont.Bold = False
mGridTextStyle.Font = aFont

End Sub

Private Sub setInitialYAxisTextStyle()
Dim aFont As StdFont

Set mYAxisTextStyle = New TextStyle
mYAxisTextStyle.Color = vbBlack
    
Set aFont = New StdFont
aFont.Italic = False
aFont.Name = "Arial"
aFont.size = 8
aFont.Underline = False
aFont.Bold = False
mYAxisTextStyle.Font = aFont

End Sub

Private Sub setInitialYCursorTextStyle()
Dim aFont As StdFont

Set mYCursorTextStyle = New TextStyle
mYCursorTextStyle.Box = True
mYCursorTextStyle.BoxColor = vbBlack
mYCursorTextStyle.BoxFillColor = vbWhite
mYCursorTextStyle.BoxFillStyle = FillSolid
mYCursorTextStyle.BoxThickness = 1
mYCursorTextStyle.Color = vbBlack
mYCursorTextStyle.PaddingX = 1
mYCursorTextStyle.PaddingY = 0
    
Set aFont = New StdFont
aFont.Italic = False
aFont.Name = "Arial"
aFont.size = 8
aFont.Underline = False
aFont.Bold = False
mYCursorTextStyle.Font = aFont

End Sub


