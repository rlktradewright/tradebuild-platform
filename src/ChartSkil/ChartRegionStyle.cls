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
                ByRef ev As PropertyChangedEventData)
                
'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                As String = "ChartRegionStyle"

Private Const ConfigSectionGridlineStyle                As String = "GridlineStyle"
Private Const ConfigSectionGridTextStyle                As String = "GridTextStyle"
Private Const ConfigSectionSessionEndGridlineStyle      As String = "SessionEndGridlineStyle"
Private Const ConfigSectionSessionStartGridlineStyle    As String = "SessionStartGridlineStyle"
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

Private mName As String
Private mStyleType As String
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
Private WithEvents mSessionEndGridlineStyle As LineStyle
Attribute mSessionEndGridlineStyle.VB_VarHelpID = -1
Private WithEvents mSessionStartGridlineStyle As LineStyle
Attribute mSessionStartGridlineStyle.VB_VarHelpID = -1
Private WithEvents mYAxisTextStyle As TextStyle
Attribute mYAxisTextStyle.VB_VarHelpID = -1
Private WithEvents mYCursorTextStyle As TextStyle
Attribute mYCursorTextStyle.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mName = GenerateGUIDString
mStyleType = "Chart Region Styles"

setInitialGridLineStyle
setInitialGridTextStyle
setInitialSessionEndGridLineStyle
setInitialSessionStartGridLineStyle
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
' IGraphicObjectStyle Interface Members
'@================================================================================

Private Property Let IGraphicObjectStyle_ConfigurationSection(ByVal RHS As ConfigurationSection)
Me.ConfigurationSection = RHS
End Property

Private Sub IGraphicObjectStyle_LoadFromConfig(ByVal pConfig As ConfigurationSection)
Me.LoadFromConfig pConfig
End Sub

Private Property Get IGraphicObjectStyle_Name() As String
IGraphicObjectStyle_Name = mName
End Property

Private Property Get IGraphicObjectStyle_ProgId() As String
IGraphicObjectStyle_ProgId = ProjectName & "." & ModuleName
End Property

Private Property Get IGraphicObjectStyle_StyleType() As String
IGraphicObjectStyle_StyleType = mStyleType
End Property

'@================================================================================
' mGridLineStyle Event Handlers
'@================================================================================

Private Sub mGridLineStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mGridLineStyle_PropertyChanged"
On Error GoTo Err

firePropertyChanged "GridLineStyle"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mGridTextStyle Event Handlers
'@================================================================================

Private Sub mGridTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mGridTextStyle_PropertyChanged"
On Error GoTo Err

firePropertyChanged "GridTextStyle"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mSessionEndGridLineStyle Event Handlers
'@================================================================================

Private Sub mSessionEndGridLineStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mSessionEndGridLineStyle_PropertyChanged"
On Error GoTo Err

firePropertyChanged "SessionEndGridLineStyle"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mSessionStartGridLineStyle Event Handlers
'@================================================================================

Private Sub mSessionStartGridLineStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mSessionStartGridLineStyle_PropertyChanged"
On Error GoTo Err

firePropertyChanged "SessionStartGridLineStyle"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mYAxisTextStyle Event Handlers
'@================================================================================

Private Sub mYAxisTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mYAxisTextStyle_PropertyChanged"
On Error GoTo Err

firePropertyChanged "YAxisTextStyle"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mYCursorTextStyle Event Handlers
'@================================================================================

Private Sub mYCursorTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mYCursorTextStyle_PropertyChanged"
On Error GoTo Err

firePropertyChanged "YCursorTextStyle"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Autoscaling() As Boolean
Autoscaling = mAutoscaling
End Property

Public Property Let Autoscaling(ByVal Value As Boolean)
Const ProcName As String = "Autoscaling"

On Error GoTo Err

mAutoscaling = Value
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingAutoscaling, mAutoscaling
firePropertyChanged "Autoscaling"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BackColor() As Long
BackColor = mBackGradientFillColors(0)
End Property

Public Property Let BackColor(ByVal Value As Long)
Const ProcName As String = "BackColor"

On Error GoTo Err

ReDim mBackGradientFillColors(0) As Long
mBackGradientFillColors(0) = Value
If Not mConfig Is Nothing Then setBackGradientFillColorsConfig
firePropertyChanged "BackColor"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BackGradientFillColors() As Long()
BackGradientFillColors = mBackGradientFillColors
End Property

Public Property Let BackGradientFillColors(ByRef Value() As Long)
Dim numColors As Long

Const ProcName As String = "BackGradientFillColors"

On Error GoTo Err

On Error Resume Next
numColors = UBound(Value) + 1
On Error GoTo Err

If numColors = 0 Then
    Erase mBackGradientFillColors
    Exit Property
End If

If numColors < 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "At least 1 background gradient fill Color must be supplied"
End If

If numColors > 5 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "At most 5 background gradient fill colors may be supplied"
End If

mBackGradientFillColors = Value

If Not mConfig Is Nothing Then setBackGradientFillColorsConfig

firePropertyChanged "BackGradientFillColors"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)

Const ProcName As String = "ConfigurationSection"

On Error GoTo Err

If Value Is Nothing Then
    RemoveFromConfig
    Set mConfig = Nothing
    Exit Property
End If

Set mConfig = Value

mConfig.SetSetting ConfigSettingName, mName
mConfig.SetSetting ConfigSettingStyleType, mStyleType

mGridLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionGridlineStyle)
mGridTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionGridTextStyle)
mSessionEndGridlineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionSessionEndGridlineStyle)
mSessionStartGridlineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionSessionStartGridlineStyle)
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

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let CursorSnapsToTickBoundaries(ByVal Value As Boolean)
Const ProcName As String = "CursorSnapsToTickBoundaries"

On Error GoTo Err

mCursorSnapsToTickBoundaries = Value
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingCursorSnapsToTickBoundaries, mCursorSnapsToTickBoundaries
firePropertyChanged "CursorSnapsToTickBoundaries"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CursorSnapsToTickBoundaries() As Boolean
CursorSnapsToTickBoundaries = mCursorSnapsToTickBoundaries
End Property

Public Property Get GridlineSpacingY() As Double
GridlineSpacingY = mGridlineSpacingY
End Property

Public Property Let GridlineSpacingY(ByVal Value As Double)
Const ProcName As String = "GridlineSpacingY"

On Error GoTo Err

mGridlineSpacingY = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingGridlineSpacingY, mGridlineSpacingY
firePropertyChanged "GridlineSpacingY"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get GridLineStyle() As LineStyle
Set GridLineStyle = mGridLineStyle
End Property

Public Property Let GridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "GridLineStyle"

On Error GoTo Err

Set mGridLineStyle = Value
If Not mConfig Is Nothing Then mGridLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionGridlineStyle)
firePropertyChanged "GridLineStyle"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get GridTextStyle() As TextStyle
Set GridTextStyle = mGridTextStyle
End Property

Public Property Let GridTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "GridTextStyle"

On Error GoTo Err

Set mGridTextStyle = Value
If Not mConfig Is Nothing Then mGridTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionGridTextStyle)
firePropertyChanged "GridTextStyle"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HasGrid() As Boolean
HasGrid = mHasGrid
End Property

Public Property Let HasGrid(ByVal val As Boolean)
Const ProcName As String = "HasGrid"

On Error GoTo Err

mHasGrid = val
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingHasGrid, mHasGrid
firePropertyChanged "HasGrid"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HasGridText() As Boolean
HasGridText = mHasGridText
End Property

Public Property Let HasGridText(ByVal val As Boolean)
Const ProcName As String = "HasGridText"

On Error GoTo Err

mHasGridText = val
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingHasGridText, mHasGridText
firePropertyChanged "HasGridText"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IntegerYScale() As Boolean
IntegerYScale = mIntegerYScale
End Property

Public Property Let IntegerYScale(ByVal Value As Boolean)
Const ProcName As String = "IntegerYScale"

On Error GoTo Err

mIntegerYScale = Value
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingIntegerYScale, mIntegerYScale
firePropertyChanged "IntegerYScale"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MinimumHeight() As Double
MinimumHeight = mMinimumHeight
End Property

Public Property Let MinimumHeight(ByVal Value As Double)
Const ProcName As String = "MinimumHeight"

On Error GoTo Err

mMinimumHeight = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingMinimumHeight, mMinimumHeight
firePropertyChanged "MinimumHeight"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Name(ByVal Value As String)
Const ProcName As String = "Name"
On Error GoTo Err

If Value = "" Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Name cannot be blank"
End If

mName = Value
If Not mConfig Is Nothing Then mConfig.InstanceQualifier = mName

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get SessionEndGridLineStyle() As LineStyle
Set SessionEndGridLineStyle = mSessionEndGridlineStyle
End Property

Public Property Let SessionEndGridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "SessionEndGridLineStyle"

On Error GoTo Err

Set mSessionEndGridlineStyle = Value
If Not mConfig Is Nothing Then mGridLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionSessionEndGridlineStyle)
firePropertyChanged "SessionEndGridLineStyle"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get SessionStartGridLineStyle() As LineStyle
Set SessionStartGridLineStyle = mSessionStartGridlineStyle
End Property

Public Property Let SessionStartGridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "SessionStartGridLineStyle"

On Error GoTo Err

Set mSessionStartGridlineStyle = Value
If Not mConfig Is Nothing Then mGridLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionSessionStartGridlineStyle)
firePropertyChanged "SessionStartGridLineStyle"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YAxisTextStyle() As TextStyle
Set YAxisTextStyle = mYAxisTextStyle
End Property

Public Property Let YAxisTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "YAxisTextStyle"

On Error GoTo Err

Set mYAxisTextStyle = Value
If Not mConfig Is Nothing Then mYAxisTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionYAxisTextStyle)
firePropertyChanged "YAxisTextStyle"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YCursorTextStyle() As TextStyle
Set YCursorTextStyle = mYCursorTextStyle
End Property

Public Property Let YCursorTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "YCursorTextStyle"

On Error GoTo Err

Set mYCursorTextStyle = Value
If Not mConfig Is Nothing Then mYCursorTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionYCursorTextStyle)
firePropertyChanged "YCursorTextStyle"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YScaleQuantum() As Double
YScaleQuantum = mYScaleQuantum
End Property

Public Property Let YScaleQuantum(ByVal Value As Double)
Const ProcName As String = "YScaleQuantum"

On Error GoTo Err

mYScaleQuantum = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingYScaleQuantum, mYScaleQuantum
firePropertyChanged "YScaleQuantum"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As ChartRegionStyle
Const ProcName As String = "clone"

On Error GoTo Err

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
clone.SessionEndGridLineStyle = mSessionEndGridlineStyle
clone.SessionStartGridLineStyle = mSessionStartGridlineStyle
clone.YAxisTextStyle = mYAxisTextStyle
clone.YCursorTextStyle = mYCursorTextStyle
clone.YScaleQuantum = mYScaleQuantum

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)
Dim ar() As String
Dim i As Long

Const ProcName As String = "LoadFromConfig"

On Error GoTo Err

Set mConfig = pConfig

getGridLineStyleFromConfig
getGridTextStyleFromConfig
getSessionEndGridLineStyleFromConfig
getSessionStartGridLineStyleFromConfig
getYAxisTextStyleFromConfig
getYCursorTextStyleFromConfig

mName = mConfig.GetSetting(ConfigSettingName)
mStyleType = mConfig.GetSetting(ConfigSettingStyleType)
HasGrid = CBool(mConfig.GetSetting(ConfigSettingHasGrid, "True"))
HasGridText = CBool(mConfig.GetSetting(ConfigSettingHasGridText, "False"))
Autoscaling = CBool(mConfig.GetSetting(ConfigSettingAutoscaling, "True"))
IntegerYScale = CBool(mConfig.GetSetting(ConfigSettingIntegerYScale, "False"))
YScaleQuantum = CDbl(mConfig.GetSetting(ConfigSettingYScaleQuantum, "0.01"))
GridlineSpacingY = CDbl(mConfig.GetSetting(ConfigSettingGridlineSpacingY, "1.8"))
MinimumHeight = CDbl(mConfig.GetSetting(ConfigSettingMinimumHeight, "0"))
CursorSnapsToTickBoundaries = CBool(mConfig.GetSetting(ConfigSettingCursorSnapsToTickBoundaries, "True"))

ar = Split(mConfig.GetSetting(ConfigSettingBackGradientFillColors), ",")
ReDim mBackGradientFillColors(UBound(ar())) As Long
For i = 0 To UBound(ar())
    mBackGradientFillColors(i) = CLng(ar(i))
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
                
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"

On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub getGridLineStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getGridLineStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionGridlineStyle)
If Not cs Is Nothing Then
    Set mGridLineStyle = New LineStyle
    mGridLineStyle.LoadFromConfig cs
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getGridTextStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getGridTextStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionGridTextStyle)
If Not cs Is Nothing Then
    Set mGridTextStyle = New TextStyle
    mGridTextStyle.LoadFromConfig cs
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getSessionEndGridLineStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getSessionEndGridLineStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionSessionEndGridlineStyle)
If Not cs Is Nothing Then
    Set mSessionEndGridlineStyle = New LineStyle
    mSessionEndGridlineStyle.LoadFromConfig cs
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getSessionStartGridLineStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getSessionStartGridLineStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionSessionStartGridlineStyle)
If Not cs Is Nothing Then
    Set mSessionStartGridlineStyle = New LineStyle
    mSessionStartGridlineStyle.LoadFromConfig cs
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getYAxisTextStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getYAxisTextStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionYAxisTextStyle)
If Not cs Is Nothing Then
    Set mYAxisTextStyle = New TextStyle
    mYAxisTextStyle.LoadFromConfig cs
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getYCursorTextStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getYCursorTextStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionYCursorTextStyle)
If Not cs Is Nothing Then
    Set mYCursorTextStyle = New TextStyle
    mYCursorTextStyle.LoadFromConfig cs
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub firePropertyChanged( _
                ByVal Name As String)
Dim ev As PropertyChangedEventData

Const ProcName As String = "firePropertyChanged"

On Error GoTo Err

Set ev.Source = Me
ev.PropertyName = Name
RaiseEvent PropertyChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setBackGradientFillColorsConfig()
Dim i As Long
Dim s As String
Const ProcName As String = "setBackGradientFillColorsConfig"

On Error GoTo Err

For i = 0 To UBound(mBackGradientFillColors)
    If i <> 0 Then s = s & ","
    s = s & CStr(mBackGradientFillColors(i))
Next
mConfig.SetSetting ConfigSettingBackGradientFillColors, s

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setInitialGridLineStyle()
Const ProcName As String = "setInitialGridLineStyle"

On Error GoTo Err

Set mGridLineStyle = New LineStyle
mGridLineStyle.Color = &HD0D0D0

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setInitialGridTextStyle()
Dim afont As StdFont

Const ProcName As String = "setInitialGridTextStyle"

On Error GoTo Err

Set mGridTextStyle = New TextStyle
mGridTextStyle.Box = True
mGridTextStyle.BoxFillWithBackgroundColor = True
mGridTextStyle.BoxStyle = LineInvisible
mGridTextStyle.Color = vbBlack
    
Set afont = New StdFont
afont.Italic = False
afont.Name = "Arial"
afont.Size = 8
afont.Underline = False
afont.Bold = False
mGridTextStyle.Font = afont

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub setInitialSessionEndGridLineStyle()
Const ProcName As String = "setInitialSessionEndGridLineStyle"

On Error GoTo Err

Set mSessionEndGridlineStyle = New LineStyle
mSessionEndGridlineStyle.Color = &HD0D0D0
mSessionEndGridlineStyle.LineStyle = LineDash

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setInitialSessionStartGridLineStyle()
Const ProcName As String = "setInitialSessionStartGridLineStyle"

On Error GoTo Err

Set mSessionStartGridlineStyle = New LineStyle
mSessionStartGridlineStyle.Color = &HD0D0D0
mSessionStartGridlineStyle.Thickness = 3

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setInitialYAxisTextStyle()
Dim afont As StdFont

Const ProcName As String = "setInitialYAxisTextStyle"

On Error GoTo Err

Set mYAxisTextStyle = New TextStyle
mYAxisTextStyle.Box = True
mYAxisTextStyle.BoxFillWithBackgroundColor = True
mYAxisTextStyle.BoxStyle = LineInvisible
mYAxisTextStyle.Color = vbBlack
    
Set afont = New StdFont
afont.Italic = False
afont.Name = "Arial"
afont.Size = 8
afont.Underline = False
afont.Bold = False
mYAxisTextStyle.Font = afont

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub setInitialYCursorTextStyle()
Dim afont As StdFont

Const ProcName As String = "setInitialYCursorTextStyle"

On Error GoTo Err

Set mYCursorTextStyle = New TextStyle
mYCursorTextStyle.Box = True
mYCursorTextStyle.BoxColor = vbBlack
mYCursorTextStyle.BoxFillColor = vbWhite
mYCursorTextStyle.BoxFillStyle = FillSolid
mYCursorTextStyle.BoxThickness = 1
mYCursorTextStyle.Color = vbBlack
mYCursorTextStyle.PaddingX = 1
mYCursorTextStyle.PaddingY = 0
mYCursorTextStyle.Justification = JustifyCentre
    
Set afont = New StdFont
afont.Italic = False
afont.Name = "Arial"
afont.Size = 8
afont.Underline = False
afont.Bold = False
mYCursorTextStyle.Font = afont

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

