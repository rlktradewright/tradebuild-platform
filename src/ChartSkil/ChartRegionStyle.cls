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

Implements Configurable
'Implements IGraphicObjectStyle

'@================================================================================
' Events
'@================================================================================

Event PropertyChanged( _
                ByRef ev As PropertyChangedEventData)
                
'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                As String = "ChartRegionStyle"

Private Const ConfigSectionCursorTextStyle              As String = "CursorTextStyle"
Private Const ConfigSectionXGridlineStyle               As String = "XGridlineStyle"
Private Const ConfigSectionXGridTextStyle               As String = "XGridTextStyle"
Private Const ConfigSectionYGridlineStyle               As String = "YGridlineStyle"
Private Const ConfigSectionYGridTextStyle               As String = "YGridTextStyle"
Private Const ConfigSectionXCursorTextStyle             As String = "XCursorTextStyle"
Private Const ConfigSectionYCursorTextStyle             As String = "YCursorTextStyle"
Private Const ConfigSectionSessionEndGridlineStyle      As String = "SessionEndGridlineStyle"
Private Const ConfigSectionSessionStartGridlineStyle    As String = "SessionStartGridlineStyle"

Private Const ConfigSettingCursorTextMode               As String = "&CursorTextMode"
Private Const ConfigSettingCursorTextPosition           As String = "&CursorTextPosition"
Private Const ConfigSettingHasXGrid                     As String = "&HasXGrid"
Private Const ConfigSettingHasXGridText                 As String = "&HasXGridText"
Private Const ConfigSettingHasYGrid                     As String = "&HasYGrid"
Private Const ConfigSettingHasYGridText                 As String = "&HasYGridText"
Private Const ConfigSettingAutoscaling                  As String = "&Autoscale"
Private Const ConfigSettingIntegerYScale                As String = "&IntegerYScale"
Private Const ConfigSettingYScaleQuantum                As String = "&YScaleQuantum"
Private Const ConfigSettingYGridlineSpacing             As String = "&YGridlineSpacing"
Private Const ConfigSettingMinimumHeight                As String = "&MinimumHeight"
Private Const ConfigSettingCursorSnapsToTickBoundaries  As String = "&CursorSnapsToTickBoundaries"
Private Const ConfigSettingBackGradientFillColors       As String = "&BackGradientFillColors"
Private Const ConfigSettingXGridTextPosition            As String = "&XGridTextPosition"
Private Const ConfigSettingYGridTextPosition            As String = "&YGridTextPosition"
Private Const ConfigSettingXCursorTextPosition          As String = "&XCursorTextPosition"
Private Const ConfigSettingYCursorTextPosition          As String = "&YCursorTextPosition"

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
Private mHasXGrid As Boolean
Private mHasYGrid As Boolean
Private mHasXGridText As Boolean
Private mHasYGridText As Boolean

Private mXGridTextPosition As XGridTextPositions
Private mYGridTextPosition As YGridTextPositions

Private mCursorTextMode As CursorTextModes
Private mCursorTextPosition As CursorTextPositions

Private mXCursorTextPosition As CursorTextPositions

Private mYCursorTextPosition As CursorTextPositions

Private mAutoscaling As Boolean
Private mIntegerYScale As Boolean
Private mYScaleQuantum As Double
Private mYGridlineSpacing As Double
Private mMinimumHeight As Double
Private mCursorSnapsToTickBoundaries As Boolean

Private mBackGradientFillColors() As Long

Private WithEvents mXGridLineStyle As LineStyle
Attribute mXGridLineStyle.VB_VarHelpID = -1
Private WithEvents mXGridTextStyle As TextStyle
Attribute mXGridTextStyle.VB_VarHelpID = -1
Private WithEvents mYGridLineStyle As LineStyle
Attribute mYGridLineStyle.VB_VarHelpID = -1
Private WithEvents mYGridTextStyle As TextStyle
Attribute mYGridTextStyle.VB_VarHelpID = -1
Private WithEvents mSessionEndGridLineStyle As LineStyle
Attribute mSessionEndGridLineStyle.VB_VarHelpID = -1
Private WithEvents mSessionStartGridLineStyle As LineStyle
Attribute mSessionStartGridLineStyle.VB_VarHelpID = -1
Private WithEvents mCursorTextStyle As TextStyle
Attribute mCursorTextStyle.VB_VarHelpID = -1
Private WithEvents mXCursorTextStyle As TextStyle
Attribute mXCursorTextStyle.VB_VarHelpID = -1
Private WithEvents mYCursorTextStyle As TextStyle
Attribute mYCursorTextStyle.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mName = GenerateGUIDString
mStyleType = "Chart Region Styles"

'setInitialXGridLineStyle
'setInitialXGridTextStyle
'setInitialYGridLineStyle
'setInitialYGridTextStyle
'setInitialSessionEndGridLineStyle
'setInitialSessionStartGridLineStyle
'setInitialCursorTextStyle
'setInitialXCursorTextStyle
'setInitialYCursorTextStyle

mHasXGrid = False
mHasXGridText = False
mHasYGrid = False
mHasYGridText = False

mXGridTextPosition = XGridTextPositionBottom
mYGridTextPosition = YGridTextPositionCentre

mCursorTextMode = CursorTextModeNone
mCursorTextPosition = CursorTextPositionLeftCursor

mXCursorTextPosition = CursorTextPositionBottomLeftFixed
mYCursorTextPosition = CursorTextPositionTopLeftFixed

mAutoscaling = True
mIntegerYScale = False
mYScaleQuantum = 0.0001
mYGridlineSpacing = 1.8
mMinimumHeight = 0.0001
mCursorSnapsToTickBoundaries = True
ReDim mBackGradientFillColors(0) As Long
mBackGradientFillColors(0) = vbWhite
End Sub

'@================================================================================
' Configurable Interface Members
'@================================================================================

Private Property Let Configurable_ConfigurationSection(ByVal RHS As ConfigurationSection)
Const ProcName As String = "Configurable_ConfigurationSection"
On Error GoTo Err

ConfigurationSection = RHS

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub Configurable_LoadFromConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "Configurable_LoadFromConfig"
On Error GoTo Err

LoadFromConfig pConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Configurable_RemoveFromConfig()
Const ProcName As String = "Configurable_RemoveFromConfig"
On Error GoTo Err

RemoveFromConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
' mCursorTextStyle Event Handlers
'@================================================================================

Private Sub mCursorTextStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mCursorTextStyle_PropertyChanged"
On Error GoTo Err

firePropertyChanged "CursorTextStyle"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mSessionEndGridLineStyle Event Handlers
'@================================================================================

Private Sub mSessionEndGridLineStyle_PropertyChanged(ev As PropertyChangedEventData)
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

Private Sub mSessionStartGridLineStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mSessionStartGridLineStyle_PropertyChanged"
On Error GoTo Err

firePropertyChanged "SessionStartGridLineStyle"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mXCursorTextStyle Event Handlers
'@================================================================================

Private Sub mXCursorTextStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mXCursorTextStyle_PropertyChanged"
On Error GoTo Err

firePropertyChanged "XCursorTextStyle"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mXGridLineStyle Event Handlers
'@================================================================================

Private Sub mXGridLineStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mXGridLineStyle_PropertyChanged"
On Error GoTo Err

firePropertyChanged "XGridLineStyle"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mXGridTextStyle Event Handlers
'@================================================================================

Private Sub mXGridTextStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mXGridTextStyle_PropertyChanged"
On Error GoTo Err

firePropertyChanged "XGridTextStyle"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mYCursorTextStyle Event Handlers
'@================================================================================

Private Sub mYCursorTextStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mYCursorTextStyle_PropertyChanged"
On Error GoTo Err

firePropertyChanged "YCursorTextStyle"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mYGridLineStyle Event Handlers
'@================================================================================

Private Sub mYXGridLineStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mYGridLineStyle_PropertyChanged"
On Error GoTo Err

firePropertyChanged "YGridLineStyle"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mYGridTextStyle Event Handlers
'@================================================================================

Private Sub mYGridTextStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mYGridTextStyle_PropertyChanged"
On Error GoTo Err

firePropertyChanged "YGridTextStyle"

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
gHandleUnexpectedError ProcName, ModuleName
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

If numColors < 1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "At least 1 background gradient fill Color must be supplied"

If numColors > 5 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "At most 5 background gradient fill colors may be supplied"

mBackGradientFillColors = Value

If Not mConfig Is Nothing Then setBackGradientFillColorsConfig

firePropertyChanged "BackGradientFillColors"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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

If Not mXGridLineStyle Is Nothing Then mXGridLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXGridlineStyle)
If Not mXGridTextStyle Is Nothing Then mXGridTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXGridTextStyle)
If Not mYGridLineStyle Is Nothing Then mYGridLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionYGridlineStyle)
If Not mYGridTextStyle Is Nothing Then mYGridTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionYGridTextStyle)
If Not mSessionEndGridLineStyle Is Nothing Then mSessionEndGridLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionSessionEndGridlineStyle)
If Not mSessionStartGridLineStyle Is Nothing Then mSessionStartGridLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionSessionStartGridlineStyle)
If Not mCursorTextStyle Is Nothing Then mCursorTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionCursorTextStyle)
If Not mXCursorTextStyle Is Nothing Then mXCursorTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXCursorTextStyle)
If Not mYCursorTextStyle Is Nothing Then mYCursorTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionYCursorTextStyle)

mConfig.SetSetting ConfigSettingHasXGrid, mHasXGrid
mConfig.SetSetting ConfigSettingHasXGridText, mHasXGridText
mConfig.SetSetting ConfigSettingHasYGrid, mHasYGrid
mConfig.SetSetting ConfigSettingHasYGridText, mHasYGridText
mConfig.SetSetting ConfigSettingCursorTextMode, mCursorTextMode
mConfig.SetSetting ConfigSettingCursorTextPosition, mCursorTextPosition
mConfig.SetSetting ConfigSettingXCursorTextPosition, mXCursorTextPosition
mConfig.SetSetting ConfigSettingYCursorTextPosition, mYCursorTextPosition
mConfig.SetSetting ConfigSettingXGridTextPosition, mXGridTextPosition
mConfig.SetSetting ConfigSettingYGridTextPosition, mYGridTextPosition
mConfig.SetSetting ConfigSettingAutoscaling, mAutoscaling
mConfig.SetSetting ConfigSettingIntegerYScale, mIntegerYScale
mConfig.SetSetting ConfigSettingYScaleQuantum, mYScaleQuantum
mConfig.SetSetting ConfigSettingYGridlineSpacing, mYGridlineSpacing
mConfig.SetSetting ConfigSettingMinimumHeight, mMinimumHeight
mConfig.SetSetting ConfigSettingCursorSnapsToTickBoundaries, mCursorSnapsToTickBoundaries

setBackGradientFillColorsConfig

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let CursorSnapsToTickBoundaries(ByVal Value As Boolean)
Const ProcName As String = "CursorSnapsToTickBoundaries"

On Error GoTo Err

mCursorSnapsToTickBoundaries = Value
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingCursorSnapsToTickBoundaries, mCursorSnapsToTickBoundaries
firePropertyChanged "CursorSnapsToTickBoundaries"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CursorSnapsToTickBoundaries() As Boolean
CursorSnapsToTickBoundaries = mCursorSnapsToTickBoundaries
End Property

Public Property Get CursorTextMode() As CursorTextModes
CursorTextMode = mCursorTextMode
End Property

Public Property Let CursorTextMode(ByVal val As CursorTextModes)
Const ProcName As String = "CursorTextMode"

On Error GoTo Err

mCursorTextMode = val
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingCursorTextMode, mCursorTextMode
firePropertyChanged "CursorTextMode"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CursorTextPosition() As CursorTextPositions
CursorTextPosition = mCursorTextPosition
End Property

Public Property Let CursorTextPosition(ByVal val As CursorTextPositions)
Const ProcName As String = "CursorTextPosition"

On Error GoTo Err

mCursorTextPosition = val
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingCursorTextPosition, mCursorTextPosition
firePropertyChanged "CursorTextPosition"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CursorTextStyle() As TextStyle
Set CursorTextStyle = mCursorTextStyle
End Property

Public Property Let CursorTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "CursorTextStyle"

On Error GoTo Err

Set mCursorTextStyle = Value
If Not mConfig Is Nothing Then mCursorTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionCursorTextStyle)
firePropertyChanged "CursorTextStyle"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YGridlineSpacing() As Double
YGridlineSpacing = mYGridlineSpacing
End Property

Public Property Let YGridlineSpacing(ByVal Value As Double)
Const ProcName As String = "YGridlineSpacing"

On Error GoTo Err

mYGridlineSpacing = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingYGridlineSpacing, mYGridlineSpacing
firePropertyChanged "YGridlineSpacing"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HasXGrid() As Boolean
HasXGrid = mHasXGrid
End Property

Public Property Let HasXGrid(ByVal val As Boolean)
Const ProcName As String = "HasXGrid"

On Error GoTo Err

mHasXGrid = val
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingHasXGrid, mHasXGrid
firePropertyChanged "HasXGrid"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HasXGridText() As Boolean
HasXGridText = mHasXGridText
End Property

Public Property Let HasXGridText(ByVal val As Boolean)
Const ProcName As String = "HasXGridText"

On Error GoTo Err

mHasXGridText = val
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingHasXGridText, mHasXGridText
firePropertyChanged "HasXGridText"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HasYGrid() As Boolean
HasYGrid = mHasYGrid
End Property

Public Property Let HasYGrid(ByVal val As Boolean)
Const ProcName As String = "HasYGrid"

On Error GoTo Err

mHasYGrid = val
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingHasYGrid, mHasYGrid
firePropertyChanged "HasYGrid"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HasYGridText() As Boolean
HasYGridText = mHasYGridText
End Property

Public Property Let HasYGridText(ByVal val As Boolean)
Const ProcName As String = "HasYGridText"

On Error GoTo Err

mHasYGridText = val
If Not mConfig Is Nothing Then mConfig.SetBooleanSetting ConfigSettingHasYGridText, mHasYGridText
firePropertyChanged "HasYGridText"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Name(ByVal Value As String)
Const ProcName As String = "Name"
On Error GoTo Err

If Value = "" Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Name cannot be blank"

mName = Value
If Not mConfig Is Nothing Then mConfig.InstanceQualifier = mName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Get SessionEndGridLineStyle() As LineStyle
Set SessionEndGridLineStyle = mSessionEndGridLineStyle
End Property

Public Property Let SessionEndGridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "SessionEndGridLineStyle"

On Error GoTo Err

Set mSessionEndGridLineStyle = Value
If Not mConfig Is Nothing Then mSessionEndGridLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionSessionEndGridlineStyle)
firePropertyChanged "SessionEndGridLineStyle"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SessionStartGridLineStyle() As LineStyle
Set SessionStartGridLineStyle = mSessionStartGridLineStyle
End Property

Public Property Let SessionStartGridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "SessionStartGridLineStyle"

On Error GoTo Err

Set mSessionStartGridLineStyle = Value
If Not mConfig Is Nothing Then mSessionStartGridLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionSessionStartGridlineStyle)
firePropertyChanged "SessionStartGridLineStyle"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XCursorTextPosition() As CursorTextPositions
XCursorTextPosition = mXCursorTextPosition
End Property

Public Property Let XCursorTextPosition(ByVal Value As CursorTextPositions)
Const ProcName As String = "XCursorTextPosition"

On Error GoTo Err

mXCursorTextPosition = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingXCursorTextPosition, mXCursorTextPosition
firePropertyChanged "XCursorTextPosition"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XCursorTextStyle() As TextStyle
Set XCursorTextStyle = mXCursorTextStyle
End Property

Public Property Let XCursorTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "XCursorTextStyle"

On Error GoTo Err

Set mXCursorTextStyle = Value
If Not mConfig Is Nothing Then mXCursorTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXCursorTextStyle)
firePropertyChanged "XCursorTextStyle"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XGridLineStyle() As LineStyle
Set XGridLineStyle = mXGridLineStyle
End Property

Public Property Let XGridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "XGridLineStyle"

On Error GoTo Err

Set mXGridLineStyle = Value
If Not mConfig Is Nothing Then mXGridLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXGridlineStyle)
firePropertyChanged "XGridLineStyle"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XGridTextPosition() As XGridTextPositions
XGridTextPosition = mXGridTextPosition
End Property

Public Property Let XGridTextPosition(ByVal Value As XGridTextPositions)
Const ProcName As String = "XGridTextPosition"

On Error GoTo Err

mXGridTextPosition = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingXGridTextPosition, mXGridTextPosition
firePropertyChanged "XGridTextPosition"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XGridTextStyle() As TextStyle
Set XGridTextStyle = mXGridTextStyle
End Property

Public Property Let XGridTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "XGridTextStyle"

On Error GoTo Err

Set mXGridTextStyle = Value
If Not mConfig Is Nothing Then mXGridTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXGridTextStyle)
firePropertyChanged "XGridTextStyle"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YCursorTextPosition() As CursorTextPositions
YCursorTextPosition = mYCursorTextPosition
End Property

Public Property Let YCursorTextPosition(ByVal Value As CursorTextPositions)
Const ProcName As String = "YCursorTextPosition"

On Error GoTo Err

mYCursorTextPosition = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingYCursorTextPosition, mYCursorTextPosition
firePropertyChanged "YCursorTextPosition"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YGridLineStyle() As LineStyle
Set YGridLineStyle = mYGridLineStyle
End Property

Public Property Let YGridLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "YGridLineStyle"

On Error GoTo Err

Set mYGridLineStyle = Value
If Not mConfig Is Nothing Then mYGridLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionYGridlineStyle)
firePropertyChanged "YGridLineStyle"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YGridTextPosition() As YGridTextPositions
YGridTextPosition = mYGridTextPosition
End Property

Public Property Let YGridTextPosition(ByVal Value As YGridTextPositions)
Const ProcName As String = "YGridTextPosition"

On Error GoTo Err

mYGridTextPosition = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingYGridTextPosition, mYGridTextPosition
firePropertyChanged "YGridTextPosition"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YGridTextStyle() As TextStyle
Set YGridTextStyle = mYGridTextStyle
End Property

Public Property Let YGridTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "YGridTextStyle"

On Error GoTo Err

Set mYGridTextStyle = Value
If Not mConfig Is Nothing Then mYGridTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionYGridTextStyle)
firePropertyChanged "YGridTextStyle"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
clone.CursorSnapsToTickBoundaries = mCursorSnapsToTickBoundaries
clone.CursorTextMode = mCursorTextMode
clone.CursorTextPosition = mCursorTextPosition
If Not mCursorTextStyle Is Nothing Then clone.CursorTextStyle = mCursorTextStyle.clone
clone.HasXGrid = mHasXGrid
clone.HasXGridText = mHasXGridText
clone.HasYGrid = mHasYGrid
clone.HasYGridText = mHasYGridText
clone.IntegerYScale = mIntegerYScale
clone.MinimumHeight = mMinimumHeight
If Not mSessionEndGridLineStyle Is Nothing Then clone.SessionEndGridLineStyle = mSessionEndGridLineStyle.clone
If Not mSessionStartGridLineStyle Is Nothing Then clone.SessionStartGridLineStyle = mSessionStartGridLineStyle.clone
clone.XCursorTextPosition = mXCursorTextPosition
If Not mXCursorTextStyle Is Nothing Then clone.XCursorTextStyle = mXCursorTextStyle.clone
If Not mXGridLineStyle Is Nothing Then clone.XGridLineStyle = mXGridLineStyle.clone
clone.XGridTextPosition = mXGridTextPosition
If Not mXGridTextStyle Is Nothing Then clone.XGridTextStyle = mXGridTextStyle.clone
clone.YCursorTextPosition = mYCursorTextPosition
If Not mYCursorTextStyle Is Nothing Then clone.YCursorTextStyle = mYCursorTextStyle.clone
clone.YGridlineSpacing = mYGridlineSpacing
If Not mYGridLineStyle Is Nothing Then clone.YGridLineStyle = mYGridLineStyle.clone
clone.YGridTextPosition = mYGridTextPosition
If Not mYGridTextStyle Is Nothing Then clone.YGridTextStyle = mYGridTextStyle.clone
clone.YScaleQuantum = mYScaleQuantum

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)
Dim ar() As String
Dim i As Long

Const ProcName As String = "LoadFromConfig"

On Error GoTo Err

Set mConfig = pConfig

getXGridLineStyleFromConfig
getXGridTextStyleFromConfig
getYGridLineStyleFromConfig
getYGridTextStyleFromConfig
getSessionEndGridLineStyleFromConfig
getSessionStartGridLineStyleFromConfig
getCursorTextStyleFromConfig
getXCursorTextStyleFromConfig
getYCursorTextStyleFromConfig

mName = mConfig.GetSetting(ConfigSettingName)
mStyleType = mConfig.GetSetting(ConfigSettingStyleType)
HasXGrid = CBool(mConfig.GetSetting(ConfigSettingHasXGrid, "True"))
HasXGridText = CBool(mConfig.GetSetting(ConfigSettingHasXGridText, "False"))
CursorTextMode = CInt(mConfig.GetSetting(ConfigSettingCursorTextMode, CStr(CursorTextModes.CursorTextModeNone)))
CursorTextPosition = CInt(mConfig.GetSetting(ConfigSettingCursorTextPosition, CStr(CursorTextPositions.CursorTextPositionLeftCursor)))
HasYGrid = CBool(mConfig.GetSetting(ConfigSettingHasYGrid, "True"))
HasYGridText = CBool(mConfig.GetSetting(ConfigSettingHasYGridText, "False"))
XGridTextPosition = CLng(mConfig.GetSetting(ConfigSettingXGridTextPosition, CStr(XGridTextPositions.XGridTextPositionBottom)))
YGridTextPosition = CLng(mConfig.GetSetting(ConfigSettingYGridTextPosition, CStr(YGridTextPositions.YGridTextPositionCentre)))
XCursorTextPosition = CLng(mConfig.GetSetting(ConfigSettingXCursorTextPosition, CStr(CursorTextPositions.CursorTextPositionBottomLeftFixed)))
YCursorTextPosition = CLng(mConfig.GetSetting(ConfigSettingYCursorTextPosition, CStr(CursorTextPositions.CursorTextPositionTopLeftFixed)))
Autoscaling = CBool(mConfig.GetSetting(ConfigSettingAutoscaling, "True"))
IntegerYScale = CBool(mConfig.GetSetting(ConfigSettingIntegerYScale, "False"))
YScaleQuantum = CDbl(mConfig.GetSetting(ConfigSettingYScaleQuantum, "0.01"))
YGridlineSpacing = CDbl(mConfig.GetSetting(ConfigSettingYGridlineSpacing, "1.8"))
MinimumHeight = CDbl(mConfig.GetSetting(ConfigSettingMinimumHeight, "0"))
CursorSnapsToTickBoundaries = CBool(mConfig.GetSetting(ConfigSettingCursorSnapsToTickBoundaries, "True"))

ar = Split(mConfig.GetSetting(ConfigSettingBackGradientFillColors), ",")
ReDim mBackGradientFillColors(UBound(ar())) As Long
For i = 0 To UBound(ar())
    mBackGradientFillColors(i) = CLng(ar(i))
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
                
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"

On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub getCursorTextStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getCursorTextStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionCursorTextStyle)
If Not cs Is Nothing Then
    Set mCursorTextStyle = New TextStyle
    mCursorTextStyle.LoadFromConfig cs
Else
    Set mCursorTextStyle = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub getSessionEndGridLineStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getSessionEndGridLineStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionSessionEndGridlineStyle)
If Not cs Is Nothing Then
    Set mSessionEndGridLineStyle = New LineStyle
    mSessionEndGridLineStyle.LoadFromConfig cs
Else
    Set mSessionEndGridLineStyle = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub getSessionStartGridLineStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getSessionStartGridLineStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionSessionStartGridlineStyle)
If Not cs Is Nothing Then
    Set mSessionStartGridLineStyle = New LineStyle
    mSessionStartGridLineStyle.LoadFromConfig cs
Else
    Set mSessionStartGridLineStyle = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub getXCursorTextStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getXCursorTextStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionXCursorTextStyle)
If Not cs Is Nothing Then
    Set mXCursorTextStyle = New TextStyle
    mXCursorTextStyle.LoadFromConfig cs
Else
    Set mXCursorTextStyle = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub getXGridLineStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getXGridLineStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionXGridlineStyle)
If Not cs Is Nothing Then
    Set mXGridLineStyle = New LineStyle
    mXGridLineStyle.LoadFromConfig cs
Else
    Set mXGridLineStyle = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub getXGridTextStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getXGridTextStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionXGridTextStyle)
If Not cs Is Nothing Then
    Set mXGridTextStyle = New TextStyle
    mXGridTextStyle.LoadFromConfig cs
Else
    Set mXGridTextStyle = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub getYCursorTextStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getYCursorTextStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionYCursorTextStyle)
If Not cs Is Nothing Then
    Set mYCursorTextStyle = New TextStyle
    mYCursorTextStyle.LoadFromConfig cs
Else
    Set mYCursorTextStyle = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub getYGridLineStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getYGridLineStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionYGridlineStyle)
If Not cs Is Nothing Then
    Set mYGridLineStyle = New LineStyle
    mYGridLineStyle.LoadFromConfig cs
Else
    Set mYGridLineStyle = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub getYGridTextStyleFromConfig()
Dim cs As ConfigurationSection
Const ProcName As String = "getYGridTextStyleFromConfig"

On Error GoTo Err

Set cs = mConfig.GetConfigurationSection(ConfigSectionYGridTextStyle)
If Not cs Is Nothing Then
    Set mYGridTextStyle = New TextStyle
    mYGridTextStyle.LoadFromConfig cs
Else
    Set mYGridTextStyle = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setInitialCursorTextStyle()
Const ProcName As String = "setInitialCursorTextStyle"
On Error GoTo Err

Set mCursorTextStyle = gDefaultTextStyle.clone
mCursorTextStyle.Box = True
mCursorTextStyle.BoxColor = vbBlack
mCursorTextStyle.BoxFillColor = vbWhite
mCursorTextStyle.BoxFillStyle = FillSolid
mCursorTextStyle.BoxThickness = 1
mCursorTextStyle.Color = vbBlack
mCursorTextStyle.PaddingX = 1
mCursorTextStyle.PaddingY = 0
mCursorTextStyle.Justification = JustifyCentre
    
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setInitialXGridLineStyle()
Const ProcName As String = "setInitialXGridLineStyle"

On Error GoTo Err

Set mXGridLineStyle = gDefaultLineStyle.clone
mXGridLineStyle.Color = DefaultGridlineColor

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setInitialXGridTextStyle()
Const ProcName As String = "setInitialXGridTextStyle"

On Error GoTo Err

Set mXGridTextStyle = gDefaultTextStyle.clone
mXGridTextStyle.Box = True
mXGridTextStyle.BoxFillWithBackgroundColor = True
mXGridTextStyle.BoxStyle = LineInvisible
mXGridTextStyle.Color = vbBlack
    
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub setInitialYGridLineStyle()
Const ProcName As String = "setInitialYGridLineStyle"

On Error GoTo Err

Set mYGridLineStyle = gDefaultLineStyle.clone
mYGridLineStyle.Color = DefaultGridlineColor

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setInitialYGridTextStyle()
Const ProcName As String = "setInitialYGridTextStyle"

On Error GoTo Err

Set mYGridTextStyle = gDefaultTextStyle.clone
mYGridTextStyle.Box = True
mYGridTextStyle.BoxFillWithBackgroundColor = True
mYGridTextStyle.BoxStyle = LineInvisible
mYGridTextStyle.Color = vbBlack
    
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub setInitialSessionEndGridLineStyle()
Const ProcName As String = "setInitialSessionEndGridLineStyle"

On Error GoTo Err

Set mSessionEndGridLineStyle = gDefaultLineStyle.clone
mSessionEndGridLineStyle.Color = DefaultGridlineColor
mSessionEndGridLineStyle.LineStyle = LineDash

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setInitialSessionStartGridLineStyle()
Const ProcName As String = "setInitialSessionStartGridLineStyle"

On Error GoTo Err

Set mSessionStartGridLineStyle = gDefaultLineStyle.clone
mSessionStartGridLineStyle.Color = DefaultGridlineColor
mSessionStartGridLineStyle.Thickness = 3

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setInitialXCursorTextStyle()
Const ProcName As String = "setInitialXCursorTextStyle"
On Error GoTo Err

Set mXCursorTextStyle = gDefaultTextStyle.clone
mXCursorTextStyle.Box = True
mXCursorTextStyle.BoxColor = vbBlack
mXCursorTextStyle.BoxFillColor = vbWhite
mXCursorTextStyle.BoxFillStyle = FillSolid
mXCursorTextStyle.BoxThickness = 1
mXCursorTextStyle.Color = vbBlack
mXCursorTextStyle.PaddingX = 1
mXCursorTextStyle.PaddingY = 0
mXCursorTextStyle.Justification = JustifyCentre
    
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setInitialYCursorTextStyle()
Const ProcName As String = "setInitialYCursorTextStyle"
On Error GoTo Err

Set mYCursorTextStyle = gDefaultTextStyle.clone
mYCursorTextStyle.Box = True
mYCursorTextStyle.BoxColor = vbBlack
mYCursorTextStyle.BoxFillColor = vbWhite
mYCursorTextStyle.BoxFillStyle = FillSolid
mYCursorTextStyle.BoxThickness = 1
mYCursorTextStyle.Color = vbBlack
mYCursorTextStyle.PaddingX = 1
mYCursorTextStyle.PaddingY = 0
mYCursorTextStyle.Justification = JustifyCentre
    
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

