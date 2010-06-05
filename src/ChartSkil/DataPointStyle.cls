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

Implements IGraphicObjectStyle

'@================================================================================
' Events
'@================================================================================

Event PropertyChanged( _
                ByRef ev As PropertyChangedEvent)
                
'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "DataPointStyle"

Private Const ConfigSettingColor                As String = "&BarColor"
Private Const ConfigSettingUpColor              As String = "&UpColor"
Private Const ConfigSettingDownColor            As String = "&DownColor"
Private Const ConfigSettingDisplayMode          As String = "&DisplayMode"
Private Const ConfigSettingLineThickness        As String = "&LineThickness"
Private Const ConfigSettingLineStyle            As String = "&LineStyle"
Private Const ConfigSettingPointStyle           As String = "&PointStyle"
Private Const ConfigSettingHistWidth            As String = "&HistBarWidth"
Private Const ConfigSettingIncludeInAutoscale   As String = "&IncludeInAutoscale"
Private Const ConfigSettingLayer                 As String = "&Layer"

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

Private mLineThickness As Long
Private mColor As Long
Private mUpColor As Long
Private mDownColor As Long
Private mLineStyle As LineStyles
Private mPointStyle As PointStyles
Private mDisplayMode As DataPointDisplayModes
Private mHistWidth As Single
Private mIncludeInAutoscale As Boolean
Private mLayer As Long

Private mPropertyFlags As DataPointPropertyFlags

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mName = GenerateGUIDString
mStyleType = "Data Point Styles"
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
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Color(ByVal Value As Long)
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(Value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"

setLongProperty mColor, DataPointPropertyColor, Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Color() As Long
Const ProcName As String = "Color"
On Error GoTo Err

checkPropertySet DataPointPropertyColor
Color = mColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
Dim failpoint As String
On Error GoTo Err

If Value Is Nothing Then
    RemoveFromConfig
    Set mConfig = Nothing
    Exit Property
End If

Set mConfig = Value

mConfig.SetSetting ConfigSettingName, mName
mConfig.SetSetting ConfigSettingStyleType, mStyleType

If IsPropertySet(DataPointPropertyLineThickness) Then mConfig.SetSetting ConfigSettingLineThickness, mLineThickness
If IsPropertySet(DataPointPropertyColor) Then mConfig.SetSetting ConfigSettingColor, mColor
If IsPropertySet(DataPointPropertyUpColor) Then mConfig.SetSetting ConfigSettingUpColor, mUpColor
If IsPropertySet(DataPointPropertyDownColor) Then mConfig.SetSetting ConfigSettingDownColor, mDownColor
If IsPropertySet(DataPointPropertyLineStyle) Then mConfig.SetSetting ConfigSettingLineStyle, mLineStyle
If IsPropertySet(DataPointPropertyPointStyle) Then mConfig.SetSetting ConfigSettingPointStyle, mPointStyle
If IsPropertySet(DataPointPropertyDisplayMode) Then mConfig.SetSetting ConfigSettingDisplayMode, mDisplayMode
If IsPropertySet(DataPointPropertyHistWidth) Then mConfig.SetSetting ConfigSettingHistWidth, mHistWidth
If IsPropertySet(DataPointPropertyIncludeInAutoscale) Then mConfig.SetSetting ConfigSettingIncludeInAutoscale, mIncludeInAutoscale
If IsPropertySet(DataPointPropertyLayer) Then mConfig.SetSetting ConfigSettingLayer, mLayer

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Property

Public Property Let DisplayMode(ByVal Value As DataPointDisplayModes)
Const ProcName As String = "DisplayMode"
Dim failpoint As String
On Error GoTo Err

setLongProperty mDisplayMode, DataPointPropertyDisplayMode, Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DisplayMode() As DataPointDisplayModes
Const ProcName As String = "DisplayMode"
On Error GoTo Err

checkPropertySet DataPointPropertyDisplayMode
DisplayMode = mDisplayMode

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let DownColor(ByVal Value As Long)
Const ProcName As String = "DownColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(Value) And Value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color or -1"

setLongProperty mDownColor, DataPointPropertyDownColor, Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DownColor() As Long
Const ProcName As String = "DownColor"
On Error GoTo Err

checkPropertySet DataPointPropertyDownColor
DownColor = mDownColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let HistogramBarWidth(ByVal Value As Single)
Const ProcName As String = "HistogramBarWidth"
Dim failpoint As String
On Error GoTo Err

If Value <= 0 Or Value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "HistogramBarWidth must be greater than zero but not greater than 1"

setSingleProperty mHistWidth, DataPointPropertyHistWidth, Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get HistogramBarWidth() As Single
Const ProcName As String = "HistogramBarWidth"
On Error GoTo Err

checkPropertySet DataPointPropertyHistWidth
HistogramBarWidth = mHistWidth

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mIncludeInAutoscale, DataPointPropertyIncludeInAutoscale, Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"
On Error GoTo Err

checkPropertySet DataPointPropertyIncludeInAutoscale
IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(Value) And Value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color or -1"

setLongProperty mLayer, BarPropertyLayer, Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

checkPropertySet BarPropertyLayer
Layer = mLayer

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let LineStyle(ByVal Value As LineStyles)
Const ProcName As String = "LineStyle"
Dim failpoint As String
On Error GoTo Err

setLongProperty mLineStyle, DataPointPropertyLineStyle, Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LineStyle() As LineStyles
Const ProcName As String = "LineStyle"
On Error GoTo Err

checkPropertySet DataPointPropertyLineStyle
LineStyle = mLineStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let LineThickness(ByVal Value As Long)
Const ProcName As String = "LineThickness"
Dim failpoint As String
On Error GoTo Err

setLongProperty mLineThickness, DataPointPropertyLineThickness, Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LineThickness() As Long
Const ProcName As String = "LineThickness"
On Error GoTo Err

checkPropertySet DataPointPropertyLineThickness
LineThickness = mLineThickness

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let PointStyle(ByVal Value As PointStyles)
Const ProcName As String = "PointStyle"
Dim failpoint As String
On Error GoTo Err

setLongProperty mPointStyle, DataPointPropertyPointStyle, Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PointStyle() As PointStyles
Const ProcName As String = "PointStyle"
On Error GoTo Err

checkPropertySet DataPointPropertyPointStyle
PointStyle = mPointStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let UpColor(ByVal Value As Long)
Const ProcName As String = "UpColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(Value) And Value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color or -1"

setLongProperty mUpColor, DataPointPropertyUpColor, Value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get UpColor() As Long
Const ProcName As String = "UpColor"
On Error GoTo Err

checkPropertySet DataPointPropertyUpColor
UpColor = mUpColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ClearProperty( _
                ByVal pPropFlag As DataPointPropertyFlags)
Const ProcName As String = "ClearProperty"
Dim failpoint As String
On Error GoTo Err

mPropertyFlags = gClearFlag(mPropertyFlags, pPropFlag)
If Not mConfig Is Nothing Then mConfig.RemoveSetting propertyFlagToConfigSetting(pPropFlag)
firePropertyChanged gDataPointPropertyFlagToString(pPropFlag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function clone() As DataPointStyle
Const ProcName As String = "clone"
Dim failpoint As String
On Error GoTo Err

Set clone = New DataPointStyle
clone.Color = mColor
clone.DisplayMode = mDisplayMode
clone.DownColor = mDownColor
clone.HistogramBarWidth = mHistWidth
clone.IncludeInAutoscale = mIncludeInAutoscale
clone.LineStyle = mLineStyle
clone.LineThickness = mLineThickness
clone.PointStyle = PointStyle
clone.UpColor = mUpColor

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function IsPropertySet( _
                ByVal pPropertyFlag As DataPointPropertyFlags) As Boolean
                
IsPropertySet = gIsFlagSet(mPropertyFlags, pPropertyFlag)

End Function

Public Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)

Const ProcName As String = "LoadFromConfig"
Dim failpoint As String
On Error GoTo Err

Set mConfig = pConfig
mName = mConfig.GetSetting(ConfigSettingName)
mStyleType = mConfig.GetSetting(ConfigSettingStyleType)

If mConfig.GetSetting(ConfigSettingColor) <> "" Then Color = mConfig.GetSetting(ConfigSettingColor)

If mConfig.GetSetting(ConfigSettingDisplayMode) <> "" Then DisplayMode = mConfig.GetSetting(ConfigSettingDisplayMode)
If mConfig.GetSetting(ConfigSettingDownColor) <> "" Then DownColor = mConfig.GetSetting(ConfigSettingDownColor)

If mConfig.GetSetting(ConfigSettingHistWidth) <> "" Then HistogramBarWidth = mConfig.GetSetting(ConfigSettingHistWidth)

If mConfig.GetSetting(ConfigSettingIncludeInAutoscale) <> "" Then IncludeInAutoscale = mConfig.GetSetting(ConfigSettingIncludeInAutoscale)

If mConfig.GetSetting(ConfigSettingLayer) <> "" Then Layer = mConfig.GetSetting(ConfigSettingLayer)

If mConfig.GetSetting(ConfigSettingLineStyle) <> "" Then LineStyle = mConfig.GetSetting(ConfigSettingLineStyle)
If mConfig.GetSetting(ConfigSettingLineThickness) <> "" Then LineThickness = mConfig.GetSetting(ConfigSettingLineThickness)

If mConfig.GetSetting(ConfigSettingPointStyle) <> "" Then PointStyle = mConfig.GetSetting(ConfigSettingPointStyle)

If mConfig.GetSetting(ConfigSettingUpColor) <> "" Then UpColor = mConfig.GetSetting(ConfigSettingUpColor)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
                
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
Dim failpoint As String
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkPropertySet( _
                ByVal pPropFlag As DataPointPropertyFlags)
Const ProcName As String = "checkPropertySet"
On Error GoTo Err

If Not IsPropertySet(pPropFlag) Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Property is not set"
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub firePropertyChanged( _
                ByVal Name As String)
Dim ev As PropertyChangedEvent

Const ProcName As String = "firePropertyChanged"
Dim failpoint As String
On Error GoTo Err

Set ev.Source = Me
ev.PropertyName = Name
RaiseEvent PropertyChanged(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function propertyFlagToConfigSetting( _
            ByVal pPropFlag As DataPointPropertyFlags) As String
Select Case pPropFlag
Case DataPointPropertyLineThickness
    propertyFlagToConfigSetting = ConfigSettingLineThickness
Case DataPointPropertyColor
    propertyFlagToConfigSetting = ConfigSettingColor
Case DataPointPropertyUpColor
    propertyFlagToConfigSetting = ConfigSettingUpColor
Case DataPointPropertyDownColor
    propertyFlagToConfigSetting = ConfigSettingDownColor
Case DataPointPropertyLineStyle
    propertyFlagToConfigSetting = ConfigSettingLineStyle
Case DataPointPropertyPointStyle
    propertyFlagToConfigSetting = ConfigSettingPointStyle
Case DataPointPropertyDisplayMode
    propertyFlagToConfigSetting = ConfigSettingDisplayMode
Case DataPointPropertyHistWidth
    propertyFlagToConfigSetting = ConfigSettingHistWidth
Case DataPointPropertyIncludeInAutoscale
    propertyFlagToConfigSetting = ConfigSettingIncludeInAutoscale
End Select
End Function

Private Sub setBooleanProperty( _
                ByRef pProp As Boolean, _
                ByVal pPropFlag As DataPointPropertyFlags, _
                ByVal Value As Boolean)
Const ProcName As String = "setBooleanProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
If Not mConfig Is Nothing Then mConfig.SetSetting propertyFlagToConfigSetting(pPropFlag), pProp
firePropertyChanged gDataPointPropertyFlagToString(pPropFlag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setLongProperty( _
                ByRef pProp As Long, _
                ByVal pPropFlag As DataPointPropertyFlags, _
                ByVal Value As Long)
Const ProcName As String = "setLongProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
If Not mConfig Is Nothing Then mConfig.SetSetting propertyFlagToConfigSetting(pPropFlag), pProp
firePropertyChanged gDataPointPropertyFlagToString(pPropFlag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setSingleProperty( _
                ByRef pProp As Single, _
                ByVal pPropFlag As DataPointPropertyFlags, _
                ByVal Value As Single)
Const ProcName As String = "setSingleProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
If Not mConfig Is Nothing Then mConfig.SetSetting propertyFlagToConfigSetting(pPropFlag), pProp
firePropertyChanged gDataPointPropertyFlagToString(pPropFlag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setPropertyFlag( _
                ByVal flag As DataPointPropertyFlags)
Const ProcName As String = "setPropertyFlag"
Dim failpoint As String
On Error GoTo Err

mPropertyFlags = gSetFlag(mPropertyFlags, flag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub







