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

Implements IConfigurable
Implements IGraphicObjectStyle

'@================================================================================
' Events
'@================================================================================

Event PropertyChanged( _
                ByRef ev As PropertyChangedEventData)
                
'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "BarStyle"

Private Const ConfigSettingTailThickness         As String = "&TailThickness"
Private Const ConfigSettingOutlineThickness      As String = "&OutlineThickness"
Private Const ConfigSettingColor                 As String = "&Color"
Private Const ConfigSettingUpColor               As String = "&UpColor"
Private Const ConfigSettingDownColor             As String = "&DownColor"
Private Const ConfigSettingDisplayMode           As String = "&DisplayMode"
Private Const ConfigSettingSolidUpBody           As String = "&SolidUpBody"
Private Const ConfigSettingThickness             As String = "&Thickness"
Private Const ConfigSettingWidth                 As String = "&Width"
Private Const ConfigSettingIncludeInAutoscale    As String = "&IncludeInAutoscale"
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

Private mLayer As Long

Private mPropertyFlags As BarPropertyFlags

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mName = GenerateGUIDString
mStyleType = "Bar Styles"
End Sub

'@================================================================================
' IConfigurable Interface Members
'@================================================================================

Private Property Let IConfigurable_ConfigurationSection(ByVal RHS As ConfigurationSection)
Const ProcName As String = "IConfigurable_ConfigurationSection"
On Error GoTo Err

ConfigurationSection = RHS

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IConfigurable_LoadFromConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "IConfigurable_LoadFromConfig"
On Error GoTo Err

LoadFromConfig pConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IConfigurable_RemoveFromConfig()
Const ProcName As String = "IConfigurable_RemoveFromConfig"
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
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Color(ByVal Value As Long)
Const ProcName As String = "Color"

On Error GoTo Err

If Not gIsValidColor(Value) And Value <> -1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Color or -1"

setLongProperty mColor, BarPropertyColor, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Color() As Long
Const ProcName As String = "Color"
On Error GoTo Err

checkPropertySet BarPropertyColor
Color = mColor

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

If IsPropertySet(BarPropertyColor) Then mConfig.SetSetting ConfigSettingColor, mColor

If IsPropertySet(BarPropertyDisplayMode) Then mConfig.SetSetting ConfigSettingDisplayMode, mDisplayMode
If IsPropertySet(BarPropertyDownColor) Then mConfig.SetSetting ConfigSettingDownColor, mDownColor

If IsPropertySet(BarPropertyIncludeInAutoscale) Then mConfig.SetSetting ConfigSettingIncludeInAutoscale, mIncludeInAutoscale

If IsPropertySet(BarPropertyLayer) Then mConfig.SetSetting ConfigSettingLayer, mLayer

If IsPropertySet(BarPropertyOutlineThickness) Then mConfig.SetSetting ConfigSettingOutlineThickness, mOutlineThickness

If IsPropertySet(BarPropertySolidUpBody) Then mConfig.SetSetting ConfigSettingSolidUpBody, mSolidUpBody

If IsPropertySet(BarPropertyTailThickness) Then mConfig.SetSetting ConfigSettingTailThickness, mTailThickness
If IsPropertySet(BarPropertyThickness) Then mConfig.SetSetting ConfigSettingThickness, mThickness

If IsPropertySet(BarPropertyUpColor) Then mConfig.SetSetting ConfigSettingUpColor, mUpColor

If IsPropertySet(BarPropertyWidth) Then mConfig.SetSetting ConfigSettingWidth, mWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Let DisplayMode(ByVal Value As BarDisplayModes)
Const ProcName As String = "DisplayMode"

On Error GoTo Err

setLongProperty mDisplayMode, BarPropertyDisplayMode, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DisplayMode() As BarDisplayModes
Const ProcName As String = "DisplayMode"
On Error GoTo Err

checkPropertySet BarPropertyDisplayMode
DisplayMode = mDisplayMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DownColor(ByVal Value As Long)
Const ProcName As String = "DownColor"

On Error GoTo Err

If Not gIsValidColor(Value) And Value <> -1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Color or -1"

setLongProperty mDownColor, BarPropertyDownColor, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DownColor() As Long
Const ProcName As String = "DownColor"
On Error GoTo Err

checkPropertySet BarPropertyDownColor
DownColor = mDownColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"

On Error GoTo Err

setBooleanProperty mIncludeInAutoscale, BarPropertyIncludeInAutoscale, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"
On Error GoTo Err

checkPropertySet BarPropertyIncludeInAutoscale
IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"

On Error GoTo Err

If Value < LayerMin Or Value > LayerMax Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid layer number"

setLongProperty mLayer, BarPropertyLayer, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

checkPropertySet BarPropertyLayer
layer = mLayer

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

Public Property Let OutlineThickness(ByVal Value As Long)
Const ProcName As String = "OutlineThickness"

On Error GoTo Err

setLongProperty mOutlineThickness, BarPropertyOutlineThickness, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OutlineThickness() As Long
Const ProcName As String = "OutlineThickness"
On Error GoTo Err

checkPropertySet BarPropertyOutlineThickness
OutlineThickness = mOutlineThickness

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let SolidUpBody(ByVal Value As Boolean)
Const ProcName As String = "SolidUpBody"

On Error GoTo Err

setBooleanProperty mSolidUpBody, BarPropertySolidUpBody, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SolidUpBody() As Boolean
Const ProcName As String = "SolidUpBody"
On Error GoTo Err

checkPropertySet BarPropertySolidUpBody
SolidUpBody = mSolidUpBody

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TailThickness(ByVal Value As Long)
Const ProcName As String = "TailThickness"

On Error GoTo Err

setLongProperty mTailThickness, BarPropertyTailThickness, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TailThickness() As Long
Const ProcName As String = "TailThickness"
On Error GoTo Err

checkPropertySet BarPropertyTailThickness
TailThickness = mTailThickness

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Thickness(ByVal Value As Long)
Const ProcName As String = "Thickness"

On Error GoTo Err

setLongProperty mThickness, BarPropertyThickness, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Thickness() As Long
Const ProcName As String = "Thickness"
On Error GoTo Err

checkPropertySet BarPropertyThickness
Thickness = mThickness

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let UpColor(ByVal Value As Long)
Const ProcName As String = "UpColor"

On Error GoTo Err

If Not gIsValidColor(Value) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Color"

setLongProperty mUpColor, BarPropertyUpColor, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get UpColor() As Long
Const ProcName As String = "UpColor"
On Error GoTo Err

checkPropertySet BarPropertyUpColor
UpColor = mUpColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Width(ByVal Value As Single)
Const ProcName As String = "Width"

On Error GoTo Err

setSingleProperty mWidth, BarPropertyWidth, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Width() As Single
Const ProcName As String = "Width"
On Error GoTo Err

checkPropertySet BarPropertyWidth
Width = mWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ClearProperty( _
                ByVal pPropFlag As BarPropertyFlags)
Const ProcName As String = "ClearProperty"

On Error GoTo Err

mPropertyFlags = gClearFlag(mPropertyFlags, pPropFlag)
If Not mConfig Is Nothing Then mConfig.RemoveSetting propertyFlagToConfigSetting(pPropFlag)
firePropertyChanged gBarPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function clone() As BarStyle
Const ProcName As String = "Clone"

On Error GoTo Err

Set clone = New BarStyle
If IsPropertySet(BarPropertyColor) Then clone.Color = mColor
If IsPropertySet(BarPropertyDisplayMode) Then clone.DisplayMode = mDisplayMode
If IsPropertySet(BarPropertyDownColor) Then clone.DownColor = mDownColor
If IsPropertySet(BarPropertyIncludeInAutoscale) Then clone.IncludeInAutoscale = mIncludeInAutoscale
If IsPropertySet(BarPropertyLayer) Then clone.layer = mLayer
If IsPropertySet(BarPropertyOutlineThickness) Then clone.OutlineThickness = mOutlineThickness
If IsPropertySet(BarPropertySolidUpBody) Then clone.SolidUpBody = mSolidUpBody
If IsPropertySet(BarPropertyTailThickness) Then clone.TailThickness = mTailThickness
If IsPropertySet(BarPropertyThickness) Then clone.Thickness = mThickness
If IsPropertySet(BarPropertyUpColor) Then clone.UpColor = mUpColor
If IsPropertySet(BarPropertyWidth) Then clone.Width = mWidth

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsPropertySet( _
                ByVal pPropertyFlag As BarPropertyFlags) As Boolean
                
IsPropertySet = gIsFlagSet(mPropertyFlags, pPropertyFlag)

End Function

Public Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"

On Error GoTo Err

Set mConfig = pConfig
mName = mConfig.GetSetting(ConfigSettingName)
mStyleType = mConfig.GetSetting(ConfigSettingStyleType)

If mConfig.GetSetting(ConfigSettingColor) <> "" Then Color = mConfig.GetSetting(ConfigSettingColor)

If mConfig.GetSetting(ConfigSettingDisplayMode) <> "" Then DisplayMode = mConfig.GetSetting(ConfigSettingDisplayMode)
If mConfig.GetSetting(ConfigSettingDownColor) <> "" Then DownColor = mConfig.GetSetting(ConfigSettingDownColor)

If mConfig.GetSetting(ConfigSettingIncludeInAutoscale) <> "" Then IncludeInAutoscale = mConfig.GetSetting(ConfigSettingIncludeInAutoscale)

If mConfig.GetSetting(ConfigSettingLayer) <> "" Then layer = mConfig.GetSetting(ConfigSettingLayer)

If mConfig.GetSetting(ConfigSettingOutlineThickness) <> "" Then OutlineThickness = mConfig.GetSetting(ConfigSettingOutlineThickness)

If mConfig.GetSetting(ConfigSettingSolidUpBody) <> "" Then SolidUpBody = mConfig.GetSetting(ConfigSettingSolidUpBody)

If mConfig.GetSetting(ConfigSettingThickness) <> "" Then Thickness = mConfig.GetSetting(ConfigSettingThickness)
If mConfig.GetSetting(ConfigSettingTailThickness) <> "" Then TailThickness = mConfig.GetSetting(ConfigSettingTailThickness)

If mConfig.GetSetting(ConfigSettingUpColor) <> "" Then UpColor = mConfig.GetSetting(ConfigSettingUpColor)
If mConfig.GetSetting(ConfigSettingWidth) <> "" Then Width = mConfig.GetSetting(ConfigSettingWidth)

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

Private Sub checkPropertySet( _
                ByVal pPropFlag As BarPropertyFlags)
Const ProcName As String = "checkPropertySet"
On Error GoTo Err

If Not IsPropertySet(pPropFlag) Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Property is not set"

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

Private Function propertyFlagToConfigSetting( _
            ByVal pPropFlag As BarPropertyFlags) As String
Const ProcName As String = "propertyFlagToConfigSetting"
On Error GoTo Err

Select Case pPropFlag
Case BarPropertyColor
    propertyFlagToConfigSetting = ConfigSettingColor
Case BarPropertyUpColor
    propertyFlagToConfigSetting = ConfigSettingUpColor
Case BarPropertyDownColor
    propertyFlagToConfigSetting = ConfigSettingDownColor
Case BarPropertyDisplayMode
    propertyFlagToConfigSetting = ConfigSettingDisplayMode
Case BarPropertySolidUpBody
    propertyFlagToConfigSetting = ConfigSettingSolidUpBody
Case BarPropertyThickness
    propertyFlagToConfigSetting = ConfigSettingThickness
Case BarPropertyWidth
    propertyFlagToConfigSetting = ConfigSettingWidth
Case BarPropertyTailThickness
    propertyFlagToConfigSetting = ConfigSettingTailThickness
Case BarPropertyOutlineThickness
    propertyFlagToConfigSetting = ConfigSettingOutlineThickness
Case BarPropertyIncludeInAutoscale
    propertyFlagToConfigSetting = ConfigSettingIncludeInAutoscale
Case BarPropertyLayer
    propertyFlagToConfigSetting = ConfigSettingLayer
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid property flag value"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setBooleanProperty( _
                ByRef pProp As Boolean, _
                ByVal pPropFlag As BarPropertyFlags, _
                ByVal Value As Boolean)
Const ProcName As String = "setBooleanProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
If Not mConfig Is Nothing Then mConfig.SetSetting propertyFlagToConfigSetting(pPropFlag), pProp
firePropertyChanged gBarPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setLongProperty( _
                ByRef pProp As Long, _
                ByVal pPropFlag As BarPropertyFlags, _
                ByVal Value As Long)
Const ProcName As String = "setLongProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
If Not mConfig Is Nothing Then mConfig.SetSetting propertyFlagToConfigSetting(pPropFlag), pProp
firePropertyChanged gBarPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSingleProperty( _
                ByRef pProp As Single, _
                ByVal pPropFlag As BarPropertyFlags, _
                ByVal Value As Single)
Const ProcName As String = "setSingleProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
If Not mConfig Is Nothing Then mConfig.SetSetting propertyFlagToConfigSetting(pPropFlag), pProp
firePropertyChanged gBarPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setPropertyFlag( _
                ByVal flag As BarPropertyFlags)
Const ProcName As String = "setPropertyFlag"

On Error GoTo Err

mPropertyFlags = gSetFlag(mPropertyFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




