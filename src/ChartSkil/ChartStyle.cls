VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements IConfigurable

'@================================================================================
' Events
'@================================================================================

Event PropertyChanged( _
                ByRef ev As PropertyChangedEventData)
                
'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                                As String = "ChartStyle"

'@================================================================================
' Member variables
'@================================================================================

Private mName                                           As String

Private WithEvents mEPhost                              As ExtendedPropertyHost
Attribute mEPhost.VB_VarHelpID = -1

Private mBasedOn                                        As ChartStyle

Private mConfig                                         As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mEPhost = New ExtendedPropertyHost
GChart.gRegisterProperties
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
' mEPhost Event Handlers
'@================================================================================

Private Sub mEPhost_ExtendedPropertyChanged(pEv As ExtendedPropertyChangedEventData)
Dim extProp As ExtendedProperty
Const ProcName As String = "mEPhost_ExtendedPropertyChanged"
On Error GoTo Err

updateConfigSetting pEv.ExtendedProperty, pEv.OldValue

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let Autoscrolling(ByVal Value As Boolean)
Const ProcName As String = "Autoscrolling"
On Error GoTo Err

setProperty GChart.gAutoscrollingProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Autoscrolling() As Boolean
Const ProcName As String = "Autoscrolling"
On Error GoTo Err

Autoscrolling = mEPhost.getLocalValue(GChart.gAutoscrollingProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get AutoscrollingProperty() As ExtendedProperty
Set AutoscrollingProperty = gAutoscrollingProperty
End Property

Friend Property Let BasedOn(ByVal Value As ChartStyle)
Const ProcName As String = "BasedOn"
On Error GoTo Err

Set mBasedOn = Value
If mBasedOn Is Nothing Then
    ' only happens for the default style when it is first created
Else
    mEPhost.Style = mBasedOn.ExtendedPropertyHost
    If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBasedOn, Value.Name
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BasedOn() As ChartStyle
Set BasedOn = mBasedOn
End Property

Public Property Let ChartBackColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "ChartBackColor"
On Error GoTo Err

setProperty GChart.gChartBackColorProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ChartBackColor() As OLE_COLOR
Const ProcName As String = "ChartBackColor"
On Error GoTo Err

ChartBackColor = mEPhost.getLocalValue(GChart.gChartBackColorProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ChartBackColorProperty() As ExtendedProperty
Set ChartBackColorProperty = gChartBackColorProperty
End Property

Friend Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If Value Is Nothing Then
    RemoveFromConfig
    Set mConfig = Nothing
    Exit Property
End If

If Value Is mConfig Then Exit Property
Set mConfig = Value

If Not mBasedOn Is Nothing Then mConfig.SetSetting ConfigSettingBasedOn, mBasedOn.Name

setConfig GChart.gPeriodWidthProperty

setConfig GChart.gAutoscrollingProperty
setConfig GChart.gChartBackColorProperty
setConfig GChart.gHorizontalMouseScrollingAllowedProperty
setConfig GChart.gHorizontalScrollBarVisibleProperty
setConfig GChart.gVerticalMouseScrollingAllowedProperty
setConfig GChart.gXAxisVisibleProperty
setConfig GChart.gYAxisVisibleProperty
setConfig GChart.gYAxisWidthCmProperty

setConfig GChart.gCrosshairLineStyleProperty
setConfig GChart.gDefaultRegionStyleProperty
setConfig GChart.gDefaultYAxisRegionStyleProperty
setConfig GChart.gXAxisRegionStyleProperty

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let CrosshairLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "CrosshairLineStyle"

On Error GoTo Err

setProperty GChart.gCrosshairLineStyleProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CrosshairLineStyle() As LineStyle
Const ProcName As String = "CrosshairLineStyle"
On Error GoTo Err

Set CrosshairLineStyle = mEPhost.getLocalValue(GChart.gCrosshairLineStyleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CrosshairLineStyleProperty() As ExtendedProperty
Set CrosshairLineStyleProperty = gCrosshairLineStyleProperty
End Property

Friend Property Let DefaultRegionStyle(ByVal Value As ChartRegionStyle)
Const ProcName As String = "DefaultRegionStyle"

On Error GoTo Err

setProperty GChart.gDefaultRegionStyleProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultRegionStyle() As ChartRegionStyle
Const ProcName As String = "DefaultRegionStyle"
On Error GoTo Err

Set DefaultRegionStyle = mEPhost.getLocalValue(GChart.gDefaultRegionStyleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultRegionStyleProperty() As ExtendedProperty
Set DefaultRegionStyleProperty = gDefaultRegionStyleProperty
End Property

Friend Property Let DefaultYAxisRegionStyle(ByVal Value As ChartRegionStyle)
Const ProcName As String = "DefaultYAxisRegionStyle"
On Error GoTo Err

setProperty GChart.gDefaultYAxisRegionStyleProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultYAxisRegionStyle() As ChartRegionStyle
Const ProcName As String = "DefaultYAxisRegionStyle"
On Error GoTo Err

Set DefaultYAxisRegionStyle = mEPhost.getLocalValue(GChart.gDefaultYAxisRegionStyleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultYAxisRegionStyleProperty() As ExtendedProperty
Set DefaultYAxisRegionStyleProperty = gDefaultYAxisRegionStyleProperty
End Property

Public Property Get ExtendedPropertyHost() As ExtendedPropertyHost
Set ExtendedPropertyHost = mEPhost
End Property

Public Property Let HorizontalMouseScrollingAllowed(ByVal Value As Boolean)
Const ProcName As String = "HorizontalMouseScrollingAllowed"
On Error GoTo Err

setProperty GChart.gHorizontalMouseScrollingAllowedProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HorizontalMouseScrollingAllowed() As Boolean
Const ProcName As String = "HorizontalMouseScrollingAllowed"
On Error GoTo Err

HorizontalMouseScrollingAllowed = mEPhost.getLocalValue(GChart.gHorizontalMouseScrollingAllowedProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HorizontalMouseScrollingAllowedProperty() As ExtendedProperty
Set HorizontalMouseScrollingAllowedProperty = gHorizontalMouseScrollingAllowedProperty
End Property

Public Property Let HorizontalScrollBarVisible(ByVal Value As Boolean)
Const ProcName As String = "HorizontalScrollBarVisible"
On Error GoTo Err

setProperty GChart.gHorizontalScrollBarVisibleProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HorizontalScrollBarVisible() As Boolean
Const ProcName As String = "HorizontalScrollBarVisible"
On Error GoTo Err

HorizontalScrollBarVisible = mEPhost.getLocalValue(GChart.gHorizontalScrollBarVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HorizontalScrollBarVisibleProperty() As ExtendedProperty
Set HorizontalScrollBarVisibleProperty = gHorizontalScrollBarVisibleProperty
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let PeriodWidth(ByVal Value As Long)
Const ProcName As String = "PeriodWidth"
On Error GoTo Err

setProperty GChart.gPeriodWidthProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PeriodWidth() As Long
Const ProcName As String = "PeriodWidth"
On Error GoTo Err

PeriodWidth = mEPhost.getLocalValue(GChart.gPeriodWidthProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PeriodWidthProperty() As ExtendedProperty
Set PeriodWidthProperty = gPeriodWidthProperty
End Property

Public Property Let VerticalMouseScrollingAllowed(ByVal Value As Boolean)
Const ProcName As String = "VerticalMouseScrollingAllowed"
On Error GoTo Err

setProperty GChart.gVerticalMouseScrollingAllowedProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get VerticalMouseScrollingAllowed() As Boolean
Const ProcName As String = "VerticalMouseScrollingAllowed"
On Error GoTo Err

VerticalMouseScrollingAllowed = mEPhost.getLocalValue(GChart.gVerticalMouseScrollingAllowedProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get VerticalMouseScrollingAllowedProperty() As ExtendedProperty
Set VerticalMouseScrollingAllowedProperty = gVerticalMouseScrollingAllowedProperty
End Property

Public Property Let XAxisVisible(ByVal Value As Boolean)
Const ProcName As String = "XAxisVisible"
On Error GoTo Err

setProperty GChart.gXAxisVisibleProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XAxisVisible() As Boolean
Const ProcName As String = "XAxisVisible"
On Error GoTo Err

XAxisVisible = mEPhost.getLocalValue(GChart.gXAxisVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XAxisVisibleProperty() As ExtendedProperty
Set XAxisVisibleProperty = gXAxisVisibleProperty
End Property

Friend Property Let XAxisRegionStyle(ByVal Value As ChartRegionStyle)
Const ProcName As String = "XAxisRegionStyle"

On Error GoTo Err

setProperty GChart.gXAxisRegionStyleProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XAxisRegionStyle() As ChartRegionStyle
Const ProcName As String = "XAxisRegionStyle"
On Error GoTo Err

Set XAxisRegionStyle = mEPhost.getLocalValue(GChart.gXAxisRegionStyleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XAxisRegionStyleProperty() As ExtendedProperty
Set XAxisRegionStyleProperty = gXAxisRegionStyleProperty
End Property

Public Property Let YAxisVisible(ByVal Value As Boolean)
Const ProcName As String = "YAxisVisible"
On Error GoTo Err

setProperty GChart.gYAxisVisibleProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YAxisVisible() As Boolean
Const ProcName As String = "YAxisVisible"
On Error GoTo Err

YAxisVisible = mEPhost.getLocalValue(GChart.gYAxisVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YAxisVisibleProperty() As ExtendedProperty
Set YAxisVisibleProperty = gYAxisVisibleProperty
End Property

Public Property Let YAxisWidthCm(ByVal Value As Single)
Const ProcName As String = "YAxisWidthCm"
On Error GoTo Err

setProperty GChart.gYAxisWidthCmProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YAxisWidthCm() As Single

Const ProcName As String = "YAxisWidthCm"
On Error GoTo Err

YAxisWidthCm = mEPhost.getLocalValue(GChart.gYAxisWidthCmProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YAxisWidthCmProperty() As ExtendedProperty
Set YAxisWidthCmProperty = gYAxisWidthCmProperty
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pName As String, _
                ByVal pBasedOn As ChartStyle, _
                ByVal pDefaultRegionStyle As ChartRegionStyle, _
                ByVal pXAxisRegionStyle As ChartRegionStyle, _
                ByVal pDefaultYAxisRegionStyle As ChartRegionStyle, _
                ByVal pCrosshairLineStyle As LineStyle)
Const ProcName As String = "Initialise"
On Error GoTo Err

mName = pName
BasedOn = pBasedOn

If Not pDefaultRegionStyle Is Nothing Then DefaultRegionStyle = pDefaultRegionStyle

If Not pXAxisRegionStyle Is Nothing Then XAxisRegionStyle = pXAxisRegionStyle

If Not pDefaultYAxisRegionStyle Is Nothing Then DefaultYAxisRegionStyle = pDefaultYAxisRegionStyle

If Not pCrosshairLineStyle Is Nothing Then CrosshairLineStyle = pCrosshairLineStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = config
If mConfig Is Nothing Then Exit Sub

mName = mConfig.InstanceQualifier

If mConfig.GetSetting(ConfigSettingBasedOn) <> "" Then BasedOn = gChartStylesManager(mConfig.GetSetting(ConfigSettingBasedOn))

loadNonObjectFromConfig GChart.gAutoscrollingProperty, "True"
loadNonObjectFromConfig GChart.gChartBackColorProperty, CStr(vbWhite)
loadNonObjectFromConfig GChart.gHorizontalMouseScrollingAllowedProperty, "True"
loadNonObjectFromConfig GChart.gHorizontalScrollBarVisibleProperty, "True"
loadNonObjectFromConfig GChart.gPeriodWidthProperty, DefaultPeriodWidth
loadNonObjectFromConfig GChart.gVerticalMouseScrollingAllowedProperty, "True"
loadNonObjectFromConfig GChart.gXAxisVisibleProperty, "True"
loadNonObjectFromConfig GChart.gYAxisVisibleProperty, "True"
loadNonObjectFromConfig GChart.gYAxisWidthCmProperty, DefaultYAxisWidthCm

loadObjectFromConfig GChart.gCrosshairLineStyleProperty, New LineStyle
loadObjectFromConfig GChart.gDefaultRegionStyleProperty, New ChartRegionStyle
loadObjectFromConfig GChart.gDefaultYAxisRegionStyleProperty, New ChartRegionStyle
loadObjectFromConfig GChart.gXAxisRegionStyleProperty, New ChartRegionStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub firePropertyChanged( _
                ByVal Name As String)
Dim ev As PropertyChangedEventData

Set ev.Source = Me
ev.PropertyName = Name
RaiseEvent PropertyChanged(ev)
End Sub

Private Function getConfigName(ByVal pExtProp As ExtendedProperty) As String
If pExtProp.ValueIsObject Then
    getConfigName = pExtProp.PropertyName
Else
    getConfigName = "&" & pExtProp.PropertyName
End If
End Function

Private Function getConfigurableObject(ByVal pValue As Variant) As IConfigurable
Set getConfigurableObject = pValue
End Function

Private Function getLocalValue(ByVal pExtProp As ExtendedProperty) As Variant
gSetVariant getLocalValue, mEPhost.getLocalValue(pExtProp)
End Function

Private Function isLocalValueSet(ByVal pExtProp As ExtendedProperty) As Boolean
Const ProcName As String = "isLocalValueSet"
On Error GoTo Err

isLocalValueSet = mEPhost.IsPropertySet(pExtProp)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub loadNonObjectFromConfig( _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pDefault As String)
Dim lConfigName As String
Const ProcName As String = "loadNonObjectFromConfig"
On Error GoTo Err

lConfigName = getConfigName(pExtProp)
If mConfig.GetSetting(lConfigName, "") <> "" Then mEPhost.SetValue pExtProp, mConfig.GetSetting(lConfigName, pDefault)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub loadObjectFromConfig( _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pNewObject As IConfigurable)
Dim lConfig As ConfigurationSection
Const ProcName As String = "loadObjectFromConfig"
On Error GoTo Err

Set lConfig = mConfig.GetConfigurationSection(getConfigName(pExtProp))
If Not lConfig Is Nothing Then
    pNewObject.LoadFromConfig lConfig
    mEPhost.SetValue pExtProp, pNewObject
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeConfig( _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pOldvalue As Variant)

Const ProcName As String = "removeConfig"
On Error GoTo Err

If pExtProp.ValueIsObject Then
    removeConfigForObject pExtProp, pOldvalue
Else
    removeConfigForNonObject pExtProp
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeConfigForNonObject(ByVal pExtProp As ExtendedProperty)
mConfig.RemoveSetting getConfigName(pExtProp)
End Sub

Private Sub removeConfigForObject(ByVal pExtProp As ExtendedProperty, ByVal pOldvalue As IConfigurable)
mConfig.RemoveConfigurationSection getConfigName(pExtProp)
' Don't execute the following line because the old value may be from the parent style,
' and we don't want to remove that from the config
'If Not pOldvalue Is Nothing Then getConfigurableObject(pOldvalue).RemoveFromConfig
End Sub

Private Sub setConfig(ByVal pExtProp As ExtendedProperty)
If Not isLocalValueSet(pExtProp) Then Exit Sub
If pExtProp.ValueIsObject Then
    setConfigForObject pExtProp
Else
    setConfigForNonObject pExtProp
End If
End Sub

Private Sub setConfigForNonObject(ByVal pExtProp As ExtendedProperty)
mConfig.SetSetting getConfigName(pExtProp), getLocalValue(pExtProp)
End Sub
               
Private Sub setConfigForObject(ByVal pExtProp As ExtendedProperty)
getConfigurableObject(getLocalValue(pExtProp)).ConfigurationSection = mConfig.AddConfigurationSection(getConfigName(pExtProp))
End Sub
               
Private Sub setProperty( _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pNewValue As Variant)
Const ProcName As String = "setProperty"
On Error GoTo Err

gSetProperty mEPhost, pExtProp, pNewValue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateConfigSetting( _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pOldvalue As Variant)
Const ProcName As String = "updateConfigSetting"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

removeConfig pExtProp, pOldvalue
setConfig pExtProp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
