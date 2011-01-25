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

Private mEPhost                                         As ExtendedPropertyHost

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
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Autoscrolling(ByVal Value As Boolean)
Const ProcName As String = "Autoscrolling"
On Error GoTo Err

setProperty GChart.gAutoscrollingProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingAutoscrolling, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Autoscrolling() As Boolean
Const ProcName As String = "Autoscrolling"
On Error GoTo Err

Autoscrolling = mEPhost.GetLocalValue(GChart.gAutoscrollingProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BasedOn() As ChartStyle
Set BasedOn = mBasedOn
End Property

Public Property Let ChartBackColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "ChartBackColor"
On Error GoTo Err

setProperty GChart.gChartBackColorProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingChartBackColor, gLongToHexString(Value)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ChartBackColor() As OLE_COLOR
Const ProcName As String = "ChartBackColor"
On Error GoTo Err

ChartBackColor = mEPhost.GetLocalValue(GChart.gChartBackColorProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

If isLocalValueSet(GChart.gTwipsPerPeriodProperty) Then mConfig.SetSetting ConfigSettingTwipsPerPeriod, mEPhost.GetLocalValue(GChart.gTwipsPerPeriodProperty)

If isLocalValueSet(GChart.gAutoscrollingProperty) Then mConfig.SetSetting ConfigSettingAutoscrolling, mEPhost.GetLocalValue(GChart.gAutoscrollingProperty)
If isLocalValueSet(GChart.gChartBackColorProperty) Then mConfig.SetSetting ConfigSettingChartBackColor, mEPhost.GetLocalValue(GChart.gChartBackColorProperty)
If isLocalValueSet(GChart.gHorizontalMouseScrollingAllowedProperty) Then mConfig.SetSetting ConfigSettingHorizontalMouseScrollingAllowed, mEPhost.GetLocalValue(GChart.gHorizontalMouseScrollingAllowedProperty)
If isLocalValueSet(GChart.gHorizontalScrollBarVisibleProperty) Then mConfig.SetSetting ConfigSettingHorizontalScrollBarVisible, mEPhost.GetLocalValue(GChart.gHorizontalScrollBarVisibleProperty)
If isLocalValueSet(GChart.gVerticalMouseScrollingAllowedProperty) Then mConfig.SetSetting ConfigSettingVerticalMouseScrollingAllowed, mEPhost.GetLocalValue(GChart.gVerticalMouseScrollingAllowedProperty)
If isLocalValueSet(GChart.gXAxisVisibleProperty) Then mConfig.SetSetting ConfigSettingXAxisVisible, mEPhost.GetLocalValue(GChart.gXAxisVisibleProperty)
If isLocalValueSet(GChart.gYAxisVisibleProperty) Then mConfig.SetSetting ConfigSettingYAxisVisible, mEPhost.GetLocalValue(GChart.gYAxisVisibleProperty)
If isLocalValueSet(GChart.gYAxisWidthCmProperty) Then mConfig.SetSetting ConfigSettingYAxisWidthCm, mEPhost.GetLocalValue(GChart.gYAxisWidthCmProperty)

If isLocalValueSet(GChart.gCrosshairLineStyleProperty) Then mEPhost.GetLocalValue(GChart.gCrosshairLineStyleProperty).ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionCrosshairLineStyle)
If isLocalValueSet(GChart.gDefaultRegionStyleProperty) Then mEPhost.GetLocalValue(GChart.gDefaultRegionStyleProperty).ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultRegionStyle)
If isLocalValueSet(GChart.gDefaultYAxisRegionStyleProperty) Then mEPhost.GetLocalValue(GChart.gDefaultYAxisRegionStyleProperty).ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultYAxisRegionStyle)
If isLocalValueSet(GChart.gXAxisRegionStyleProperty) Then mEPhost.GetLocalValue(GChart.gXAxisRegionStyleProperty).ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXAxisRegionStyle)
If isLocalValueSet(GChart.gXCursorTextStyleProperty) Then mEPhost.GetLocalValue(GChart.gXCursorTextStyleProperty).ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXCursorTextStyle)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let CrosshairLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "CrosshairLineStyle"

On Error GoTo Err

Dim prevValue As LineStyle
If setProperty(GChart.gCrosshairLineStyleProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionCrosshairLineStyle)
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CrosshairLineStyle() As LineStyle
Const ProcName As String = "CrosshairLineStyle"
On Error GoTo Err

Set CrosshairLineStyle = mEPhost.GetLocalValue(GChart.gCrosshairLineStyleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let DefaultRegionStyle(ByVal Value As ChartRegionStyle)
Const ProcName As String = "DefaultRegionStyle"

On Error GoTo Err

Dim prevValue As ChartRegionStyle
If setProperty(GChart.gDefaultRegionStyleProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultRegionStyle)
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DefaultRegionStyle() As ChartRegionStyle
Const ProcName As String = "DefaultRegionStyle"
On Error GoTo Err

Set DefaultRegionStyle = mEPhost.GetLocalValue(GChart.gDefaultRegionStyleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let DefaultYAxisRegionStyle(ByVal Value As ChartRegionStyle)
Const ProcName As String = "DefaultYAxisRegionStyle"

On Error GoTo Err

Dim prevValue As ChartRegionStyle
If setProperty(GChart.gDefaultYAxisRegionStyleProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultYAxisRegionStyle)
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DefaultYAxisRegionStyle() As ChartRegionStyle
Const ProcName As String = "DefaultYAxisRegionStyle"
On Error GoTo Err

Set DefaultYAxisRegionStyle = mEPhost.GetLocalValue(GChart.gDefaultYAxisRegionStyleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Get ExtendedPropertyHost() As ExtendedPropertyHost
Set ExtendedPropertyHost = mEPhost
End Property

Public Property Let HorizontalMouseScrollingAllowed(ByVal Value As Boolean)
Const ProcName As String = "HorizontalMouseScrollingAllowed"
On Error GoTo Err

setProperty GChart.gHorizontalMouseScrollingAllowedProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingHorizontalMouseScrollingAllowed, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HorizontalMouseScrollingAllowed() As Boolean
Const ProcName As String = "HorizontalMouseScrollingAllowed"
On Error GoTo Err

HorizontalMouseScrollingAllowed = mEPhost.GetLocalValue(GChart.gHorizontalMouseScrollingAllowedProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HorizontalScrollBarVisible(ByVal Value As Boolean)
Const ProcName As String = "HorizontalScrollBarVisible"
On Error GoTo Err

setProperty GChart.gHorizontalScrollBarVisibleProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingHorizontalScrollBarVisible, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HorizontalScrollBarVisible() As Boolean
Const ProcName As String = "HorizontalScrollBarVisible"
On Error GoTo Err

HorizontalScrollBarVisible = mEPhost.GetLocalValue(GChart.gHorizontalScrollBarVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let TwipsPerPeriod(ByVal Value As Long)
Const ProcName As String = "TwipsPerPeriod"
On Error GoTo Err

setProperty GChart.gTwipsPerPeriodProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingTwipsPerPeriod, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get TwipsPerPeriod() As Long
Const ProcName As String = "TwipsPerPeriod"
On Error GoTo Err

TwipsPerPeriod = mEPhost.GetLocalValue(GChart.gTwipsPerPeriodProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let VerticalMouseScrollingAllowed(ByVal Value As Boolean)
Const ProcName As String = "VerticalMouseScrollingAllowed"
On Error GoTo Err

setProperty GChart.gVerticalMouseScrollingAllowedProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingVerticalMouseScrollingAllowed, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get VerticalMouseScrollingAllowed() As Boolean
Const ProcName As String = "VerticalMouseScrollingAllowed"
On Error GoTo Err

VerticalMouseScrollingAllowed = mEPhost.GetLocalValue(GChart.gVerticalMouseScrollingAllowedProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let XAxisVisible(ByVal Value As Boolean)
Const ProcName As String = "XAxisVisible"
On Error GoTo Err

setProperty GChart.gXAxisVisibleProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingXAxisVisible, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get XAxisVisible() As Boolean
Const ProcName As String = "XAxisVisible"
On Error GoTo Err

XAxisVisible = mEPhost.GetLocalValue(GChart.gXAxisVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let XAxisRegionStyle(ByVal Value As ChartRegionStyle)
Const ProcName As String = "XAxisRegionStyle"

On Error GoTo Err

Dim prevValue As ChartRegionStyle
If setProperty(GChart.gXAxisRegionStyleProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXAxisRegionStyle)
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get XAxisRegionStyle() As ChartRegionStyle
Const ProcName As String = "XAxisRegionStyle"
On Error GoTo Err

Set XAxisRegionStyle = mEPhost.GetLocalValue(GChart.gXAxisRegionStyleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let XCursorTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "XCursorTextStyle"

On Error GoTo Err

Dim prevValue As TextStyle
If setProperty(GChart.gXCursorTextStyleProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXCursorTextStyle)
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get XCursorTextStyle() As TextStyle
Const ProcName As String = "XCursorTextStyle"
On Error GoTo Err

Set XCursorTextStyle = mEPhost.GetLocalValue(GChart.gXCursorTextStyleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let YAxisVisible(ByVal Value As Boolean)
Const ProcName As String = "YAxisVisible"
On Error GoTo Err

setProperty GChart.gYAxisVisibleProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingYAxisVisible, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YAxisVisible() As Boolean
Const ProcName As String = "YAxisVisible"
On Error GoTo Err

YAxisVisible = mEPhost.GetLocalValue(GChart.gYAxisVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let YAxisWidthCm(ByVal Value As Single)
Const ProcName As String = "YAxisWidthCm"
On Error GoTo Err

setProperty GChart.gYAxisWidthCmProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingYAxisWidthCm, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YAxisWidthCm() As Single

Const ProcName As String = "YAxisWidthCm"
On Error GoTo Err

YAxisWidthCm = mEPhost.GetLocalValue(GChart.gYAxisWidthCmProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
                ByVal pCrosshairLineStyle As LineStyle, _
                ByVal pXCursorTextStyle As TextStyle)
Const ProcName As String = "Initialise"
On Error GoTo Err

mName = pName
BasedOn = pBasedOn

If Not pDefaultRegionStyle Is Nothing Then DefaultRegionStyle = pDefaultRegionStyle

If Not pXAxisRegionStyle Is Nothing Then XAxisRegionStyle = pXAxisRegionStyle

If Not pDefaultYAxisRegionStyle Is Nothing Then DefaultYAxisRegionStyle = pDefaultYAxisRegionStyle

If Not pCrosshairLineStyle Is Nothing Then CrosshairLineStyle = pCrosshairLineStyle

If Not pXCursorTextStyle Is Nothing Then XCursorTextStyle = pXCursorTextStyle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = config
If mConfig Is Nothing Then Exit Sub

mName = mConfig.InstanceQualifier

If mConfig.GetSetting(ConfigSettingBasedOn) <> "" Then BasedOn = gChartStylesManager(mConfig.GetSetting(ConfigSettingBasedOn))

If mConfig.GetSetting(ConfigSettingAutoscrolling) <> "" Then Autoscrolling = mConfig.GetSetting(ConfigSettingAutoscrolling, "True")
If mConfig.GetSetting(ConfigSettingChartBackColor) <> "" Then ChartBackColor = mConfig.GetSetting(ConfigSettingChartBackColor, CStr(vbWhite))
If mConfig.GetSetting(ConfigSettingHorizontalMouseScrollingAllowed) <> "" Then HorizontalMouseScrollingAllowed = mConfig.GetSetting(ConfigSettingHorizontalMouseScrollingAllowed, "true")
If mConfig.GetSetting(ConfigSettingHorizontalScrollBarVisible) <> "" Then HorizontalScrollBarVisible = mConfig.GetSetting(ConfigSettingHorizontalScrollBarVisible, "true")
If mConfig.GetSetting(ConfigSettingTwipsPerPeriod) <> "" Then TwipsPerPeriod = mConfig.GetSetting(ConfigSettingTwipsPerPeriod, "120")
If mConfig.GetSetting(ConfigSettingVerticalMouseScrollingAllowed) <> "" Then VerticalMouseScrollingAllowed = mConfig.GetSetting(ConfigSettingVerticalMouseScrollingAllowed, "true")
If mConfig.GetSetting(ConfigSettingXAxisVisible) <> "" Then XAxisVisible = mConfig.GetSetting(ConfigSettingXAxisVisible, "true")
If mConfig.GetSetting(ConfigSettingYAxisVisible) <> "" Then YAxisVisible = mConfig.GetSetting(ConfigSettingYAxisVisible, "true")
If mConfig.GetSetting(ConfigSettingYAxisWidthCm) <> "" Then YAxisWidthCm = mConfig.GetSetting(ConfigSettingYAxisWidthCm, "1.8")

Dim ls As LineStyle
If Not mConfig.GetConfigurationSection(ConfigSectionCrosshairLineStyle) Is Nothing Then
    Set ls = New LineStyle
    ls.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionCrosshairLineStyle)
    CrosshairLineStyle = ls
End If

Dim crs As ChartRegionStyle
If Not mConfig.GetConfigurationSection(ConfigSectionDefaultRegionStyle) Is Nothing Then
    Set crs = New ChartRegionStyle
    crs.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionDefaultRegionStyle)
    DefaultRegionStyle = crs
End If

If Not mConfig.GetConfigurationSection(ConfigSectionDefaultYAxisRegionStyle) Is Nothing Then
    Set crs = New ChartRegionStyle
    crs.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionDefaultYAxisRegionStyle)
    DefaultYAxisRegionStyle = crs
End If

If Not mConfig.GetConfigurationSection(ConfigSectionXAxisRegionStyle) Is Nothing Then
    Set crs = New ChartRegionStyle
    crs.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionXAxisRegionStyle)
    XAxisRegionStyle = crs
End If

Dim ts As TextStyle
If Not mConfig.GetConfigurationSection(ConfigSectionXCursorTextStyle) Is Nothing Then
    Set ts = New TextStyle
    ts.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionXCursorTextStyle)
    XCursorTextStyle = ts
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Private Function isLocalValueSet(ByVal pExtProp As ExtendedProperty) As Boolean
Const ProcName As String = "isLocalValueSet"
On Error GoTo Err

isLocalValueSet = mEPhost.IsPropertySet(pExtProp)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function setProperty( _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pNewValue As Variant, _
                Optional ByRef pPrevValue As Variant) As Boolean
Const ProcName As String = "setProperty"
On Error GoTo Err

setProperty = gSetProperty(mEPhost, pExtProp, pNewValue, pPrevValue)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function


