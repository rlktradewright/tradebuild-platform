VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "StudyValueConfiguration"

Private Const ConfigSectionBarStyle                             As String = "BarStyle"
Private Const ConfigSectionDataPointStyle                       As String = "DataPointStyle"
Private Const ConfigSectionLineStyle                            As String = "LineStyle"
Private Const ConfigSectionTextStyle                            As String = "TextStyle"

Private Const ConfigSettingBarFormatterFactoryName              As String = "&BarFormatterFactoryName"
Private Const ConfigSettingBarFormatterLibraryName              As String = "&BarFormatterLibraryName"
Private Const ConfigSettingChartRegionName                      As String = "&ChartRegionName"
Private Const ConfigSettingIncludeInChart                       As String = "&IncludeInChart"
Private Const ConfigSettingLayer                                As String = "&Layer"
Private Const ConfigSettingOffsetX                              As String = "&X"
Private Const ConfigSettingOffsetY                              As String = "&Y"
Private Const ConfigSettingValueName                            As String = "&ValueName"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mIncludeInChart As Boolean
Private mValueName As String

Private mBarStyle As BarStyle
Private mDataPointStyle As DataPointStyle
Private mLineStyle As LineStyle
Private mTextStyle As TextStyle

Private mChartRegionName As String
Private mLayer As Long
Private mOffsetX As Double
Private mOffsetY As Double

Private mBarFormatterFactoryName As String
Private mBarFormatterLibraryName As String

Private mConfig As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mLayer = LayerNone
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

Public Property Let BarFormatterFactoryName(ByVal Value As String)
Const ProcName As String = "BarFormatterFactoryName"
On Error GoTo Err

mBarFormatterFactoryName = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BarFormatterFactoryName() As String
Const ProcName As String = "BarFormatterFactoryName"
On Error GoTo Err

BarFormatterFactoryName = mBarFormatterFactoryName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BarFormatterLibraryName(ByVal Value As String)
Const ProcName As String = "BarFormatterLibraryName"
On Error GoTo Err

mBarFormatterLibraryName = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BarFormatterLibraryName() As String
Const ProcName As String = "BarFormatterLibraryName"
On Error GoTo Err

BarFormatterLibraryName = mBarFormatterLibraryName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BarStyle(ByVal Value As BarStyle)
Const ProcName As String = "BarStyle"
On Error GoTo Err

Set mBarStyle = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BarStyle() As BarStyle
Const ProcName As String = "BarStyle"
On Error GoTo Err

Set BarStyle = mBarStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ChartRegionName(ByVal Value As String)
Const ProcName As String = "ChartRegionName"
On Error GoTo Err

mChartRegionName = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ChartRegionName() As String
Const ProcName As String = "ChartRegionName"
On Error GoTo Err

ChartRegionName = mChartRegionName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If Value Is mConfig Then Exit Property
Set mConfig = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DataPointStyle(ByVal Value As DataPointStyle)
Const ProcName As String = "DataPointStyle"
On Error GoTo Err

Set mDataPointStyle = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DataPointStyle() As DataPointStyle
Const ProcName As String = "DataPointStyle"
On Error GoTo Err

Set DataPointStyle = mDataPointStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let IncludeInChart(ByVal Value As Boolean)
Const ProcName As String = "IncludeInChart"
On Error GoTo Err

mIncludeInChart = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IncludeInChart() As Boolean
Const ProcName As String = "IncludeInChart"
On Error GoTo Err

IncludeInChart = mIncludeInChart

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Layer(ByVal Value As Long)
Const ProcName As String = "Layer"
On Error GoTo Err

mLayer = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Layer() As Long
Const ProcName As String = "Layer"
On Error GoTo Err

Layer = mLayer

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let LineStyle(ByVal Value As LineStyle)
Const ProcName As String = "LineStyle"
On Error GoTo Err

Set mLineStyle = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LineStyle() As LineStyle
Const ProcName As String = "LineStyle"
On Error GoTo Err

Set LineStyle = mLineStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let OffsetX(ByVal Value As Double)
Const ProcName As String = "OffsetX"
On Error GoTo Err

mOffsetX = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OffsetX() As Double
Const ProcName As String = "OffsetX"
On Error GoTo Err

OffsetX = mOffsetX

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let OffsetY(ByVal Value As Double)
Const ProcName As String = "OffsetY"
On Error GoTo Err

mOffsetY = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OffsetY() As Double
Const ProcName As String = "OffsetY"
On Error GoTo Err

OffsetY = mOffsetY

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TextStyle(ByVal Value As TextStyle)
Const ProcName As String = "TextStyle"
On Error GoTo Err

Set mTextStyle = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextStyle() As TextStyle
Const ProcName As String = "TextStyle"
On Error GoTo Err

Set TextStyle = mTextStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let ValueName(ByVal Value As String)
Const ProcName As String = "ValueName"
On Error GoTo Err

mValueName = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ValueName() As String
Const ProcName As String = "ValueName"
On Error GoTo Err

ValueName = mValueName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = config

gLogger.Log "Loading StudyValueConfiguration: " & mConfig.Path, ProcName, ModuleName

Dim cs As ConfigurationSection
Set cs = mConfig.GetConfigurationSection(ConfigSectionBarStyle)
If Not cs Is Nothing Then
    Set mBarStyle = New BarStyle
    mBarStyle.LoadFromConfig cs
End If

Set cs = mConfig.GetConfigurationSection(ConfigSectionDataPointStyle)
If Not cs Is Nothing Then
    Set mDataPointStyle = New DataPointStyle
    mDataPointStyle.LoadFromConfig cs
End If

Set cs = mConfig.GetConfigurationSection(ConfigSectionLineStyle)
If Not cs Is Nothing Then
    Set mLineStyle = New LineStyle
    mLineStyle.LoadFromConfig cs
End If

Set cs = mConfig.GetConfigurationSection(ConfigSectionTextStyle)
If Not cs Is Nothing Then
    Set mTextStyle = New TextStyle
    mTextStyle.LoadFromConfig cs
End If

mBarFormatterFactoryName = mConfig.GetSetting(ConfigSettingBarFormatterFactoryName, "")
mBarFormatterLibraryName = mConfig.GetSetting(ConfigSettingBarFormatterLibraryName, "")
mIncludeInChart = mConfig.GetSetting(ConfigSettingIncludeInChart)
mValueName = mConfig.GetSetting(ConfigSettingValueName)
mChartRegionName = mConfig.GetSetting(ConfigSettingChartRegionName)
mLayer = mConfig.GetSetting(ConfigSettingLayer)
mOffsetX = mConfig.GetSetting(ConfigSettingOffsetX)
mOffsetY = mConfig.GetSetting(ConfigSettingOffsetY)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub storeSettings()
Const ProcName As String = "storeSettings"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

If Not mBarStyle Is Nothing Then mBarStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionBarStyle)
If Not mDataPointStyle Is Nothing Then mDataPointStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDataPointStyle)
If Not mLineStyle Is Nothing Then mLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionLineStyle)
If Not mTextStyle Is Nothing Then mTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionTextStyle)

mConfig.SetSetting ConfigSettingBarFormatterFactoryName, mBarFormatterFactoryName
mConfig.SetSetting ConfigSettingBarFormatterLibraryName, mBarFormatterLibraryName
mConfig.SetSetting ConfigSettingIncludeInChart, mIncludeInChart
mConfig.SetSetting ConfigSettingValueName, mValueName
mConfig.SetSetting ConfigSettingChartRegionName, mChartRegionName
mConfig.SetSetting ConfigSettingLayer, mLayer
mConfig.SetSetting ConfigSettingOffsetX, mOffsetX
mConfig.SetSetting ConfigSettingOffsetY, mOffsetY

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


