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

Private Const ConfigSectionBarFormatterFactory                  As String = "BarFormatterFactory"
Private Const ConfigSectionBarStyle                             As String = "BarStyle"
Private Const ConfigSectionDataPointStyle                       As String = "DataPointStyle"
Private Const ConfigSectionLineStyle                            As String = "LineStyle"
Private Const ConfigSectionTextStyle                            As String = "TextStyle"

Private Const ConfigSettingChartRegionName                      As String = "&ChartRegionName"
Private Const ConfigSettingIncludeInChart                       As String = "&IncludeInChart"
Private Const ConfigSettingLayer                                As String = "&Layer"
Private Const ConfigSettingOffsetX                              As String = "&X"
Private Const ConfigSettingOffsetY                              As String = "&Y"
Private Const ConfigSettingProgId                               As String = "&ProgId"
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

Private mBarFormatterFactory As BarFormatterFactory
Private mBars As Bars

Private mConfig As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

mLayer = LayerNumbers.LayerHighestUser
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

Public Property Get BarFormatterFactory() As BarFormatterFactory
Set BarFormatterFactory = mBarFormatterFactory
End Property

Public Property Get Bars() As Bars
Set Bars = mBars
End Property
    
Public Property Let BarStyle(ByVal value As BarStyle)
Set mBarStyle = value
storeSettings
End Property

Public Property Get BarStyle() As BarStyle
Set BarStyle = mBarStyle
End Property

Public Property Let ChartRegionName(ByVal value As String)
mChartRegionName = value
storeSettings
End Property

Public Property Get ChartRegionName() As String
ChartRegionName = mChartRegionName
End Property

Friend Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
If value Is mConfig Then Exit Property
Set mConfig = value
storeSettings
End Property

Public Property Let DataPointStyle(ByVal value As DataPointStyle)
Set mDataPointStyle = value
storeSettings
End Property

Public Property Get DataPointStyle() As DataPointStyle
Set DataPointStyle = mDataPointStyle
End Property

Public Property Let IncludeInChart(ByVal value As Boolean)
mIncludeInChart = value
storeSettings
End Property

Public Property Get IncludeInChart() As Boolean
IncludeInChart = mIncludeInChart
End Property

Public Property Let Layer(ByVal value As Long)
mLayer = value
storeSettings
End Property

Public Property Get Layer() As Long
Layer = mLayer
End Property

Public Property Let LineStyle(ByVal value As LineStyle)
Set mLineStyle = value
storeSettings
End Property

Public Property Get LineStyle() As LineStyle
Set LineStyle = mLineStyle
End Property

Public Property Let OffsetX(ByVal value As Double)
mOffsetX = value
storeSettings
End Property

Public Property Get OffsetX() As Double
OffsetX = mOffsetX
End Property

Public Property Let OffsetY(ByVal value As Double)
mOffsetY = value
storeSettings
End Property

Public Property Get OffsetY() As Double
OffsetY = mOffsetY
End Property

Public Property Let TextStyle(ByVal value As TextStyle)
Set mTextStyle = value
storeSettings
End Property

Public Property Get TextStyle() As TextStyle
Set TextStyle = mTextStyle
End Property

Friend Property Let ValueName(ByVal value As String)
mValueName = value
storeSettings
End Property

Public Property Get ValueName() As String
ValueName = mValueName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub RemoveFromConfig()
If Not mConfig Is Nothing Then mConfig.Remove
End Sub

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection, _
                ByVal pBars As Bars)
Dim cs As ConfigurationSection

Set mBars = pBars
Set mConfig = config

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

Set cs = mConfig.GetConfigurationSection(ConfigSectionBarFormatterFactory)
If Not cs Is Nothing Then
    Set mBarFormatterFactory = CreateObject(cs.GetSetting(ConfigSettingProgId))
    mBarFormatterFactory.LoadFromConfig cs
End If

mIncludeInChart = mConfig.GetSetting(ConfigSettingIncludeInChart)
mValueName = mConfig.GetSetting(ConfigSettingValueName)
mChartRegionName = mConfig.GetSetting(ConfigSettingChartRegionName)
mLayer = mConfig.GetSetting(ConfigSettingLayer)
mOffsetX = mConfig.GetSetting(ConfigSettingOffsetX)
mOffsetY = mConfig.GetSetting(ConfigSettingOffsetY)
End Sub

Public Sub SetBarFormatterFactory(ByVal formatterFactory As BarFormatterFactory, ByVal Bars As Bars)
If Not formatterFactory Is Nothing And Bars Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "SetBarFormatterFactory", _
            "A Bars object must be supplied"
End If
Set mBarFormatterFactory = formatterFactory
Set mBars = Bars
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub storeSettings()
If mConfig Is Nothing Then Exit Sub

If Not mBarStyle Is Nothing Then mBarStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionBarStyle)
If Not mDataPointStyle Is Nothing Then mDataPointStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDataPointStyle)
If Not mLineStyle Is Nothing Then mLineStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionLineStyle)
If Not mTextStyle Is Nothing Then mTextStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionTextStyle)
If Not mBarFormatterFactory Is Nothing Then
    Dim bfSect As ConfigurationSection
    Set bfSect = mConfig.AddConfigurationSection(ConfigSectionBarFormatterFactory)
    bfSect.SetSetting ConfigSettingProgId, GetProgIdFromObject(mBarFormatterFactory)
    mBarFormatterFactory.ConfigurationSection = bfSect
End If

mConfig.SetSetting ConfigSettingIncludeInChart, mIncludeInChart
mConfig.SetSetting ConfigSettingValueName, mValueName
mConfig.SetSetting ConfigSettingChartRegionName, mChartRegionName
mConfig.SetSetting ConfigSettingLayer, mLayer
mConfig.SetSetting ConfigSettingOffsetX, mOffsetX
mConfig.SetSetting ConfigSettingOffsetY, mOffsetY
End Sub


