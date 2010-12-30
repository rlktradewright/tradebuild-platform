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
Const ProcName As String = "BarFormatterFactory"
On Error GoTo Err

Set BarFormatterFactory = mBarFormatterFactory

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Bars() As Bars
Const ProcName As String = "Bars"
On Error GoTo Err

Set Bars = mBars

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property
    
Public Property Let BarStyle(ByVal value As BarStyle)
Const ProcName As String = "BarStyle"
On Error GoTo Err

Set mBarStyle = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BarStyle() As BarStyle
Const ProcName As String = "BarStyle"
On Error GoTo Err

Set BarStyle = mBarStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ChartRegionName(ByVal value As String)
Const ProcName As String = "ChartRegionName"
On Error GoTo Err

mChartRegionName = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ChartRegionName() As String
Const ProcName As String = "ChartRegionName"
On Error GoTo Err

ChartRegionName = mChartRegionName

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If value Is mConfig Then Exit Property
Set mConfig = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let DataPointStyle(ByVal value As DataPointStyle)
Const ProcName As String = "DataPointStyle"
On Error GoTo Err

Set mDataPointStyle = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DataPointStyle() As DataPointStyle
Const ProcName As String = "DataPointStyle"
On Error GoTo Err

Set DataPointStyle = mDataPointStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let IncludeInChart(ByVal value As Boolean)
Const ProcName As String = "IncludeInChart"
On Error GoTo Err

mIncludeInChart = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IncludeInChart() As Boolean
Const ProcName As String = "IncludeInChart"
On Error GoTo Err

IncludeInChart = mIncludeInChart

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Layer(ByVal value As Long)
Const ProcName As String = "Layer"
On Error GoTo Err

mLayer = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Layer() As Long
Const ProcName As String = "Layer"
On Error GoTo Err

Layer = mLayer

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let LineStyle(ByVal value As LineStyle)
Const ProcName As String = "LineStyle"
On Error GoTo Err

Set mLineStyle = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LineStyle() As LineStyle
Const ProcName As String = "LineStyle"
On Error GoTo Err

Set LineStyle = mLineStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let OffsetX(ByVal value As Double)
Const ProcName As String = "OffsetX"
On Error GoTo Err

mOffsetX = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get OffsetX() As Double
Const ProcName As String = "OffsetX"
On Error GoTo Err

OffsetX = mOffsetX

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let OffsetY(ByVal value As Double)
Const ProcName As String = "OffsetY"
On Error GoTo Err

mOffsetY = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get OffsetY() As Double
Const ProcName As String = "OffsetY"
On Error GoTo Err

OffsetY = mOffsetY

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let TextStyle(ByVal value As TextStyle)
Const ProcName As String = "TextStyle"
On Error GoTo Err

Set mTextStyle = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get TextStyle() As TextStyle
Const ProcName As String = "TextStyle"
On Error GoTo Err

Set TextStyle = mTextStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let ValueName(ByVal value As String)
Const ProcName As String = "ValueName"
On Error GoTo Err

mValueName = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ValueName() As String
Const ProcName As String = "ValueName"
On Error GoTo Err

ValueName = mValueName

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection, _
                ByVal pBars As Bars)
Dim cs As ConfigurationSection

Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub SetBarFormatterFactory(ByVal formatterFactory As BarFormatterFactory, ByVal Bars As Bars)
Const ProcName As String = "SetBarFormatterFactory"
On Error GoTo Err

If Not formatterFactory Is Nothing And Bars Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "A Bars object must be supplied"
End If
Set mBarFormatterFactory = formatterFactory
Set mBars = Bars

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


