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

Private Const ConfigSettingMinimumTicksHeight           As String = "&MinimumTicksHeight"
Private Const ConfigSettingTwipsPerPeriod               As String = "&TwipsPerPeriod"

Private Const ConfigSectionDefaultRegionStyle           As String = "DefaultRegionStyle"
Private Const ConfigSectionDefaultYAxisRegionStyle      As String = "DefaultYAxisRegionStyle"
Private Const ConfigSectionXAxisRegionStyle             As String = "XAxisRegionStyle"
Private Const ConfigSectionVolumeRegionStyle            As String = "VolumeRegionStyle"
Private Const ConfigSectionBarsStyle                    As String = "BarsStyle"
Private Const ConfigSectionVolumeStyle                  As String = "VolumeStyle"

'@================================================================================
' Member variables
'@================================================================================

Private mName                               As String
Private mMinimumTicksHeight                 As Long
Private mTwipsPerPeriod                     As Long
Private mDefaultRegionStyle                 As ChartRegionStyle
Private mDefaultYAxisRegionStyle            As ChartRegionStyle
Private mXAxisRegionStyle                   As ChartRegionStyle
Private mVolumeRegionStyle                  As ChartRegionStyle
Private mBarsStyle                          As BarStyle
Private mVolumeStyle                        As DataPointStyle

Private mConfig                             As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get BarsStyle() As BarStyle
Set BarsStyle = mBarsStyle
End Property

Friend Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If value Is mConfig Then Exit Property
Set mConfig = value
If Not mConfig Is Nothing Then
    mConfig.SetSetting ConfigSettingMinimumTicksHeight, mMinimumTicksHeight
    mConfig.SetSetting ConfigSettingTwipsPerPeriod, mTwipsPerPeriod
    mDefaultRegionStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultRegionStyle)
    mVolumeRegionStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionVolumeRegionStyle)
    mXAxisRegionStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXAxisRegionStyle)
    mDefaultYAxisRegionStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultYAxisRegionStyle)
    mBarsStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionBarsStyle)
    mVolumeStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionVolumeStyle)
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DefaultRegionStyle() As ChartRegionStyle
Const ProcName As String = "DefaultRegionStyle"
On Error GoTo Err

Set DefaultRegionStyle = mDefaultRegionStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DefaultYAxisRegionStyle() As ChartRegionStyle
Const ProcName As String = "DefaultYAxisRegionStyle"
On Error GoTo Err

Set DefaultYAxisRegionStyle = mDefaultYAxisRegionStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MinimumTicksHeight() As Long
Const ProcName As String = "MinimumTicksHeight"
On Error GoTo Err

MinimumTicksHeight = mMinimumTicksHeight

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Get TwipsPerPeriod() As Long
Const ProcName As String = "TwipsPerPeriod"
On Error GoTo Err

TwipsPerPeriod = mTwipsPerPeriod

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get VolumeRegionStyle() As ChartRegionStyle
Const ProcName As String = "VolumeRegionStyle"
On Error GoTo Err

Set VolumeRegionStyle = mVolumeRegionStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get VolumeStyle() As DataPointStyle
Const ProcName As String = "VolumeStyle"
On Error GoTo Err

Set VolumeStyle = mVolumeStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get XAxisRegionStyle() As ChartRegionStyle
Const ProcName As String = "XAxisRegionStyle"
On Error GoTo Err

Set XAxisRegionStyle = mXAxisRegionStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function CloneWithNewName(ByVal pName As String) As ChartStyle
Const ProcName As String = "CloneWithNewName"
On Error GoTo Err

Set CloneWithNewName = New ChartStyle
CloneWithNewName.Initialise pName, _
                mTwipsPerPeriod, _
                mMinimumTicksHeight, _
                mDefaultRegionStyle.Clone, _
                mVolumeRegionStyle.Clone, _
                mXAxisRegionStyle.Clone, _
                mDefaultYAxisRegionStyle.Clone, _
                mBarsStyle.Clone, _
                mVolumeStyle.Clone

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub Initialise( _
                ByVal pName As String, _
                ByVal pTwipsPerPeriod As Long, _
                ByVal pMinimumTicksHeight As Long, _
                ByVal pDefaultRegionStyle As ChartRegionStyle, _
                ByVal pVolumeRegionStyle As ChartRegionStyle, _
                ByVal pXAxisRegionStyle As ChartRegionStyle, _
                ByVal pDefaultYAxisRegionStyle As ChartRegionStyle, _
                ByVal pBarsStyle As BarStyle, _
                ByVal pVolumeStyle As DataPointStyle)
Const ProcName As String = "Initialise"
On Error GoTo Err

mName = pName
mTwipsPerPeriod = pTwipsPerPeriod
mMinimumTicksHeight = MinimumTicksHeight

If Not pDefaultRegionStyle Is Nothing Then
    Set mDefaultRegionStyle = pDefaultRegionStyle
Else
    Set mDefaultRegionStyle = GetDefaultChartRegionStyle.Clone
    mDefaultRegionStyle.Autoscaling = True
    mDefaultRegionStyle.GridlineSpacingY = 1.8
    mDefaultRegionStyle.HasGrid = True
    mDefaultRegionStyle.CursorSnapsToTickBoundaries = True
End If

If Not pVolumeRegionStyle Is Nothing Then
    Set mVolumeRegionStyle = pVolumeRegionStyle
Else
    Set mVolumeRegionStyle = mDefaultRegionStyle.Clone
    mVolumeRegionStyle.GridlineSpacingY = 0.8
    mVolumeRegionStyle.MinimumHeight = 10
    mVolumeRegionStyle.IntegerYScale = True
End If

If Not pXAxisRegionStyle Is Nothing Then
    Set mXAxisRegionStyle = pXAxisRegionStyle
Else
    Set mXAxisRegionStyle = mDefaultRegionStyle.Clone
    mXAxisRegionStyle.HasGrid = False
    mXAxisRegionStyle.HasGridText = True
End If

If Not pDefaultYAxisRegionStyle Is Nothing Then
    Set mDefaultYAxisRegionStyle = pDefaultYAxisRegionStyle
Else
    Set mDefaultYAxisRegionStyle = mDefaultRegionStyle.Clone
    mDefaultYAxisRegionStyle.HasGrid = False
End If

If Not pBarsStyle Is Nothing Then
    Set mBarsStyle = pBarsStyle
Else
    Set mBarsStyle = GetDefaultBarStyle.Clone
    mBarsStyle.Thickness = 2
    mBarsStyle.Width = 0.6
    mBarsStyle.DisplayMode = BarDisplayModeCandlestick
    mBarsStyle.DownColor = &H43FC2
    mBarsStyle.IncludeInAutoscale = True
    mBarsStyle.OutlineThickness = 1
    mBarsStyle.SolidUpBody = False
    mBarsStyle.TailThickness = 1
    mBarsStyle.UpColor = &H1D9311
End If

If Not pVolumeStyle Is Nothing Then
    Set mVolumeStyle = pVolumeStyle
Else
    Set mVolumeStyle = GetDefaultDataPointStyle.Clone
    mVolumeStyle.DisplayMode = DataPointDisplayModeHistogram
    mVolumeStyle.DownColor = &H43FC2
    mVolumeStyle.HistogramBarWidth = 0.6
    mVolumeStyle.IncludeInAutoscale = True
    mVolumeStyle.LineStyle = LineSolid
    mVolumeStyle.LineThickness = 1
    mVolumeStyle.PointStyle = PointRound
    mVolumeStyle.UpColor = &H1D9311
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Dim ar() As String
Dim i As Long

Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = config
If mConfig Is Nothing Then Exit Sub

mName = mConfig.InstanceQualifier
mMinimumTicksHeight = mConfig.GetSetting(ConfigSettingMinimumTicksHeight, 20)
mTwipsPerPeriod = mConfig.GetSetting(ConfigSettingTwipsPerPeriod, 100)

Set mDefaultRegionStyle = New ChartRegionStyle
mDefaultRegionStyle.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionDefaultRegionStyle)

Set mVolumeRegionStyle = New ChartRegionStyle
mVolumeRegionStyle.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionVolumeRegionStyle)

Set mXAxisRegionStyle = New ChartRegionStyle
mXAxisRegionStyle.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionXAxisRegionStyle)

Set mDefaultYAxisRegionStyle = New ChartRegionStyle
mDefaultYAxisRegionStyle.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionDefaultYAxisRegionStyle)

Set mBarsStyle = New BarStyle
mBarsStyle.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionBarsStyle)

Set mVolumeStyle = New DataPointStyle
mVolumeStyle.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionVolumeStyle)

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


