VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartSpecifier"
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

Private Const ProjectName                               As String = "TradeBuildUI26"
Private Const ModuleName                                As String = "ChartSpecifier"

Private Const ConfigSettingTimeframe                    As String = ".Timefame"
Private Const ConfigSettingInitialNumberOfBars          As String = ".InitialNumberOfBars"
Private Const ConfigSettingIncludeBarsOutsideSession    As String = ".IncludeBarsOutsideSession"
Private Const ConfigSettingMinimumTicksHeight           As String = ".MinimumTicksHeight"
Private Const ConfigSettingTwipsPerBar                  As String = ".TwipsPerBar"
Private Const ConfigSettingChartBackColor               As String = ".ChartBackColor"

Private Const ConfigSectionDefaultRegionStyle           As String = "DefaultRegionStyle"
Private Const ConfigSectionDefaultYAxisRegionStyle      As String = "DefaultYAxisRegionStyle"
Private Const ConfigSectionXAxisRegionStyle             As String = "XAxisRegionStyle"
Private Const ConfigSectionVolumeRegionStyle            As String = "VolumeRegionStyle"
Private Const ConfigSectionBarsStyle                    As String = "BarsStyle"
Private Const ConfigSectionVolumeStyle                  As String = "VolumeStyle"

'@================================================================================
' Member variables
'@================================================================================

Private mTimeframe                          As TimePeriod
Private mInitialNumberOfBars                As Integer
Private mIncludeBarsOutsideSession          As Boolean
Private mMinimumTicksHeight                 As Integer
Private mDefaultRegionStyle                 As ChartRegionStyle
Private mDefaultYAxisRegionStyle            As ChartRegionStyle
Private mXAxisRegionStyle                   As ChartRegionStyle
Private mVolumeRegionStyle                  As ChartRegionStyle
Private mBarsStyle                          As BarStyle
Private mVolumeStyle                        As DataPointStyle
Private mTwipsPerBar                        As Long
Private mChartBackColor                     As Long

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

Public Property Get ChartBackColor() As OLE_COLOR
ChartBackColor = mChartBackColor
End Property

Public Property Let ChartBackColor(ByVal value As OLE_COLOR)
mChartBackColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingChartBackColor, CStr(value)
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
If value Is mConfig Then Exit Property
Set mConfig = value
If Not mConfig Is Nothing Then
    mConfig.SetSetting ConfigSettingTimeframe, mTimeframe.toString
    mConfig.SetSetting ConfigSettingInitialNumberOfBars, mInitialNumberOfBars
    mConfig.SetBooleanSetting ConfigSettingIncludeBarsOutsideSession, mIncludeBarsOutsideSession
    mConfig.SetSetting ConfigSettingMinimumTicksHeight, mMinimumTicksHeight
    mConfig.SetSetting ConfigSettingChartBackColor, CStr(mChartBackColor)
    mDefaultRegionStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultRegionStyle)
    mVolumeRegionStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionVolumeRegionStyle)
    mXAxisRegionStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionXAxisRegionStyle)
    mDefaultYAxisRegionStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultYAxisRegionStyle)
    mBarsStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionBarsStyle)
    mVolumeStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionVolumeStyle)
    mConfig.SetSetting ConfigSettingTwipsPerBar, mTwipsPerBar
End If
End Property

Public Property Get DefaultRegionStyle() As ChartRegionStyle
Set DefaultRegionStyle = mDefaultRegionStyle
End Property

Public Property Get DefaultYAxisRegionStyle() As ChartRegionStyle
Set DefaultYAxisRegionStyle = mDefaultYAxisRegionStyle
End Property

Public Property Let InitialNumberOfBars( _
                ByVal value As Integer)
mInitialNumberOfBars = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingInitialNumberOfBars, mInitialNumberOfBars
End Property

Public Property Get InitialNumberOfBars() As Integer
InitialNumberOfBars = mInitialNumberOfBars
End Property

Public Property Get IncludeBarsOutsideSession() As Boolean
IncludeBarsOutsideSession = mIncludeBarsOutsideSession
End Property

Public Property Get MinimumTicksHeight() As Integer
MinimumTicksHeight = mMinimumTicksHeight
End Property

Public Property Get Timeframe() As TimePeriod
Set Timeframe = mTimeframe
End Property

Public Property Let Timeframe(ByVal value As TimePeriod)
Set mTimeframe = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingTimeframe, mTimeframe.toString
End Property

Public Property Get TwipsPerBar() As Long
TwipsPerBar = mTwipsPerBar
End Property

Public Property Let TwipsPerBar(ByVal value As Long)
mTwipsPerBar = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingTwipsPerBar, mTwipsPerBar
End Property

Public Property Get VolumeRegionStyle() As ChartRegionStyle
Set VolumeRegionStyle = mVolumeRegionStyle
End Property

Public Property Get VolumeStyle() As DataPointStyle
Set VolumeStyle = mVolumeStyle
End Property

Public Property Get XAxisRegionStyle() As ChartRegionStyle
Set XAxisRegionStyle = mXAxisRegionStyle
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Clone() As ChartSpecifier
Set Clone = New ChartSpecifier
Clone.Initialise mTimeframe, _
                mInitialNumberOfBars, _
                mIncludeBarsOutsideSession, _
                mMinimumTicksHeight, _
                mDefaultRegionStyle, _
                mVolumeRegionStyle, _
                mXAxisRegionStyle, _
                mDefaultYAxisRegionStyle, _
                mBarsStyle, _
                mVolumeStyle
Clone.ChartBackColor = mChartBackColor
Clone.TwipsPerBar = mTwipsPerBar
End Function

Friend Sub Initialise( _
                ByVal Timeframe As TimePeriod, _
                ByVal InitialNumberOfBars As Integer, _
                ByVal IncludeBarsOutsideSession As Boolean, _
                ByVal MinimumTicksHeight As Integer, _
                ByVal DefaultRegionStyle As ChartRegionStyle, _
                ByVal VolumeRegionStyle As ChartRegionStyle, _
                ByVal XAxisRegionStyle As ChartRegionStyle, _
                ByVal DefaultYAxisRegionStyle As ChartRegionStyle, _
                ByVal pBarsStyle As BarStyle, _
                ByVal VolumeStyle As DataPointStyle)
Set mTimeframe = Timeframe
validateTimeFrame Timeframe
mInitialNumberOfBars = InitialNumberOfBars
mIncludeBarsOutsideSession = IncludeBarsOutsideSession
mMinimumTicksHeight = MinimumTicksHeight
Set mDefaultRegionStyle = DefaultRegionStyle.Clone
Set mVolumeRegionStyle = VolumeRegionStyle.Clone
Set mXAxisRegionStyle = XAxisRegionStyle.Clone
Set mDefaultYAxisRegionStyle = DefaultYAxisRegionStyle.Clone
Set mBarsStyle = pBarsStyle.Clone
Set mVolumeStyle = VolumeStyle.Clone

mChartBackColor = &H602008

End Sub

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Dim ar() As String
Dim i As Long

Set mConfig = config
If mConfig Is Nothing Then Exit Sub

Set mTimeframe = TimePeriodFromString(mConfig.GetSetting(ConfigSettingTimeframe, "5 Minutes"))

mInitialNumberOfBars = mConfig.GetSetting(ConfigSettingInitialNumberOfBars, 500)
mIncludeBarsOutsideSession = CBool(mConfig.GetSetting(ConfigSettingIncludeBarsOutsideSession, "False"))
mMinimumTicksHeight = mConfig.GetSetting(ConfigSettingMinimumTicksHeight, 20)

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

mChartBackColor = mConfig.GetSetting(ConfigSettingChartBackColor, CStr(&H602008))

mTwipsPerBar = mConfig.GetSetting(ConfigSettingTwipsPerBar, 100)

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub validateTimeFrame(ByVal tf As TimePeriod)

Select Case tf.Units
Case TimePeriodUnits.TimePeriodSecond, _
        TimePeriodUnits.TimePeriodMinute, _
        TimePeriodUnits.TimePeriodHour, _
        TimePeriodUnits.TimePeriodDay, _
        TimePeriodUnits.TimePeriodWeek, _
        TimePeriodUnits.TimePeriodMonth, _
        TimePeriodUnits.TimePeriodYear, _
        TimePeriodUnits.TimePeriodVolume, _
        TimePeriodUnits.TimePeriodTickMovement
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "validateTimeFrame", _
            "Time period units not supported"
End Select

End Sub

