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

Private Const ProjectName                   As String = "TradeBuildUI26"
Private Const ModuleName                    As String = "ChartSpecifier"

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
Private mBarsStyle                          As barStyle
Private mVolumeStyle                        As dataPointStyle


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

Public Property Get BarsStyle() As barStyle
Set BarsStyle = mBarsStyle
End Property

Public Property Get DefaultRegionStyle() As ChartRegionStyle
Set DefaultRegionStyle = mDefaultRegionStyle
End Property

Public Property Get DefaultYAxisRegionStyle() As ChartRegionStyle
Set DefaultYAxisRegionStyle = mDefaultYAxisRegionStyle
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

Public Property Get timeframe() As TimePeriod
Set timeframe = mTimeframe
End Property

Friend Property Let timeframe(ByVal value As TimePeriod)
Set mTimeframe = value
End Property

Public Property Get VolumeRegionStyle() As ChartRegionStyle
Set VolumeRegionStyle = mVolumeRegionStyle
End Property

Public Property Get VolumeStyle() As dataPointStyle
Set VolumeStyle = mVolumeStyle
End Property

Public Property Get XAxisRegionStyle() As ChartRegionStyle
Set XAxisRegionStyle = mXAxisRegionStyle
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal timeframe As TimePeriod, _
                ByVal InitialNumberOfBars As Integer, _
                ByVal IncludeBarsOutsideSession As Boolean, _
                ByVal MinimumTicksHeight As Integer, _
                ByVal DefaultRegionStyle As ChartRegionStyle, _
                ByVal VolumeRegionStyle As ChartRegionStyle, _
                ByVal XAxisRegionStyle As ChartRegionStyle, _
                ByVal DefaultYAxisRegionStyle As ChartRegionStyle, _
                ByVal pBarsStyle As barStyle, _
                ByVal VolumeStyle As dataPointStyle)
Set mTimeframe = timeframe
validateTimeFrame (timeframe)
mInitialNumberOfBars = InitialNumberOfBars
mIncludeBarsOutsideSession = IncludeBarsOutsideSession
mMinimumTicksHeight = MinimumTicksHeight
Set mDefaultRegionStyle = DefaultRegionStyle
Set mVolumeRegionStyle = VolumeRegionStyle
Set mXAxisRegionStyle = XAxisRegionStyle
Set mDefaultYAxisRegionStyle = DefaultYAxisRegionStyle
Set mBarsStyle = pBarsStyle
Set mVolumeStyle = VolumeStyle
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub validateTimeFrame(ByVal tf As TimePeriod)

Select Case tf.units
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

