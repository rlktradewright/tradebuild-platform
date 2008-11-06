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

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Imports ChartSkil26
Imports TimeframeUtils26
Imports TradeBuild26

Public Class ChartSpecifier

    Private mTimeframe As TimePeriod
    Private mInitialNumberOfBars As Integer = 200
    Private mIncludeBarsOutsideSession As Boolean = False
    Private mMinimumTicksHeight As Integer = 20
    Private mDefaultRegionStyle As ChartRegionStyle
    Private mDefaultYAxisRegionStyle As ChartRegionStyle
    Private mXAxisRegionStyle As ChartRegionStyle
    Private mVolumeRegionStyle As ChartRegionStyle

    Public Sub New(ByVal timeframe As TimePeriod, _
                   ByVal initialNumberOfBars As Integer, _
                   ByVal includeBarsOutsideSession As Boolean, _
                   ByVal minimumTicksHeight As Integer, _
                   ByVal defaultRegionStyle As ChartRegionStyle, _
                   ByVal volumeRegionstyle As ChartRegionStyle, _
                   ByVal xAxisRegionStyle As ChartRegionStyle, _
                   ByVal defaultYAxisRegionStyle As ChartRegionStyle)
        mTimeframe = timeframe
        validateTimeFrame (timeframe)
        mInitialNumberOfBars = initialNumberOfBars
        mIncludeBarsOutsideSession = includeBarsOutsideSession
        mMinimumTicksHeight = minimumTicksHeight
        mDefaultRegionStyle = DefaultRegionStyle
        mVolumeRegionStyle = volumeRegionStyle
        mXAxisRegionStyle = xAxisRegionStyle
        mDefaultYAxisRegionStyle = defaultYAxisRegionStyle
    End Sub

    Public Sub New(ByVal timeframe As TimePeriod)
        mTimeframe = timeframe
        validateTimeFrame (timeframe)
    End Sub

    Public Sub New(ByVal timeframe As TimePeriod, _
                   ByVal initialNumberOfBars As Integer)
        mTimeframe = timeframe
        validateTimeFrame (timeframe)
        mInitialNumberOfBars = initialNumberOfBars
    End Sub

    Public Sub New(ByVal timeframe As TimePeriod, _
                   ByVal initialNumberOfBars As Integer, _
                   ByVal includeBarsOutsideSession As Boolean)
        mTimeframe = timeframe
        validateTimeFrame (timeframe)
        mInitialNumberOfBars = initialNumberOfBars
        mIncludeBarsOutsideSession = includeBarsOutsideSession
    End Sub

    Public Sub New(ByVal timeframe As TimePeriod, _
                   ByVal initialNumberOfBars As Integer, _
                   ByVal includeBarsOutsideSession As Boolean, _
                   ByVal minimumTicksHeight As Integer)
        mTimeframe = timeframe
        validateTimeFrame (timeframe)
        mInitialNumberOfBars = initialNumberOfBars
        mIncludeBarsOutsideSession = includeBarsOutsideSession
        mMinimumTicksHeight = minimumTicksHeight
    End Sub

    Public ReadOnly Property defaultRegionStyle() As ChartRegionStyle
        Get
            Return mDefaultRegionStyle
        End Get
    End Property

    Public ReadOnly Property defaultYAxisRegionStyle() As ChartRegionStyle
        Get
            Return mDefaultYAxisRegionStyle
        End Get
    End Property

    Public ReadOnly Property initialNumberOfBars() As Integer
        Get
            Return mInitialNumberOfBars
        End Get
    End Property

    Public ReadOnly Property includeBarsOutsideSession() As Boolean
        Get
            Return mIncludeBarsOutsideSession
        End Get
    End Property

    Public ReadOnly Property minimumTicksHeight() As Integer
        Get
            Return mMinimumTicksHeight
        End Get
    End Property

    Public Property timeframe() As TimePeriod
        Get
            Return mTimeframe
        End Get
        Friend Set(ByVal value As TimePeriod)
            mTimeframe = value
        End Set
    End Property

    Public ReadOnly Property volumeRegionStyle() As ChartRegionStyle
        Get
            Return mVolumeRegionStyle
        End Get
    End Property

    Public ReadOnly Property xAxisRegionStyle() As ChartRegionStyle
        Get
            Return mXAxisRegionStyle
        End Get
    End Property

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
                Throw New ArgumentException("Time period units not supported")
        End Select

    End Sub

End Class

