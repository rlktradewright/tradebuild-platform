VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartUtils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
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


Private Const ModuleName                As String = "ChartUtils"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Public Type RegionEntry
    Region              As ChartRegion
    usageCount          As Long
End Type

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

Public Property Get ChartRegionNameCustom() As String
ChartRegionNameCustom = RegionNameCustom
End Property

Public Property Get ChartRegionNameDefault() As String
ChartRegionNameDefault = RegionNameDefault
End Property

Public Property Get ChartRegionNamePrice() As String
ChartRegionNamePrice = RegionNamePrice
End Property

Public Property Get ChartRegionNameVolume() As String
ChartRegionNameVolume = RegionNameVolume
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateAbstractChartTool( _
                ByVal pChart As ChartController, _
                ByVal tool As IChartTool, _
                Optional ByVal pointerIcon As IPictureDisp) As AbstractChartTool
Dim failpoint As Long
On Error GoTo Err

Set CreateAbstractChartTool = New AbstractChartTool
CreateAbstractChartTool.Initialise pChart, tool, pointerIcon

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "CreateAbstractChartTool" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function CreateChartManager( _
                ByVal studyManager As studyManager, _
                ByVal pChart As ChartController) As ChartManager
Dim failpoint As Long
On Error GoTo Err

Set CreateChartManager = New ChartManager
CreateChartManager.Initialise studyManager, pChart

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "CreateChartManager" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function CreateChartSpecifier( _
                Optional ByVal Timeframe As TimePeriod, _
                Optional ByVal InitialNumberOfBars As Integer = 200, _
                Optional ByVal IncludeBarsOutsideSession As Boolean = False, _
                Optional ByVal MinimumTicksHeight As Integer = 20, _
                Optional ByVal DefaultRegionStyle As ChartRegionStyle, _
                Optional ByVal VolumeRegionStyle As ChartRegionStyle, _
                Optional ByVal XAxisRegionStyle As ChartRegionStyle, _
                Optional ByVal DefaultYAxisRegionStyle As ChartRegionStyle, _
                Optional ByVal pBarsStyle As BarStyle, _
                Optional ByVal VolumeStyle As DataPointStyle) As ChartSpecifier

Dim failpoint As Long
On Error GoTo Err

If Timeframe Is Nothing Then Set Timeframe = GetTimePeriod(5, TimePeriodMinute)

Set CreateChartSpecifier = New ChartSpecifier
CreateChartSpecifier.Initialise Timeframe, _
                InitialNumberOfBars, _
                IncludeBarsOutsideSession, _
                MinimumTicksHeight, _
                DefaultRegionStyle, _
                VolumeRegionStyle, _
                XAxisRegionStyle, _
                DefaultYAxisRegionStyle, _
                pBarsStyle, _
                VolumeStyle

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "CreateChartSpecifier" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function GetDefaultStudyConfiguration( _
                ByVal Name As String, _
                ByVal studyLibName As String) As StudyConfiguration
Set GetDefaultStudyConfiguration = gGetDefaultStudyConfiguration(Name, studyLibName)
End Function

Public Function LoadChartSpecifierFromConfig( _
                ByVal config As ConfigurationSection) As ChartSpecifier
Dim failpoint As Long
On Error GoTo Err

If config Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "LoadChartSpecifierFromConfig", _
            "config argument is Nothing"
End If

Set LoadChartSpecifierFromConfig = New ChartSpecifier
LoadChartSpecifierFromConfig.LoadFromConfig config

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "LoadChartSpecifierFromConfig" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Sub LoadDefaultStudyConfigurationsFromConfig( _
                ByVal config As ConfigurationSection)
gLoadDefaultStudyConfigurationsFromConfig config
End Sub

Public Sub SetDefaultStudyConfiguration( _
                ByVal value As StudyConfiguration)
gSetDefaultStudyConfiguration value
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


