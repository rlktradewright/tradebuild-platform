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

Private Const ProjectName                As String = "ChartUtils26"
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

'@================================================================================
' Methods
'@================================================================================

Public Function CreateAbstractChartTool( _
                ByVal controller As chartController, _
                ByVal tool As IChartTool, _
                Optional ByVal pointerIcon As IPictureDisp) As AbstractChartTool
Dim failpoint As Long
On Error GoTo Err

Set CreateAbstractChartTool = New AbstractChartTool
CreateAbstractChartTool.Initialise controller, tool, pointerIcon

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
                ByVal chartController As chartController) As ChartManager
Dim failpoint As Long
On Error GoTo Err

Set CreateChartManager = New ChartManager
CreateChartManager.Initialise studyManager, chartController

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "CreateChartManager" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function CreateChartSpecifier( _
                Optional ByVal timeframe As TimePeriod, _
                Optional ByVal InitialNumberOfBars As Integer = 200, _
                Optional ByVal IncludeBarsOutsideSession As Boolean = False, _
                Optional ByVal MinimumTicksHeight As Integer = 20, _
                Optional ByVal DefaultRegionStyle As ChartRegionStyle, _
                Optional ByVal VolumeRegionStyle As ChartRegionStyle, _
                Optional ByVal XAxisRegionStyle As ChartRegionStyle, _
                Optional ByVal DefaultYAxisRegionStyle As ChartRegionStyle, _
                Optional ByVal pBarsStyle As barStyle, _
                Optional ByVal VolumeStyle As dataPointStyle) As ChartSpecifier

Dim failpoint As Long
On Error GoTo Err

If timeframe Is Nothing Then Set timeframe = GetTimePeriod(5, TimePeriodMinute)

Set CreateChartSpecifier = New ChartSpecifier
CreateChartSpecifier.Initialise timeframe, _
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

'@================================================================================
' Helper Functions
'@================================================================================


