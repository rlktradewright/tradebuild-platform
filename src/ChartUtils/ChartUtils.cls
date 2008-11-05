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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Public Type RegionEntry
    region              As ChartRegion
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
Set CreateAbstractChartTool = New AbstractChartTool
CreateAbstractChartTool.Initialise controller, tool, pointerIcon
End Function

Public Function CreateChartManager( _
                ByVal studyManager As studyManager, _
                ByVal chartController As chartController) As ChartManager
Set CreateChartManager = New ChartManager
CreateChartManager.Initialise studyManager, chartController
End Function

'@================================================================================
' Helper Functions
'@================================================================================


