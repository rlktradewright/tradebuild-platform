VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mChart As Chart

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let Chart( _
                ByVal value As Chart)
Set mChart = value
End Property

Public Property Get allowHorizontalMouseScrolling() As Boolean
allowHorizontalMouseScrolling = mChart.allowHorizontalMouseScrolling
End Property

Public Property Let allowHorizontalMouseScrolling(ByVal value As Boolean)
mChart.allowHorizontalMouseScrolling = value
End Property

Public Property Get allowVerticalMouseScrolling() As Boolean
allowVerticalMouseScrolling = mChart.allowVerticalMouseScrolling
End Property

Public Property Let allowVerticalMouseScrolling(ByVal value As Boolean)
mChart.allowVerticalMouseScrolling = value
End Property

Public Property Get autoscale() As Boolean
autoscale = mChart.autoscale
End Property

Public Property Let autoscale(ByVal value As Boolean)
mChart.autoscale = value
End Property

Public Property Get chartBackColor() As OLE_COLOR
chartBackColor = mChart.backColor
End Property

Public Property Let chartBackColor(ByVal val As OLE_COLOR)
mChart.chartBackColor = val
End Property

Public Property Get chartLeft() As Single
chartLeft = mChart.chartLeft
End Property

Public Property Get chartWidth() As Single
chartWidth = mChart.chartWidth
End Property

Public Property Get currentPeriodNumber() As Long
currentPeriodNumber = mChart.currentPeriodNumber
End Property

Public Property Get firstVisiblePeriod() As Long
firstVisiblePeriod = mChart.firstVisiblePeriod
End Property

Public Property Let firstVisiblePeriod(ByVal value As Long)
mChart.firstVisiblePeriod = value
End Property

Public Property Get gridColor() As OLE_COLOR
gridColor = mChart.gridColor
End Property

Public Property Let gridColor(ByVal val As OLE_COLOR)
mChart.gridColor = val
End Property

Public Property Get gridTextColor() As OLE_COLOR
gridTextColor = mChart.gridTextColor
End Property

Public Property Let gridTextColor(ByVal val As OLE_COLOR)
mChart.gridTextColor = val
End Property

Public Property Get lastVisiblePeriod() As Long
lastVisiblePeriod = mChart.lastVisiblePeriod
End Property

Public Property Let lastVisiblePeriod(ByVal value As Long)
mChart.lastVisiblePeriod = value
End Property

Public Property Get periodLengthMinutes() As Long
periodLengthMinutes = mChart.periodLengthMinutes
End Property

Public Property Let periodLengthMinutes(ByVal val As Long)
mChart.periodLengthMinutes = val
End Property

Public Property Get Periods() As Periods
Set Periods = mPeriods
End Property

Public Property Get sessionStartTime() As Date
sessionStartTime = mChart.sessionStartTime
End Property

Public Property Let sessionStartTime(ByVal val As Date)
mChart.sessionStartTime = val
End Property

Public Property Get showCrosshairs() As Boolean
showCrosshairs = mChart.showCrosshairs
End Property

Public Property Let showCrosshairs(ByVal val As Boolean)
mChart.showCrosshairs = val
End Property

Public Property Get showGrid() As Boolean
showGrid = mChart.showGrid
End Property

Public Property Let showGrid(ByVal val As Boolean)
mChart.showGrid = val
End Property

Public Property Get showHorizontalScrollBar() As Boolean
showHorizontalScrollBar = mChart.showHorizontalScrollBar
End Property

Public Property Let showHorizontalScrollBar(ByVal val As Boolean)
mChart.showHorizontalScrollBar = val
End Property

Public Property Get suppressDrawing() As Boolean
suppressDrawing = mChart.suppressDrawing
End Property

Public Property Let suppressDrawing(ByVal val As Boolean)
mChart.suppressDrawing = val
End Property

Public Property Get twipsPerBar() As Long
twipsPerBar = mChart.twipsPerBar
End Property

Public Property Let twipsPerBar(ByVal val As Long)
mChart.twipsPerBar = val
End Property

Public Property Let verticalGridSpacing(ByVal value As Long)
mChart.verticalGridSpacing = value
End Property

Public Property Get verticalGridSpacing() As Long
verticalGridSpacing = mChart.verticalGridSpacing
End Property

Public Property Let verticalGridUnits(ByVal value As TimeUnits)
mChart.verticalGridUnits = value
End Property

Public Property Get verticalGridUnits() As TimeUnits
verticalGridUnits = mChart.verticalGridUnits
End Property

Public Property Get YAxisPosition() As Long
YAxisPosition = mChart.YAxisPosition
End Property

Public Property Get YAxisWidthCm() As Single
YAxisWidthCm = mChart.YAxisWidthCm
End Property

Public Property Let YAxisWidthCm(ByVal value As Single)
mChart.YAxisWidthCm = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function addChartRegion(ByVal percentheight As Double, _
                    Optional ByVal minimumPercentHeight As Double, _
                    Optional ByVal name As String) As ChartRegion
Set addChartRegion = mChart.addChartRegion(percentheight, minimumPercentHeight, name)
End Function

Public Function addperiod(ByVal timestamp As Date) As period
Set addperiod = mChart.addperiod(timestamp)
End Function

Public Sub clearChart()
mChart.clearChart
End Sub

Public Sub refresh()
mChart.refresh
End Sub

Public Sub removeChartRegion( _
                    ByVal region As ChartRegion)
mChart.removeChartRegion region
End Sub

Public Sub scrollX(ByVal value As Long)
mChart.scrollX value
End Sub

'================================================================================
' Helper Functions
'================================================================================



