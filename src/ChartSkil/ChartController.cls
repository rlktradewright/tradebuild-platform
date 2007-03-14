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

Event ChartCleared()
Event RegionsChanged(ev As CollectionChangeEvent)
Event PeriodsChanged(ev As CollectionChangeEvent)

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

Private mChartRef                   As WeakReference
Attribute mChartRef.VB_VarHelpID = -1

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get allowHorizontalMouseScrolling() As Boolean
allowHorizontalMouseScrolling = chartObj.allowHorizontalMouseScrolling
End Property

Public Property Let allowHorizontalMouseScrolling(ByVal value As Boolean)
chartObj.allowHorizontalMouseScrolling = value
End Property

Public Property Get allowVerticalMouseScrolling() As Boolean
allowVerticalMouseScrolling = chartObj.allowVerticalMouseScrolling
End Property

Public Property Let allowVerticalMouseScrolling(ByVal value As Boolean)
chartObj.allowVerticalMouseScrolling = value
End Property

Public Property Get autoscroll() As Boolean
autoscroll = chartObj.autoscroll
End Property

Public Property Let autoscroll(ByVal value As Boolean)
chartObj.autoscroll = value
End Property

Friend Property Let Chart( _
                ByVal value As Chart)
Set mChartRef = createWeakReference(value)
End Property

Public Property Get chartBackColor() As OLE_COLOR
chartBackColor = chartObj.chartBackColor
End Property

Public Property Let chartBackColor(ByVal val As OLE_COLOR)
chartObj.chartBackColor = val
End Property

Public Property Get chartLeft() As Single
chartLeft = chartObj.chartLeft
End Property

Public Property Get chartWidth() As Single
chartWidth = chartObj.chartWidth
End Property

Public Property Get currentPeriodNumber() As Long
currentPeriodNumber = chartObj.currentPeriodNumber
End Property

Public Property Get currentSessionEndTime() As Date
currentSessionEndTime = chartObj.currentSessionEndTime
End Property

Public Property Get currentSessionStartTime() As Date
currentSessionStartTime = chartObj.currentSessionStartTime
End Property

Public Property Get defaultRegionStyle() As ChartRegionStyle
Set defaultRegionStyle = chartObj.defaultRegionStyle
End Property

Public Property Let defaultRegionStyle( _
                ByVal value As ChartRegionStyle)
chartObj.defaultRegionStyle = value
End Property

Public Property Get firstVisiblePeriod() As Long
firstVisiblePeriod = chartObj.firstVisiblePeriod
End Property

Public Property Let firstVisiblePeriod(ByVal value As Long)
chartObj.firstVisiblePeriod = value
End Property

Public Property Get lastVisiblePeriod() As Long
lastVisiblePeriod = chartObj.lastVisiblePeriod
End Property

Public Property Let lastVisiblePeriod(ByVal value As Long)
chartObj.lastVisiblePeriod = value
End Property

Public Property Get periodLength() As Long
periodLength = chartObj.periodLength
End Property

Public Property Get Periods() As Periods
Set Periods = chartObj.Periods
End Property

Public Property Get periodUnits() As TimePeriodUnits
periodUnits = chartObj.periodUnits
End Property

Public Property Let pointerStyle(ByVal value As PointerStyles)
chartObj.pointerStyle = value
End Property

Public Property Get PointerCrosshairsColor() As Long
PointerCrosshairsColor = chartObj.PointerCrosshairsColor
End Property

Public Property Let PointerCrosshairsColor(ByVal value As Long)
chartObj.PointerCrosshairsColor = value
End Property

Public Property Get PointerDiscColor() As Long
PointerDiscColor = chartObj.PointerDiscColor
End Property

Public Property Let PointerDiscColor(ByVal value As Long)
chartObj.PointerDiscColor = value
End Property

Public Property Get pointerStyle() As PointerStyles
pointerStyle = chartObj.pointerStyle
End Property

Public Property Get sessionEndTime() As Date
sessionEndTime = chartObj.sessionEndTime
End Property

Public Property Let sessionEndTime(ByVal val As Date)
chartObj.sessionEndTime = val
End Property

Public Property Get sessionStartTime() As Date
sessionStartTime = chartObj.sessionStartTime
End Property

Public Property Let sessionStartTime(ByVal val As Date)
chartObj.sessionStartTime = val
End Property

Public Property Get showHorizontalScrollBar() As Boolean
showHorizontalScrollBar = chartObj.showHorizontalScrollBar
End Property

Public Property Let showHorizontalScrollBar(ByVal val As Boolean)
chartObj.showHorizontalScrollBar = val
End Property

Public Property Get showToolbar() As Boolean
showToolbar = chartObj.showToolbar
End Property

Public Property Let showToolbar(ByVal val As Boolean)
chartObj.showToolbar = val
End Property

Public Property Get suppressDrawing() As Boolean
suppressDrawing = chartObj.suppressDrawing
End Property

Public Property Let suppressDrawing(ByVal val As Boolean)
chartObj.suppressDrawing = val
End Property

Public Property Get twipsPerBar() As Long
twipsPerBar = chartObj.twipsPerBar
End Property

Public Property Let twipsPerBar(ByVal val As Long)
chartObj.twipsPerBar = val
End Property

Public Property Get YAxisPosition() As Long
YAxisPosition = chartObj.YAxisPosition
End Property

Public Property Get YAxisWidthCm() As Single
YAxisWidthCm = chartObj.YAxisWidthCm
End Property

Public Property Let YAxisWidthCm(ByVal value As Single)
chartObj.YAxisWidthCm = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function addChartRegion(ByVal percentheight As Double, _
                    Optional ByVal minimumPercentHeight As Double, _
                    Optional ByVal style As ChartRegionStyle, _
                    Optional ByVal name As String) As ChartRegion
Set addChartRegion = chartObj.addChartRegion(percentheight, minimumPercentHeight, style, name)
End Function

Public Function addPeriod(ByVal timestamp As Date) As period
Set addPeriod = chartObj.addPeriod(timestamp)
End Function

Public Sub clearChart()
chartObj.clearChart
End Sub

Public Sub displayGrid()
chartObj.displayGrid
End Sub

Friend Sub fireChartCleared()
RaiseEvent ChartCleared
End Sub

Friend Sub fireRegionsChanged(ev As CollectionChangeEvent)
RaiseEvent RegionsChanged(ev)
End Sub

Friend Sub firePeriodsChanged(ev As CollectionChangeEvent)
RaiseEvent PeriodsChanged(ev)
End Sub

Public Function getChartRegion( _
                ByVal name As String) As ChartRegion
Set getChartRegion = chartObj.getChartRegion(name)
End Function

Public Sub hideGrid()
chartObj.hideGrid
End Sub

Public Function isGridHidden() As Boolean
isGridHidden = chartObj.isGridHidden
End Function

Public Function isTimeInSession(ByVal timestamp As Date) As Boolean
isTimeInSession = chartObj.isTimeInSession(timestamp)
End Function

Public Sub refresh()
chartObj.refresh
End Sub

Public Sub removeChartRegion( _
                    ByVal region As ChartRegion)
chartObj.removeChartRegion region
End Sub

Public Sub scrollX(ByVal value As Long)
chartObj.scrollX value
End Sub

Public Sub setPeriodParameters( _
                ByVal periodLength As Long, _
                ByVal periodUnits As TimePeriodUnits)
chartObj.setPeriodParameters periodLength, periodUnits
End Sub

Public Sub setVerticalGridParameters( _
                ByVal verticalGridSpacing As Long, _
                ByVal verticalGridUnits As TimePeriodUnits)
chartObj.setVerticalGridParameters verticalGridSpacing, verticalGridUnits
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function chartObj() As Chart
Set chartObj = mChartRef.Target
End Function

