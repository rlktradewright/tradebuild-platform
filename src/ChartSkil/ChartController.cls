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
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event PointerModeChanged()
Event RegionSelected(ByVal Region As ChartRegion)

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

Public Property Get Autoscrolling() As Boolean
Autoscrolling = chartObj.Autoscrolling
End Property

Public Property Let Autoscrolling(ByVal value As Boolean)
chartObj.Autoscrolling = value
End Property

Public Property Let BarTimePeriod( _
                ByVal value As TimePeriod)
chartObj.BarTimePeriod = value
End Property

Public Property Get BarTimePeriod() As TimePeriod
Set BarTimePeriod = chartObj.BarTimePeriod
End Property

Friend Property Let Chart( _
                ByVal value As Chart)
Set mChartRef = CreateWeakReference(value)
End Property

Public Property Get ChartBackColor() As OLE_COLOR
ChartBackColor = chartObj.ChartBackColor
End Property

Public Property Let ChartBackColor(ByVal val As OLE_COLOR)
chartObj.ChartBackColor = val
End Property

Public Property Get ChartLeft() As Single
ChartLeft = chartObj.ChartLeft
End Property

Public Property Get ChartWidth() As Single
ChartWidth = chartObj.ChartWidth
End Property

Public Property Get CurrentPeriodNumber() As Long
CurrentPeriodNumber = chartObj.CurrentPeriodNumber
End Property

Public Property Get CurrentSessionEndTime() As Date
CurrentSessionEndTime = chartObj.CurrentSessionEndTime
End Property

Public Property Get CurrentSessionStartTime() As Date
CurrentSessionStartTime = chartObj.CurrentSessionStartTime
End Property

Public Property Get FirstVisiblePeriod() As Long
FirstVisiblePeriod = chartObj.FirstVisiblePeriod
End Property

Public Property Let FirstVisiblePeriod(ByVal value As Long)
chartObj.FirstVisiblePeriod = value
End Property

Public Property Get HorizontalMouseScrollingAllowed() As Boolean
HorizontalMouseScrollingAllowed = chartObj.HorizontalMouseScrollingAllowed
End Property

Public Property Let HorizontalMouseScrollingAllowed(ByVal value As Boolean)
chartObj.HorizontalMouseScrollingAllowed = value
End Property

Public Property Get HorizontalScrollBarVisible() As Boolean
HorizontalScrollBarVisible = chartObj.HorizontalScrollBarVisible
End Property

Public Property Let HorizontalScrollBarVisible(ByVal val As Boolean)
chartObj.HorizontalScrollBarVisible = val
End Property

Public Property Get IsDrawingEnabled() As Boolean
IsDrawingEnabled = chartObj.IsDrawingEnabled
End Property

Public Property Get IsGridHidden() As Boolean
IsGridHidden = chartObj.IsGridHidden
End Property

Public Property Get LastVisiblePeriod() As Long
LastVisiblePeriod = chartObj.LastVisiblePeriod
End Property

Public Property Let LastVisiblePeriod(ByVal value As Long)
chartObj.LastVisiblePeriod = value
End Property

Public Property Get Periods() As Periods
Set Periods = chartObj.Periods
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

Public Property Get PointerIcon() As IPictureDisp
Set PointerIcon = chartObj.PointerIcon
End Property

Public Property Let PointerIcon(ByVal value As IPictureDisp)
chartObj.PointerIcon = value
End Property

Public Property Get PointerMode() As PointerModes
PointerMode = chartObj.PointerMode
End Property

Public Property Get PointerStyle() As PointerStyles
PointerStyle = chartObj.PointerStyle
End Property

Public Property Let PointerStyle(ByVal value As PointerStyles)
chartObj.PointerStyle = value
End Property

Public Property Get Regions() As ChartRegions
Set Regions = chartObj.Regions
End Property

Public Property Get SessionEndTime() As Date
SessionEndTime = chartObj.SessionEndTime
End Property

Public Property Let SessionEndTime(ByVal val As Date)
chartObj.SessionEndTime = val
End Property

Public Property Get SessionStartTime() As Date
SessionStartTime = chartObj.SessionStartTime
End Property

Public Property Let SessionStartTime(ByVal val As Date)
chartObj.SessionStartTime = val
End Property

Public Property Get TwipsPerBar() As Long
TwipsPerBar = chartObj.TwipsPerBar
End Property

Public Property Let TwipsPerBar(ByVal val As Long)
chartObj.TwipsPerBar = val
End Property

Public Property Let VerticalGridTimePeriod( _
                ByVal value As TimePeriod)
chartObj.VerticalGridTimePeriod = value
End Property

Public Property Get VerticalGridTimePeriod() As TimePeriod
Set VerticalGridTimePeriod = chartObj.VerticalGridTimePeriod
End Property

Public Property Get VerticalMouseScrollingAllowed() As Boolean
VerticalMouseScrollingAllowed = chartObj.VerticalMouseScrollingAllowed
End Property

Public Property Let VerticalMouseScrollingAllowed(ByVal value As Boolean)
chartObj.VerticalMouseScrollingAllowed = value
End Property

Public Property Get XAxisRegion() As ChartRegion
Set XAxisRegion = chartObj.XAxisRegion
End Property

Public Property Get XAxisVisible() As Boolean
XAxisVisible = chartObj.XAxisVisible
End Property

Public Property Let XAxisVisible(ByVal value As Boolean)
chartObj.XAxisVisible = value
End Property

Public Property Let XCursorTextStyle(ByVal value As TextStyle)
chartObj.XCursorTextStyle = value
End Property

Public Property Get XCursorTextStyle() As TextStyle
Set XCursorTextStyle = chartObj.XCursorTextStyle
End Property

Public Property Get YAxisPosition() As Long
YAxisPosition = chartObj.YAxisPosition
End Property

Public Property Get YAxisVisible() As Boolean
YAxisVisible = chartObj.YAxisVisible
End Property

Public Property Let YAxisVisible(ByVal value As Boolean)
chartObj.YAxisVisible = value
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

Public Sub ClearChart()
chartObj.ClearChart
End Sub

Public Sub DisableDrawing()
chartObj.DisableDrawing
End Sub

Public Sub EnableDrawing()
chartObj.EnableDrawing
End Sub

Public Sub Finish()
chartObj.Finish
End Sub

Friend Sub Finished()
Set mChartRef = Nothing
End Sub

Friend Sub fireChartCleared()
RaiseEvent ChartCleared
End Sub

Friend Sub fireKeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Friend Sub fireKeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Friend Sub fireKeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Friend Sub firePointerModeChanged()
RaiseEvent PointerModeChanged
End Sub

Friend Sub fireRegionSelected(ByVal Region As ChartRegion)
RaiseEvent RegionSelected(Region)
End Sub

Public Function GetXFromTimestamp( _
                ByVal Timestamp As Date, _
                Optional ByVal forceNewPeriod As Boolean, _
                Optional ByVal duplicateNumber As Long) As Double
GetXFromTimestamp = chartObj.GetXFromTimestamp(Timestamp, forceNewPeriod, duplicateNumber)
End Function

Public Sub HideGrid()
chartObj.HideGrid
End Sub

Public Function IsTimeInSession(ByVal Timestamp As Date) As Boolean
IsTimeInSession = chartObj.IsTimeInSession(Timestamp)
End Function

Public Sub ScrollX(ByVal value As Long)
chartObj.ScrollX value
End Sub

Public Sub SetPointerModeDefault()
chartObj.SetPointerModeDefault
End Sub

Public Sub SetPointerModeSelection()
chartObj.SetPointerModeSelection
End Sub

Public Sub SetPointerModeTool( _
                Optional ByVal toolPointerStyle As PointerStyles = PointerTool, _
                Optional ByVal icon As IPictureDisp)
chartObj.SetPointerModeTool toolPointerStyle, icon
End Sub

Public Sub ShowGrid()
chartObj.ShowGrid
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function chartObj() As Chart
Set chartObj = mChartRef.Target
End Function

