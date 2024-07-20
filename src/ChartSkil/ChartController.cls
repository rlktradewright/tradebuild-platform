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

Private Const ModuleName                As String = "ChartController"

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
Const ProcName As String = "Autoscrolling"

On Error GoTo Err

Autoscrolling = chartObj.Autoscrolling

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Autoscrolling(ByVal Value As Boolean)
Const ProcName As String = "Autoscrolling"

On Error GoTo Err

chartObj.Autoscrolling = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Chart( _
                ByVal Value As Chart)
Const ProcName As String = "Chart"

On Error GoTo Err

Set mChartRef = CreateWeakReference(Value)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ChartBackColor() As OLE_COLOR
Const ProcName As String = "ChartBackColor"

On Error GoTo Err

ChartBackColor = chartObj.ChartBackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ChartBackColor(ByVal val As OLE_COLOR)
Const ProcName As String = "ChartBackColor"

On Error GoTo Err

chartObj.ChartBackColor = val

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ChartID() As String
ChartID = chartObj.ChartID
End Property

Public Property Get ChartLeft() As Single
Const ProcName As String = "ChartLeft"

On Error GoTo Err

ChartLeft = chartObj.ChartLeft

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ChartWidth() As Single
Const ProcName As String = "ChartWidth"

On Error GoTo Err

ChartWidth = chartObj.ChartWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CrosshairLineStyle() As LineStyle
Const ProcName As String = "CrosshairLineStyle"
On Error GoTo Err

Set CrosshairLineStyle = chartObj.CrosshairLineStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let CrosshairLineStyle(ByVal Value As LineStyle)
Const ProcName As String = "CrosshairLineStyle"

On Error GoTo Err

chartObj.CrosshairLineStyle = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentPeriodNumber() As Long
Const ProcName As String = "CurrentPeriodNumber"

On Error GoTo Err

CurrentPeriodNumber = chartObj.CurrentPeriodNumber

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentSessionEndTime() As Date
Const ProcName As String = "CurrentSessionEndTime"

On Error GoTo Err

CurrentSessionEndTime = chartObj.CurrentSessionEndTime

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentSessionStartTime() As Date
Const ProcName As String = "CurrentSessionStartTime"

On Error GoTo Err

CurrentSessionStartTime = chartObj.CurrentSessionStartTime

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultRegionStyle() As ChartRegionStyle
Const ProcName As String = "DefaultRegionStyle"
On Error GoTo Err

Set DefaultRegionStyle = chartObj.DefaultRegionStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DefaultRegionStyle(ByVal Value As ChartRegionStyle)
Const ProcName As String = "DefaultRegionStyle"

On Error GoTo Err

chartObj.DefaultRegionStyle = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultYAxisRegionStyle() As ChartRegionStyle
Const ProcName As String = "DefaultYAxisRegionStyle"
On Error GoTo Err

Set DefaultYAxisRegionStyle = chartObj.DefaultYAxisRegionStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DefaultYAxisRegionStyle(ByVal Value As ChartRegionStyle)
Const ProcName As String = "DefaultYAxisRegionStyle"

On Error GoTo Err

chartObj.DefaultYAxisRegionStyle = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FirstVisiblePeriod() As Long
Const ProcName As String = "FirstVisiblePeriod"

On Error GoTo Err

FirstVisiblePeriod = chartObj.FirstVisiblePeriod

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let FirstVisiblePeriod(ByVal Value As Long)
Const ProcName As String = "FirstVisiblePeriod"

On Error GoTo Err

chartObj.FirstVisiblePeriod = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HorizontalMouseScrollingAllowed() As Boolean
Const ProcName As String = "HorizontalMouseScrollingAllowed"

On Error GoTo Err

HorizontalMouseScrollingAllowed = chartObj.HorizontalMouseScrollingAllowed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HorizontalMouseScrollingAllowed(ByVal Value As Boolean)
Const ProcName As String = "HorizontalMouseScrollingAllowed"

On Error GoTo Err

chartObj.HorizontalMouseScrollingAllowed = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HorizontalScrollBarVisible() As Boolean
Const ProcName As String = "HorizontalScrollBarVisible"

On Error GoTo Err

HorizontalScrollBarVisible = chartObj.HorizontalScrollBarVisible

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HorizontalScrollBarVisible(ByVal val As Boolean)
Const ProcName As String = "HorizontalScrollBarVisible"

On Error GoTo Err

chartObj.HorizontalScrollBarVisible = val

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsDrawingEnabled() As Boolean
Const ProcName As String = "IsDrawingEnabled"

On Error GoTo Err

IsDrawingEnabled = chartObj.IsDrawingEnabled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsGridHidden() As Boolean
Const ProcName As String = "IsGridHidden"

On Error GoTo Err

IsGridHidden = chartObj.IsGridHidden

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LastVisiblePeriod() As Long
Const ProcName As String = "LastVisiblePeriod"

On Error GoTo Err

LastVisiblePeriod = chartObj.LastVisiblePeriod

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let LastVisiblePeriod(ByVal Value As Long)
Const ProcName As String = "LastVisiblePeriod"

On Error GoTo Err

chartObj.LastVisiblePeriod = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Periods() As Periods
Const ProcName As String = "Periods"

On Error GoTo Err

Set Periods = chartObj.Periods

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PeriodWidth() As Long
Const ProcName As String = "PeriodWidth"
On Error GoTo Err

PeriodWidth = chartObj.PeriodWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PeriodWidth(ByVal val As Long)
Const ProcName As String = "PeriodWidth"
On Error GoTo Err

chartObj.PeriodWidth = val

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointerCrosshairsColor() As Long
Const ProcName As String = "PointerCrosshairsColor"

On Error GoTo Err

PointerCrosshairsColor = chartObj.PointerCrosshairsColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PointerCrosshairsColor(ByVal Value As Long)
Const ProcName As String = "PointerCrosshairsColor"

On Error GoTo Err

chartObj.PointerCrosshairsColor = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointerDiscColor() As Long
Const ProcName As String = "PointerDiscColor"

On Error GoTo Err

PointerDiscColor = chartObj.PointerDiscColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PointerDiscColor(ByVal Value As Long)
Const ProcName As String = "PointerDiscColor"

On Error GoTo Err

chartObj.PointerDiscColor = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointerIcon() As IPictureDisp
Const ProcName As String = "PointerIcon"

On Error GoTo Err

Set PointerIcon = chartObj.PointerIcon

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PointerIcon(ByVal Value As IPictureDisp)
Const ProcName As String = "PointerIcon"

On Error GoTo Err

chartObj.PointerIcon = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointerMode() As PointerModes
Const ProcName As String = "PointerMode"

On Error GoTo Err

PointerMode = chartObj.PointerMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointerStyle() As PointerStyles
Const ProcName As String = "PointerStyle"

On Error GoTo Err

PointerStyle = chartObj.PointerStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PointerStyle(ByVal Value As PointerStyles)
Const ProcName As String = "PointerStyle"

On Error GoTo Err

chartObj.PointerStyle = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Regions() As ChartRegions
Const ProcName As String = "Regions"

On Error GoTo Err

Set Regions = chartObj.Regions

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SessionEndTime() As Date
Const ProcName As String = "SessionEndTime"

On Error GoTo Err

SessionEndTime = chartObj.SessionEndTime

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let SessionEndTime(ByVal val As Date)
Const ProcName As String = "SessionEndTime"

On Error GoTo Err

chartObj.SessionEndTime = val

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SessionStartTime() As Date
Const ProcName As String = "SessionStartTime"

On Error GoTo Err

SessionStartTime = chartObj.SessionStartTime

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let SessionStartTime(ByVal val As Date)
Const ProcName As String = "SessionStartTime"

On Error GoTo Err

chartObj.SessionStartTime = val

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Style(ByVal Value As ChartStyle)
Const ProcName As String = "Style"
On Error GoTo Err

chartObj.Style = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Style() As ChartStyle
Set Style = chartObj.Style
End Property

Public Property Let TimePeriod( _
                ByVal Value As TimePeriod)
Const ProcName As String = "TimePeriod"
On Error GoTo Err

chartObj.TimePeriod = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TimePeriod() As TimePeriod
Const ProcName As String = "TimePeriod"
On Error GoTo Err

Set TimePeriod = chartObj.TimePeriod

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let VerticalGridTimePeriod( _
                ByVal Value As TimePeriod)
Const ProcName As String = "VerticalGridTimePeriod"

On Error GoTo Err

chartObj.VerticalGridTimePeriod = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get VerticalGridTimePeriod() As TimePeriod
Const ProcName As String = "VerticalGridTimePeriod"

On Error GoTo Err

Set VerticalGridTimePeriod = chartObj.VerticalGridTimePeriod

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get VerticalMouseScrollingAllowed() As Boolean
Const ProcName As String = "VerticalMouseScrollingAllowed"

On Error GoTo Err

VerticalMouseScrollingAllowed = chartObj.VerticalMouseScrollingAllowed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let VerticalMouseScrollingAllowed(ByVal Value As Boolean)
Const ProcName As String = "VerticalMouseScrollingAllowed"

On Error GoTo Err

chartObj.VerticalMouseScrollingAllowed = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XAxisRegion() As ChartRegion
Const ProcName As String = "XAxisRegion"

On Error GoTo Err

Set XAxisRegion = chartObj.XAxisRegion

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XAxisRegionStyle() As ChartRegionStyle
Const ProcName As String = "XAxisRegionStyle"
On Error GoTo Err

Set XAxisRegionStyle = chartObj.XAxisRegionStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let XAxisRegionStyle(ByVal Value As ChartRegionStyle)
Const ProcName As String = "XAxisRegionStyle"

On Error GoTo Err

chartObj.XAxisRegionStyle = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get XAxisVisible() As Boolean
Const ProcName As String = "XAxisVisible"

On Error GoTo Err

XAxisVisible = chartObj.XAxisVisible

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let XAxisVisible(ByVal Value As Boolean)
Const ProcName As String = "XAxisVisible"

On Error GoTo Err

chartObj.XAxisVisible = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YAxisPosition() As Long
Const ProcName As String = "YAxisPosition"

On Error GoTo Err

YAxisPosition = chartObj.YAxisPosition

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YAxisVisible() As Boolean
Const ProcName As String = "YAxisVisible"

On Error GoTo Err

YAxisVisible = chartObj.YAxisVisible

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let YAxisVisible(ByVal Value As Boolean)
Const ProcName As String = "YAxisVisible"

On Error GoTo Err

chartObj.YAxisVisible = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YAxisWidthCm() As Single
Const ProcName As String = "YAxisWidthCm"

On Error GoTo Err

YAxisWidthCm = chartObj.YAxisWidthCm

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let YAxisWidthCm(ByVal Value As Single)
Const ProcName As String = "YAxisWidthCm"

On Error GoTo Err

chartObj.YAxisWidthCm = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Public Sub ClearChart()
Const ProcName As String = "ClearChart"

On Error GoTo Err

chartObj.ClearChart

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DisableDrawing()
Const ProcName As String = "DisableDrawing"
On Error GoTo Err

chartObj.DisableDrawing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub EnableDrawing()
Const ProcName As String = "EnableDrawing"
On Error GoTo Err

chartObj.EnableDrawing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

chartObj.Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Finished()
Const ProcName As String = "Finished"
On Error GoTo Err

Set mChartRef = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
Const ProcName As String = "GetXFromTimestamp"
On Error GoTo Err

GetXFromTimestamp = chartObj.GetXFromTimestamp(Timestamp, forceNewPeriod, duplicateNumber)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub HideGrid()
Const ProcName As String = "HideGrid"
On Error GoTo Err

chartObj.HideGrid

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function IsTimeInSession(ByVal Timestamp As Date) As Boolean
Const ProcName As String = "IsTimeInSession"
On Error GoTo Err

IsTimeInSession = chartObj.IsTimeInSession(Timestamp)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub ScrollX(ByVal Value As Long)
Const ProcName As String = "ScrollX"
On Error GoTo Err

chartObj.ScrollX Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetPointerModeDefault()
Const ProcName As String = "SetPointerModeDefault"
On Error GoTo Err

chartObj.SetPointerModeDefault

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetPointerModeSelection()
Const ProcName As String = "SetPointerModeSelection"
On Error GoTo Err

chartObj.SetPointerModeSelection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetPointerModeTool( _
                Optional ByVal toolPointerStyle As PointerStyles = PointerTool, _
                Optional ByVal icon As IPictureDisp)
Const ProcName As String = "SetPointerModeTool"
On Error GoTo Err

chartObj.SetPointerModeTool toolPointerStyle, icon

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ShowGrid()
Const ProcName As String = "ShowGrid"

On Error GoTo Err

chartObj.ShowGrid

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function chartObj() As Chart
Const ProcName As String = "chartObj"
On Error GoTo Err

Set chartObj = mChartRef.Target

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

