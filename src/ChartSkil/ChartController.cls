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
Dim failpoint As String
On Error GoTo Err

Autoscrolling = chartObj.Autoscrolling

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Autoscrolling(ByVal Value As Boolean)
Const ProcName As String = "Autoscrolling"
Dim failpoint As String
On Error GoTo Err

chartObj.Autoscrolling = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let PeriodLength( _
                ByVal Value As TimePeriod)
Const ProcName As String = "PeriodLength"
Dim failpoint As String
On Error GoTo Err

chartObj.PeriodLength = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PeriodLength() As TimePeriod
Const ProcName As String = "PeriodLength"
Dim failpoint As String
On Error GoTo Err

Set PeriodLength = chartObj.PeriodLength

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let Chart( _
                ByVal Value As Chart)
Const ProcName As String = "Chart"
Dim failpoint As String
On Error GoTo Err

Set mChartRef = CreateWeakReference(Value)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ChartBackColor() As OLE_COLOR
Const ProcName As String = "ChartBackColor"
Dim failpoint As String
On Error GoTo Err

ChartBackColor = chartObj.ChartBackColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ChartBackColor(ByVal val As OLE_COLOR)
Const ProcName As String = "ChartBackColor"
Dim failpoint As String
On Error GoTo Err

chartObj.ChartBackColor = val

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ChartLeft() As Single
Const ProcName As String = "ChartLeft"
Dim failpoint As String
On Error GoTo Err

ChartLeft = chartObj.ChartLeft

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ChartWidth() As Single
Const ProcName As String = "ChartWidth"
Dim failpoint As String
On Error GoTo Err

ChartWidth = chartObj.ChartWidth

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CurrentPeriodNumber() As Long
Const ProcName As String = "CurrentPeriodNumber"
Dim failpoint As String
On Error GoTo Err

CurrentPeriodNumber = chartObj.CurrentPeriodNumber

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CurrentSessionEndTime() As Date
Const ProcName As String = "CurrentSessionEndTime"
Dim failpoint As String
On Error GoTo Err

CurrentSessionEndTime = chartObj.CurrentSessionEndTime

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CurrentSessionStartTime() As Date
Const ProcName As String = "CurrentSessionStartTime"
Dim failpoint As String
On Error GoTo Err

CurrentSessionStartTime = chartObj.CurrentSessionStartTime

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get FirstVisiblePeriod() As Long
Const ProcName As String = "FirstVisiblePeriod"
Dim failpoint As String
On Error GoTo Err

FirstVisiblePeriod = chartObj.FirstVisiblePeriod

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let FirstVisiblePeriod(ByVal Value As Long)
Const ProcName As String = "FirstVisiblePeriod"
Dim failpoint As String
On Error GoTo Err

chartObj.FirstVisiblePeriod = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HorizontalMouseScrollingAllowed() As Boolean
Const ProcName As String = "HorizontalMouseScrollingAllowed"
Dim failpoint As String
On Error GoTo Err

HorizontalMouseScrollingAllowed = chartObj.HorizontalMouseScrollingAllowed

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HorizontalMouseScrollingAllowed(ByVal Value As Boolean)
Const ProcName As String = "HorizontalMouseScrollingAllowed"
Dim failpoint As String
On Error GoTo Err

chartObj.HorizontalMouseScrollingAllowed = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HorizontalScrollBarVisible() As Boolean
Const ProcName As String = "HorizontalScrollBarVisible"
Dim failpoint As String
On Error GoTo Err

HorizontalScrollBarVisible = chartObj.HorizontalScrollBarVisible

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HorizontalScrollBarVisible(ByVal val As Boolean)
Const ProcName As String = "HorizontalScrollBarVisible"
Dim failpoint As String
On Error GoTo Err

chartObj.HorizontalScrollBarVisible = val

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsDrawingEnabled() As Boolean
Const ProcName As String = "IsDrawingEnabled"
Dim failpoint As String
On Error GoTo Err

IsDrawingEnabled = chartObj.IsDrawingEnabled

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsGridHidden() As Boolean
Const ProcName As String = "IsGridHidden"
Dim failpoint As String
On Error GoTo Err

IsGridHidden = chartObj.IsGridHidden

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LastVisiblePeriod() As Long
Const ProcName As String = "LastVisiblePeriod"
Dim failpoint As String
On Error GoTo Err

LastVisiblePeriod = chartObj.LastVisiblePeriod

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let LastVisiblePeriod(ByVal Value As Long)
Const ProcName As String = "LastVisiblePeriod"
Dim failpoint As String
On Error GoTo Err

chartObj.LastVisiblePeriod = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Periods() As Periods
Const ProcName As String = "Periods"
Dim failpoint As String
On Error GoTo Err

Set Periods = chartObj.Periods

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PointerCrosshairsColor() As Long
Const ProcName As String = "PointerCrosshairsColor"
Dim failpoint As String
On Error GoTo Err

PointerCrosshairsColor = chartObj.PointerCrosshairsColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let PointerCrosshairsColor(ByVal Value As Long)
Const ProcName As String = "PointerCrosshairsColor"
Dim failpoint As String
On Error GoTo Err

chartObj.PointerCrosshairsColor = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PointerDiscColor() As Long
Const ProcName As String = "PointerDiscColor"
Dim failpoint As String
On Error GoTo Err

PointerDiscColor = chartObj.PointerDiscColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let PointerDiscColor(ByVal Value As Long)
Const ProcName As String = "PointerDiscColor"
Dim failpoint As String
On Error GoTo Err

chartObj.PointerDiscColor = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PointerIcon() As IPictureDisp
Const ProcName As String = "PointerIcon"
Dim failpoint As String
On Error GoTo Err

Set PointerIcon = chartObj.PointerIcon

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let PointerIcon(ByVal Value As IPictureDisp)
Const ProcName As String = "PointerIcon"
Dim failpoint As String
On Error GoTo Err

chartObj.PointerIcon = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PointerMode() As PointerModes
Const ProcName As String = "PointerMode"
Dim failpoint As String
On Error GoTo Err

PointerMode = chartObj.PointerMode

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PointerStyle() As PointerStyles
Const ProcName As String = "PointerStyle"
Dim failpoint As String
On Error GoTo Err

PointerStyle = chartObj.PointerStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let PointerStyle(ByVal Value As PointerStyles)
Const ProcName As String = "PointerStyle"
Dim failpoint As String
On Error GoTo Err

chartObj.PointerStyle = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Regions() As ChartRegions
Const ProcName As String = "Regions"
Dim failpoint As String
On Error GoTo Err

Set Regions = chartObj.Regions

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get SessionEndTime() As Date
Const ProcName As String = "SessionEndTime"
Dim failpoint As String
On Error GoTo Err

SessionEndTime = chartObj.SessionEndTime

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let SessionEndTime(ByVal val As Date)
Const ProcName As String = "SessionEndTime"
Dim failpoint As String
On Error GoTo Err

chartObj.SessionEndTime = val

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get SessionStartTime() As Date
Const ProcName As String = "SessionStartTime"
Dim failpoint As String
On Error GoTo Err

SessionStartTime = chartObj.SessionStartTime

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let SessionStartTime(ByVal val As Date)
Const ProcName As String = "SessionStartTime"
Dim failpoint As String
On Error GoTo Err

chartObj.SessionStartTime = val

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get TwipsPerBar() As Long
Const ProcName As String = "TwipsPerBar"
Dim failpoint As String
On Error GoTo Err

TwipsPerBar = chartObj.TwipsPerBar

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let TwipsPerBar(ByVal val As Long)
Const ProcName As String = "TwipsPerBar"
Dim failpoint As String
On Error GoTo Err

chartObj.TwipsPerBar = val

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let VerticalGridTimePeriod( _
                ByVal Value As TimePeriod)
Const ProcName As String = "VerticalGridTimePeriod"
Dim failpoint As String
On Error GoTo Err

chartObj.VerticalGridTimePeriod = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get VerticalGridTimePeriod() As TimePeriod
Const ProcName As String = "VerticalGridTimePeriod"
Dim failpoint As String
On Error GoTo Err

Set VerticalGridTimePeriod = chartObj.VerticalGridTimePeriod

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get VerticalMouseScrollingAllowed() As Boolean
Const ProcName As String = "VerticalMouseScrollingAllowed"
Dim failpoint As String
On Error GoTo Err

VerticalMouseScrollingAllowed = chartObj.VerticalMouseScrollingAllowed

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let VerticalMouseScrollingAllowed(ByVal Value As Boolean)
Const ProcName As String = "VerticalMouseScrollingAllowed"
Dim failpoint As String
On Error GoTo Err

chartObj.VerticalMouseScrollingAllowed = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get XAxisRegion() As ChartRegion
Const ProcName As String = "XAxisRegion"
Dim failpoint As String
On Error GoTo Err

Set XAxisRegion = chartObj.XAxisRegion

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get XAxisVisible() As Boolean
Const ProcName As String = "XAxisVisible"
Dim failpoint As String
On Error GoTo Err

XAxisVisible = chartObj.XAxisVisible

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let XAxisVisible(ByVal Value As Boolean)
Const ProcName As String = "XAxisVisible"
Dim failpoint As String
On Error GoTo Err

chartObj.XAxisVisible = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let XCursorTextStyle(ByVal Value As TextStyle)
Const ProcName As String = "XCursorTextStyle"
Dim failpoint As String
On Error GoTo Err

chartObj.XCursorTextStyle = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get XCursorTextStyle() As TextStyle
Const ProcName As String = "XCursorTextStyle"
Dim failpoint As String
On Error GoTo Err

Set XCursorTextStyle = chartObj.XCursorTextStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YAxisPosition() As Long
Const ProcName As String = "YAxisPosition"
Dim failpoint As String
On Error GoTo Err

YAxisPosition = chartObj.YAxisPosition

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YAxisVisible() As Boolean
Const ProcName As String = "YAxisVisible"
Dim failpoint As String
On Error GoTo Err

YAxisVisible = chartObj.YAxisVisible

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let YAxisVisible(ByVal Value As Boolean)
Const ProcName As String = "YAxisVisible"
Dim failpoint As String
On Error GoTo Err

chartObj.YAxisVisible = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YAxisWidthCm() As Single
Const ProcName As String = "YAxisWidthCm"
Dim failpoint As String
On Error GoTo Err

YAxisWidthCm = chartObj.YAxisWidthCm

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let YAxisWidthCm(ByVal Value As Single)
Const ProcName As String = "YAxisWidthCm"
Dim failpoint As String
On Error GoTo Err

chartObj.YAxisWidthCm = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Public Sub ClearChart()
Const ProcName As String = "ClearChart"
Dim failpoint As String
On Error GoTo Err

chartObj.ClearChart

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub DisableDrawing()
Const ProcName As String = "DisableDrawing"
Dim failpoint As String
On Error GoTo Err

chartObj.DisableDrawing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub EnableDrawing()
Const ProcName As String = "EnableDrawing"
Dim failpoint As String
On Error GoTo Err

chartObj.EnableDrawing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
Dim failpoint As String
On Error GoTo Err

chartObj.Finish

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Finished()
Const ProcName As String = "Finished"
Dim failpoint As String
On Error GoTo Err

Set mChartRef = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
Dim failpoint As String
On Error GoTo Err

GetXFromTimestamp = chartObj.GetXFromTimestamp(Timestamp, forceNewPeriod, duplicateNumber)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub HideGrid()
Const ProcName As String = "HideGrid"
Dim failpoint As String
On Error GoTo Err

chartObj.HideGrid

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function IsTimeInSession(ByVal Timestamp As Date) As Boolean
Const ProcName As String = "IsTimeInSession"
Dim failpoint As String
On Error GoTo Err

IsTimeInSession = chartObj.IsTimeInSession(Timestamp)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub ScrollX(ByVal Value As Long)
Const ProcName As String = "ScrollX"
Dim failpoint As String
On Error GoTo Err

chartObj.ScrollX Value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub SetPointerModeDefault()
Const ProcName As String = "SetPointerModeDefault"
Dim failpoint As String
On Error GoTo Err

chartObj.SetPointerModeDefault

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub SetPointerModeSelection()
Const ProcName As String = "SetPointerModeSelection"
Dim failpoint As String
On Error GoTo Err

chartObj.SetPointerModeSelection

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub SetPointerModeTool( _
                Optional ByVal toolPointerStyle As PointerStyles = PointerTool, _
                Optional ByVal icon As IPictureDisp)
Const ProcName As String = "SetPointerModeTool"
Dim failpoint As String
On Error GoTo Err

chartObj.SetPointerModeTool toolPointerStyle, icon

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ShowGrid()
Const ProcName As String = "ShowGrid"
Dim failpoint As String
On Error GoTo Err

chartObj.ShowGrid

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function chartObj() As Chart
Const ProcName As String = "chartObj"
Dim failpoint As String
On Error GoTo Err

Set chartObj = mChartRef.Target

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

