VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LineTool"
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

Implements IChartTool

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

Private Const ProjectName                   As String = "ChartTools"
Private Const ModuleName                    As String = "LineTool"

'@================================================================================
' Member variables
'@================================================================================

Private mAbstractTool                       As AbstractChartTool

Private mRegion                             As ChartRegion

Private mLineSeries                         As LineSeries
Private WithEvents mLine                    As ChartSkil26.Line
Attribute mLine.VB_VarHelpID = -1

Private mLineStyle                          As LineStyle

Private mStartPoint1                        As ChartSkil26.Point
Private mStartPoint2                        As ChartSkil26.Point

Private mLayer                              As LayerNumbers

Private mPrevX                              As Double
Private mPrevY                              As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IChartTool Interface Members
'@================================================================================

Private Sub IChartTool_Click(ByVal X As Double, ByVal Y As Double)
Const ProcName As String = "IChartTool_Click"
On Error GoTo Err

If mLine Is Nothing Then
    Set mLineSeries = mRegion.AddGraphicObjectSeries(New LineSeries, mLayer)
    If mLayer = -1 Then mLayer = mLineSeries.Layer  ' get the layer allocated
    mLineSeries.Extended = True
    mLineSeries.IncludeInAutoscale = False
    mLineSeries.Color = &H808080
    
    Set mLine = mLineSeries.Add(, mLineStyle)
    mLine.IsSelectable = True
    
    mLine.Layer = LayerNumbers.LayerMax
    mLine.SetPosition NewPoint(X, Y), NewPoint(X, Y)
    mAbstractTool.AddSelectableObject mLine
Else
    mLine.Layer = mLayer
    Set mStartPoint1 = Nothing
    Set mStartPoint2 = Nothing
    mAbstractTool.Finish
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IChartTool_Create(ByVal X As Double, ByVal Y As Double)

Const ProcName As String = "IChartTool_Create"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IChartTool_DblClick()

Const ProcName As String = "IChartTool_DblClick"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IChartTool_Delete()
Const ProcName As String = "IChartTool_Delete"
On Error GoTo Err

mRegion.RemoveGraphicObject mLine
mRegion.RemoveGraphicObjectSeries mLineSeries

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IChartTool_Deselected()

Const ProcName As String = "IChartTool_Deselected"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IChartTool_EndResize()

Const ProcName As String = "IChartTool_EndResize"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IChartTool_Move(ByVal X As Double, ByVal Y As Double)
Const ProcName As String = "IChartTool_Move"
On Error GoTo Err

mLine.Point1 = NewPoint(mLine.Point1.X + X - mPrevX, mLine.Point1.Y + Y - mPrevY)
mLine.Point2 = NewPoint(mLine.Point2.X + X - mPrevX, mLine.Point2.Y + Y - mPrevY)
mPrevX = X
mPrevY = Y

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IChartTool_MoveResize(ByVal selHandleIndex As Long, ByVal X As Double, ByVal Y As Double)
Const ProcName As String = "IChartTool_MoveResize"
On Error GoTo Err

If selHandleIndex = 1 Then
    mLine.Point1 = NewPoint(X, Y)
Else
    mLine.Point2 = NewPoint(X, Y)
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Property Let IChartTool_Region(ByVal RHS As ChartSkil26.ChartRegion)
Const ProcName As String = "IChartTool_Region"
On Error GoTo Err

Set mRegion = RHS
If mLineStyle Is Nothing Then Set mLineStyle = New LineStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub IChartTool_ResizeCreate(ByVal X As Double, ByVal Y As Double)
Const ProcName As String = "IChartTool_ResizeCreate"
On Error GoTo Err

If Not mLine Is Nothing Then mLine.Point2 = NewPoint(X, Y)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IChartTool_StartMove(ByVal X As Single, ByVal Y As Single)
Const ProcName As String = "IChartTool_StartMove"
On Error GoTo Err

mPrevX = X
mPrevY = Y

Set mStartPoint1 = mLine.Point1
Set mStartPoint2 = mLine.Point2

mLine.Layer = LayerNumbers.LayerMax

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IChartTool_StartResize(ByVal selHandleIndex As Long, ByVal X As Single, ByVal Y As Single)
Const ProcName As String = "IChartTool_StartResize"
On Error GoTo Err

If mLine Is Nothing Then Exit Sub

Set mStartPoint1 = mLine.Point1
Set mStartPoint2 = mLine.Point2

mLine.Layer = LayerNumbers.LayerMax

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IChartTool_Undo()
Const ProcName As String = "IChartTool_Undo"
On Error GoTo Err

mLine.SetPosition mStartPoint1, mStartPoint2
mLine.Layer = mLayer
Set mStartPoint1 = Nothing
Set mStartPoint2 = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mLine Event Handlers
'@================================================================================

Private Sub mLine_SelectionStateChanged()
Const ProcName As String = "mLine_SelectionStateChanged"
On Error GoTo Err

If mLine.IsSelected Then
    If Not mAbstractTool.IsSelected Then
        If mAbstractTool.SetSelected Then
            mAbstractTool.AddSelectionHandle mLine.Point1.X, mLine.Point1.Y
            mAbstractTool.AddSelectionHandle mLine.Point2.X, mLine.Point2.Y
        Else
            mRegion.ClearSelection
        End If
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Region() As ChartRegion
Const ProcName As String = "Region"
On Error GoTo Err

Set Region = mRegion

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pChart As ChartController, _
                ByVal pLineStyle As LineStyle, _
                ByVal Layer As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mLineStyle = pLineStyle
mLayer = Layer
Set mAbstractTool = CreateAbstractChartTool(pChart, Me)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

