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
If mLine Is Nothing Then
    Set mLineSeries = mRegion.AddGraphicObjectSeries(New LineSeries, mLayer)
    If mLayer = -1 Then mLayer = mLineSeries.Layer  ' get the layer allocated
    mLineSeries.Extended = True
    mLineSeries.IncludeInAutoscale = False
    
    Set mLine = mLineSeries.Add(, mLineStyle)
    mLine.IsSelectable = True
    
    mLine.Layer = LayerNumbers.LayerMax
    mLine.SetPosition mRegion.NewPoint(X, Y), mRegion.NewPoint(X, Y)
    mAbstractTool.AddSelectableObject mLine
Else
    mLine.Layer = mLayer
    Set mStartPoint1 = Nothing
    Set mStartPoint2 = Nothing
    mAbstractTool.Finish
End If
End Sub

Private Sub IChartTool_Create(ByVal X As Double, ByVal Y As Double)

End Sub

Private Sub IChartTool_DblClick()

End Sub

Private Sub IChartTool_Delete()
mRegion.RemoveGraphicObject mLine
mRegion.RemoveGraphicObjectSeries mLineSeries
End Sub

Private Sub IChartTool_Deselected()

End Sub

Private Sub IChartTool_EndResize()

End Sub

Private Sub IChartTool_Move(ByVal X As Double, ByVal Y As Double)
mLine.Point1 = mRegion.NewPoint(mLine.Point1.X + X - mPrevX, mLine.Point1.Y + Y - mPrevY)
mLine.Point2 = mRegion.NewPoint(mLine.Point2.X + X - mPrevX, mLine.Point2.Y + Y - mPrevY)
mPrevX = X
mPrevY = Y
End Sub

Private Sub IChartTool_MoveResize(ByVal selHandleIndex As Long, ByVal X As Double, ByVal Y As Double)
If selHandleIndex = 1 Then
    mLine.Point1 = mRegion.NewPoint(X, Y)
Else
    mLine.Point2 = mRegion.NewPoint(X, Y)
End If

End Sub

Private Property Let IChartTool_Region(ByVal RHS As ChartSkil26.ChartRegion)
Set mRegion = RHS
If mLineStyle Is Nothing Then Set mLineStyle = New LineStyle
End Property

Private Sub IChartTool_ResizeCreate(ByVal X As Double, ByVal Y As Double)
If Not mLine Is Nothing Then mLine.Point2 = mRegion.NewPoint(X, Y)
End Sub

Private Sub IChartTool_StartMove(ByVal X As Single, ByVal Y As Single)
mPrevX = X
mPrevY = Y

Set mStartPoint1 = mLine.Point1
Set mStartPoint2 = mLine.Point2

mLine.Layer = LayerNumbers.LayerMax
End Sub

Private Sub IChartTool_StartResize(ByVal selHandleIndex As Long, ByVal X As Single, ByVal Y As Single)
If mLine Is Nothing Then Exit Sub

Set mStartPoint1 = mLine.Point1
Set mStartPoint2 = mLine.Point2

mLine.Layer = LayerNumbers.LayerMax
End Sub

Private Sub IChartTool_Undo()
mLine.SetPosition mStartPoint1, mStartPoint2
mLine.Layer = mLayer
Set mStartPoint1 = Nothing
Set mStartPoint2 = Nothing
End Sub

'@================================================================================
' mLine Event Handlers
'@================================================================================

Private Sub mLine_SelectionStateChanged()
Dim failpoint As Long
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
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "mLine_SelectionStateChanged" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Region() As ChartRegion
Set Region = mRegion
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pChart As ChartController, _
                ByVal pLineStyle As LineStyle, _
                ByVal Layer As Long)
Set mLineStyle = pLineStyle
mLayer = Layer
Set mAbstractTool = CreateAbstractChartTool(pChart, Me)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

