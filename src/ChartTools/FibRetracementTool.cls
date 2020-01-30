VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FibRetracementTool"
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

Public Type FibLineSpecifier
    Percentage          As Single
    Style               As LineStyle
End Type

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "FibRetracementTool"

'@================================================================================
' Member variables
'@================================================================================

Private mAbstractTool                       As AbstractChartTool

Private mRegion                             As ChartRegion

Private mLineSpecs()                        As FibLineSpecifier

Private mLineSeries                         As LineSeries

Private WithEvents mLine1                   As ChartSkil27.Line
Attribute mLine1.VB_VarHelpID = -1
Private WithEvents mLine2                   As ChartSkil27.Line
Attribute mLine2.VB_VarHelpID = -1
Private WithEvents mLine3                   As ChartSkil27.Line
Attribute mLine3.VB_VarHelpID = -1
Private WithEvents mLine4                   As ChartSkil27.Line
Attribute mLine4.VB_VarHelpID = -1
Private WithEvents mLine5                   As ChartSkil27.Line
Attribute mLine5.VB_VarHelpID = -1
Private WithEvents mLine6                   As ChartSkil27.Line
Attribute mLine6.VB_VarHelpID = -1
Private WithEvents mLine7                   As ChartSkil27.Line
Attribute mLine7.VB_VarHelpID = -1
Private WithEvents mLine8                   As ChartSkil27.Line
Attribute mLine8.VB_VarHelpID = -1

Private mTextSeries                         As TextSeries

Private mControlPoint1                      As Point
Private mControlPoint2                      As Point

Private mLineStyle                          As LineStyle
Private mTextStyle                          As TextStyle

Private mStartPoint1                        As ChartSkil27.Point
Private mStartPoint2                        As ChartSkil27.Point

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

If mLineSeries Is Nothing Then
    
    Set mControlPoint1 = NewPoint(X, Y)
    Set mControlPoint2 = NewPoint(X, Y)
    
    Set mLineSeries = mRegion.AddGraphicObjectSeries(New LineSeries, mLayer)
    If mLayer = -1 Then mLayer = mLineSeries.Layer  ' get the layer allocated
    mLineSeries.Extended = True
    mLineSeries.IncludeInAutoscale = False
    
    If mTextStyle Is Nothing Then
        Set mTextStyle = New TextStyle
        
        Dim aFont As New StdFont
        aFont.Name = "Lucida Console"
        aFont.Size = 8
        
        mTextStyle.Font = aFont
        mTextStyle.Box = False
        mTextStyle.Align = AlignBoxCentreLeft
    End If
    Set mTextSeries = mRegion.AddGraphicObjectSeries(New TextSeries, mLayer)
    mTextSeries.Style = mTextStyle
    
    Set mLine1 = addRetracementLine(0, X, Y)
    
    If UBound(mLineSpecs) >= 1 Then Set mLine2 = addRetracementLine(1, X, Y)
    
    If UBound(mLineSpecs) >= 2 Then Set mLine3 = addRetracementLine(2, X, Y)
    
    If UBound(mLineSpecs) >= 3 Then Set mLine4 = addRetracementLine(3, X, Y)
    
    If UBound(mLineSpecs) >= 4 Then Set mLine5 = addRetracementLine(4, X, Y)
    
    If UBound(mLineSpecs) >= 5 Then Set mLine6 = addRetracementLine(5, X, Y)
    
    If UBound(mLineSpecs) >= 6 Then Set mLine7 = addRetracementLine(6, X, Y)
    
    If UBound(mLineSpecs) >= 7 Then Set mLine8 = addRetracementLine(7, X, Y)
    
Else
    Dim ln As ChartSkil27.Line
    For Each ln In mLineSeries
        ln.Layer = mLayer
    Next
    Dim txt As ChartSkil27.Text
    For Each txt In mTextSeries
        txt.Layer = mLayer
    Next
    Set mStartPoint1 = Nothing
    Set mStartPoint2 = Nothing
    mAbstractTool.Finish
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub IChartTool_Create(ByVal X As Double, ByVal Y As Double)
Const ProcName As String = "IChartTool_Create"
On Error GoTo Err

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub IChartTool_DblClick()
Const ProcName As String = "IChartTool_DblClick"
On Error GoTo Err

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub IChartTool_Delete()
Const ProcName As String = "IChartTool_Delete"
On Error GoTo Err

mRegion.RemoveGraphicObjectSeries mLineSeries
mRegion.RemoveGraphicObjectSeries mTextSeries

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub IChartTool_Deselected()
Const ProcName As String = "IChartTool_Deselected"
On Error GoTo Err

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub IChartTool_EndResize()
Const ProcName As String = "IChartTool_EndResize"
On Error GoTo Err

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub IChartTool_Move(ByVal X As Double, ByVal Y As Double)
Const ProcName As String = "IChartTool_Move"
On Error GoTo Err

Set mControlPoint1 = NewPoint(mControlPoint1.X + X - mPrevX, mControlPoint1.Y + Y - mPrevY)
Set mControlPoint2 = NewPoint(mControlPoint2.X + X - mPrevX, mControlPoint2.Y + Y - mPrevY)

Dim ln As ChartSkil27.Line
For Each ln In mLineSeries
    ln.Point1 = NewPoint(ln.Point1.X + X - mPrevX, ln.Point1.Y + Y - mPrevY)
    ln.Point2 = NewPoint(ln.Point2.X + X - mPrevX, ln.Point2.Y + Y - mPrevY)
Next
Dim i As Long
For i = 1 To mTextSeries.Count
    Dim txt As ChartSkil27.Text: Set txt = mTextSeries(i)
    txt.Position = NewPoint(txt.Position.X + X - mPrevX, txt.Position.Y + Y - mPrevY)
    SetText txt, i
Next
mPrevX = X
mPrevY = Y

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub IChartTool_MoveResize(ByVal selHandleIndex As Long, ByVal X As Double, ByVal Y As Double)
Const ProcName As String = "IChartTool_MoveResize"
On Error GoTo Err

If selHandleIndex = 1 Then
    Set mControlPoint1 = NewPoint(X, Y)
ElseIf selHandleIndex = 2 Then
    Set mControlPoint1 = NewPoint(mControlPoint1.X, Y)
ElseIf selHandleIndex = 3 Then
    Set mControlPoint1 = NewPoint(mControlPoint1.X, Y)
    Set mControlPoint2 = NewPoint(X, mControlPoint2.Y)
ElseIf selHandleIndex = 4 Then
    Set mControlPoint1 = NewPoint(X, mControlPoint1.Y)
ElseIf selHandleIndex = 5 Then
    Set mControlPoint2 = NewPoint(X, mControlPoint2.Y)
ElseIf selHandleIndex = 6 Then
    Set mControlPoint1 = NewPoint(X, mControlPoint1.Y)
    Set mControlPoint2 = NewPoint(mControlPoint2.X, Y)
ElseIf selHandleIndex = 7 Then
    Set mControlPoint2 = NewPoint(mControlPoint2.X, Y)
Else
    Set mControlPoint2 = NewPoint(X, Y)
End If
drawLines
moveSelectionhandles

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Property Let IChartTool_Region(ByVal RHS As ChartSkil27.ChartRegion)
Const ProcName As String = "IChartTool_Region"
On Error GoTo Err

Set mRegion = RHS
Dim i As Long
For i = 0 To UBound(mLineSpecs)
    If mLineSpecs(i).Style Is Nothing Then
        Set mLineSpecs(i).Style = New LineStyle
    End If
Next

Exit Property

Err:
gNotifyUnhandledError ProcName, ModuleName
End Property

Private Sub IChartTool_ResizeCreate(ByVal X As Double, ByVal Y As Double)
Const ProcName As String = "IChartTool_ResizeCreate"
On Error GoTo Err

If mLineSeries Is Nothing Then Exit Sub

Set mControlPoint2 = NewPoint(X, Y)

Dim i As Long
For i = 1 To mLineSeries.Count
    Dim ln As ChartSkil27.Line: Set ln = mLineSeries(i)
    Dim retr As Double: retr = mControlPoint2.Y - _
                                (mControlPoint2.Y - mControlPoint1.Y) * _
                                mLineSpecs(i - 1).Percentage / 100
    ln.Point1 = NewPoint(mControlPoint1.X, retr)
    ln.Point2 = NewPoint(X, retr)
    
    Dim txt As Text: Set txt = mTextSeries(i)
    txt.Position = ln.Point2
    SetText txt, i
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName

End Sub

Private Sub IChartTool_StartMove(ByVal X As Single, ByVal Y As Single)
Const ProcName As String = "IChartTool_StartMove"
On Error GoTo Err

mPrevX = X
mPrevY = Y

startMove

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub IChartTool_StartResize(ByVal selHandleIndex As Long, ByVal X As Single, ByVal Y As Single)
Const ProcName As String = "IChartTool_StartResize"
On Error GoTo Err

startMove

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub IChartTool_Undo()
Const ProcName As String = "IChartTool_Undo"
On Error GoTo Err

Set mControlPoint1 = mStartPoint1
Set mControlPoint2 = mStartPoint2

Dim i As Long
For i = 1 To mLineSeries.Count
    Dim ln As ChartSkil27.Line: Set ln = mLineSeries(i)
    Dim retr As Double: retr = mControlPoint2.Y - _
                                (mControlPoint2.Y - mControlPoint1.Y) * _
                                mLineSpecs(i - 1).Percentage / 100
    ln.Point1 = NewPoint(mControlPoint1.X, retr)
    ln.Point2 = NewPoint(mControlPoint2.X, retr)
    ln.Layer = mLayer
    Dim txt As Text: Set txt = mTextSeries(i)
    txt.Position = ln.Point2
    SetText txt, i
    txt.Layer = mLayer
Next

Set mStartPoint1 = Nothing
Set mStartPoint2 = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mLine Event Handlers
'@================================================================================

Private Sub mLine1_SelectionStateChanged()
Const ProcName As String = "mLine1_SelectionStateChanged"
On Error GoTo Err

If mLine1.IsSelected Then notifySelected

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mLine2_SelectionStateChanged()
Const ProcName As String = "mLine2_SelectionStateChanged"
On Error GoTo Err

If mLine2.IsSelected Then notifySelected

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mLine3_SelectionStateChanged()
Const ProcName As String = "mLine3_SelectionStateChanged"
On Error GoTo Err

If mLine3.IsSelected Then notifySelected

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mLine4_SelectionStateChanged()
Const ProcName As String = "mLine4_SelectionStateChanged"
On Error GoTo Err

If mLine4.IsSelected Then notifySelected

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mLine5_SelectionStateChanged()
Const ProcName As String = "mLine5_SelectionStateChanged"
On Error GoTo Err

If mLine5.IsSelected Then notifySelected

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mLine6_SelectionStateChanged()
Const ProcName As String = "mLine6_SelectionStateChanged"
On Error GoTo Err

If mLine6.IsSelected Then notifySelected

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mLine7_SelectionStateChanged()
Const ProcName As String = "mLine7_SelectionStateChanged"
On Error GoTo Err

If mLine7.IsSelected Then notifySelected

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mLine8_SelectionStateChanged()
Const ProcName As String = "mLine8_SelectionStateChanged"
On Error GoTo Err

If mLine8.IsSelected Then notifySelected

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pChart As ChartController, _
                ByRef pLineSpecs() As FibLineSpecifier, _
                ByVal Layer As Long, _
                ByVal pTextStyle As TextStyle)
Const ProcName As String = "Initialise"
On Error GoTo Err

mLineSpecs = pLineSpecs
mLayer = Layer
Set mTextStyle = pTextStyle
Set mAbstractTool = CreateAbstractChartTool(pChart, Me)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addRetracementLine( _
                ByVal index As Long, _
                ByVal X As Double, _
                ByVal Y As Double)
Const ProcName As String = "addRetracementLine"
On Error GoTo Err

Dim ln As ChartSkil27.Line: Set ln = mLineSeries.Add(, mLineSpecs(index).Style)
ln.IsSelectable = True
ln.Layer = LayerNumbers.LayerMax
ln.SetPosition NewPoint(X, Y), NewPoint(X, Y)
mAbstractTool.AddSelectableObject ln

Dim txt As ChartSkil27.Text: Set txt = mTextSeries.Add
txt.Color = mLineSpecs(index).Style.Color
txt.Position = NewPoint(X, Y)
txt.Offset = NewSize(0.2, 0)
txt.Text = mLineSpecs(index).Percentage & "%"

Set addRetracementLine = ln

Exit Function

Err:
gNotifyUnhandledError ProcName, ModuleName
End Function

Private Sub createSelectionhandles()
Const ProcName As String = "createSelectionhandles"
On Error GoTo Err

mAbstractTool.AddSelectionHandle mControlPoint1.X, mControlPoint1.Y, True, True
mAbstractTool.AddSelectionHandle (mControlPoint1.X + mControlPoint2.X) / 2, mControlPoint1.Y, False, True
mAbstractTool.AddSelectionHandle mControlPoint2.X, mControlPoint1.Y, True, True

mAbstractTool.AddSelectionHandle mControlPoint1.X, (mControlPoint1.Y + mControlPoint2.Y) / 2, True, False
mAbstractTool.AddSelectionHandle mControlPoint2.X, (mControlPoint1.Y + mControlPoint2.Y) / 2, True, False

mAbstractTool.AddSelectionHandle mControlPoint1.X, mControlPoint2.Y, True, True
mAbstractTool.AddSelectionHandle (mControlPoint1.X + mControlPoint2.X) / 2, mControlPoint2.Y, False, True
mAbstractTool.AddSelectionHandle mControlPoint2.X, mControlPoint2.Y, True, True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub drawLines()
Const ProcName As String = "drawLines"
On Error GoTo Err

Dim i As Long
For i = 1 To mLineSeries.Count
    Dim ln As ChartSkil27.Line: Set ln = mLineSeries(i)
    Dim retr As Double: retr = mControlPoint2.Y - _
                                (mControlPoint2.Y - mControlPoint1.Y) * _
                                mLineSpecs(i - 1).Percentage / 100
    ln.Point1 = NewPoint(mControlPoint1.X, retr)
    ln.Point2 = NewPoint(mControlPoint2.X, retr)
    Dim txt As ChartSkil27.Text: Set txt = mTextSeries(i)
    txt.Position = ln.Point2
    SetText txt, i
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub moveSelectionhandles()
Const ProcName As String = "moveSelectionhandles"
On Error GoTo Err

mAbstractTool.MoveSelectionHandle 1, mControlPoint1.X, mControlPoint1.Y
mAbstractTool.MoveSelectionHandle 2, (mControlPoint1.X + mControlPoint2.X) / 2, mControlPoint1.Y
mAbstractTool.MoveSelectionHandle 3, mControlPoint2.X, mControlPoint1.Y

mAbstractTool.MoveSelectionHandle 4, mControlPoint1.X, (mControlPoint1.Y + mControlPoint2.Y) / 2
mAbstractTool.MoveSelectionHandle 5, mControlPoint2.X, (mControlPoint1.Y + mControlPoint2.Y) / 2

mAbstractTool.MoveSelectionHandle 6, mControlPoint1.X, mControlPoint2.Y
mAbstractTool.MoveSelectionHandle 7, (mControlPoint1.X + mControlPoint2.X) / 2, mControlPoint2.Y
mAbstractTool.MoveSelectionHandle 8, mControlPoint2.X, mControlPoint2.Y

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub notifySelected()
Const ProcName As String = "notifySelected"
On Error GoTo Err

If Not mAbstractTool.IsSelected Then
    If mAbstractTool.SetSelected Then
        createSelectionhandles
    Else
        mRegion.ClearSelection
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub SetText( _
                ByVal txt As Text, _
                ByVal index As Long)
Const ProcName As String = "SetText"
On Error GoTo Err

txt.Text = mLineSpecs(index - 1).Percentage & "% (" & mRegion.FormatYValue(txt.Position.Y) & ")"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub startMove()
Const ProcName As String = "startMove"
On Error GoTo Err

Set mStartPoint1 = mControlPoint1
Set mStartPoint2 = mControlPoint2

Dim ln As ChartSkil27.Line
For Each ln In mLineSeries
    ln.Layer = LayerNumbers.LayerMax
Next
Dim txt As ChartSkil27.Text
For Each txt In mTextSeries
    txt.Layer = LayerNumbers.LayerMax
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub
