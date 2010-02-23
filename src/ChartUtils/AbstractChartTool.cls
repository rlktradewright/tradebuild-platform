VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AbstractChartTool"
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

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum SelectionModes
    SelectionModeNone
    SelectionModeCreate
    SelectionModeTool
    SelectionModeSelectionHandle1
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "AbstractChartTool"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mChartEV                 As ChartController
Attribute mChartEV.VB_VarHelpID = -1
Private mChart                              As ChartController

Private mRegion                             As ChartRegion
Private WithEvents mRegionEV                As ChartRegion
Attribute mRegionEV.VB_VarHelpID = -1

Private mSelectionHandleSeries              As DataPointSeries
Private mSelectionHandles                   As Collection

Private mSelectableObjects                  As Collection

Private mSelMode                            As SelectionModes

Private mPrevX                              As Double
Private mPrevY                              As Double

Private mToolRef                            As WeakReference
Private mTool                               As IChartTool

Private mToolInControl                      As Boolean

Private mPointerIcon                        As IPictureDisp

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mChart Event Handlers
'@================================================================================

Private Sub mChartEV_KeyDown(KeyCode As Integer, Shift As Integer)

Const ProcName As String = "mChartEV_KeyDown"
Dim failpoint As String
On Error GoTo Err

Select Case KeyCode
Case vbKeyDelete
    If IsSelected Then
        mTool.Delete
        Delete
        KeyCode = 0
    End If
Case vbKeyEscape
    If Undo Then KeyCode = 0
End Select

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mChartEV_PointerModeChanged()
Const ProcName As String = "mChartEV_PointerModeChanged"
Dim failpoint As String
On Error GoTo Err

If mChart.PointerMode = PointerModeDefault Then
    Deselected
    mTool.Deselected
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mChartEV_RegionSelected(ByVal Region As ChartSkil26.ChartRegion)
Const ProcName As String = "mChartEV_RegionSelected"
Dim failpoint As String
On Error GoTo Err

If mRegion Is Nothing Then
    
    Set mTool = mToolRef.Target ' convert to a strong ref for efficiency, now that we know
                                ' that this tool object is actually being used. This will also
                                ' keep both objects alive
    Set mRegion = Region
    Set mRegionEV = Region
    mTool.Region = Region

    Set mSelectionHandleSeries = mRegion.AddGraphicObjectSeries(New DataPointSeries, LayerNumbers.LayerMax)
    mSelectionHandleSeries.IncludeInAutoscale = False
    mSelectionHandleSeries.DisplayMode = DataPointDisplayModePoint
    mSelectionHandleSeries.PointStyle = PointSquare
    mSelectionHandleSeries.LineThickness = 5
    mSelectionHandleSeries.Color = vbBlack
    
    Set mSelectionHandles = New Collection
    
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mRegionEV Event Handlers
'@================================================================================

Private Sub mRegionEV_DblCLick()
Const ProcName As String = "mRegionEV_DblCLick"
Dim failpoint As String
On Error GoTo Err

mTool.DblClick

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mRegionEV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim obj As IGraphicObject

Const ProcName As String = "mRegionEV_MouseDown"
Dim failpoint As String
On Error GoTo Err

Debug.Print "x=" & X & " y=" & Y

mPrevX = X
mPrevY = Y

Select Case mSelMode
Case SelectionModeNone
    mToolInControl = True
    mSelMode = SelectionModeCreate
    mTool.Create X, Y
Case SelectionModeCreate
    
Case SelectionModeTool
    If Not mToolInControl Then
        For Each obj In mSelectableObjects
            If obj.Selected Then Exit For
        Next
    
        If obj Is Nothing Then
            ' the user has deselected
            Deselected
            mTool.Deselected
        Else
            
            mToolInControl = True
            mChart.SetPointerModeTool
            
            StartMove
            If SelectionHandleHit(X, Y) Then
                mTool.StartResize mSelMode - SelectionModes.SelectionModeSelectionHandle1 + 1, X, Y
            Else
                mTool.StartMove X, Y
            End If
            
        End If
    End If
Case Else

End Select

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mRegionEV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim selHandle As SelectionHandle

'If Not CBool(Button And MouseButtonConstants.vbLeftButton) Then Exit Sub
Const ProcName As String = "mRegionEV_MouseMove"
Dim failpoint As String
On Error GoTo Err

If Not mToolInControl Then Exit Sub

Debug.Print "x=" & X & " y=" & Y

Select Case mSelMode
Case SelectionModeNone

Case SelectionModeCreate
    mTool.ResizeCreate X, Y
Case SelectionModeTool
    mTool.Move X, Y
    
    For Each selHandle In mSelectionHandles
        selHandle.X = selHandle.X + X - mPrevX
        selHandle.Y = selHandle.Y + Y - mPrevY
    Next
    
    mPrevX = X
    mPrevY = Y
Case Else
    Set selHandle = mSelectionHandles(mSelMode - SelectionModes.SelectionModeSelectionHandle1 + 1)
    mTool.MoveResize mSelMode - SelectionModes.SelectionModeSelectionHandle1 + 1, _
                    IIf(selHandle.AllowXMove, X, selHandle.X), _
                    IIf(selHandle.AllowYMove, Y, selHandle.Y)
    If selHandle.AllowXMove Then selHandle.X = X
    If selHandle.AllowYMove Then selHandle.Y = Y
End Select

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mRegionEV_MouseUp( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)

Const ProcName As String = "mRegionEV_MouseUp"
Dim failpoint As String
On Error GoTo Err

mTool.Click X, Y

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get IsSelected() As Boolean
Select Case mSelMode
Case SelectionModeNone

Case SelectionModeCreate

Case Else
    IsSelected = True
End Select
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddSelectableObject( _
                ByVal obj As IGraphicObject)
Const ProcName As String = "AddSelectableObject"
Dim failpoint As String
On Error GoTo Err

mSelectableObjects.Add obj

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub AddSelectionHandle( _
                ByVal X As Double, _
                ByVal Y As Double, _
                Optional ByVal AllowXMove As Boolean = True, _
                Optional ByVal AllowYMove As Boolean = True)
Dim dp As DataPoint
Dim selHandle As New SelectionHandle

Const ProcName As String = "AddSelectionHandle"
Dim failpoint As String
On Error GoTo Err

Set dp = mSelectionHandleSeries.Add(CDate(0))
dp.X = X
dp.DataValue = Y
selHandle.Initialise dp, AllowXMove, AllowYMove
mSelectionHandles.Add selHandle

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub Finish()

Const ProcName As String = "Finish"
Dim failpoint As String
On Error GoTo Err

mToolInControl = False

Select Case mSelMode
Case SelectionModeNone

Case SelectionModeCreate
    mSelMode = SelectionModeNone
    Set mRegionEV = Nothing
    Set mChartEV = Nothing
Case SelectionModeTool

Case Else
    mSelMode = SelectionModeTool
End Select

mChart.SetPointerModeSelection
mTool.EndResize

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Initialise( _
                ByVal pChart As ChartController, _
                ByVal tool As IChartTool, _
                ByVal pointerIcon As IPictureDisp)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mChart = pChart
Set mChartEV = pChart
Set mToolRef = CreateWeakReference(tool)
Set mPointerIcon = pointerIcon
Set mSelectableObjects = New Collection
mChart.SetPointerModeTool

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub MoveSelectionHandle( _
                ByVal index As Long, _
                ByVal X As Double, _
                ByVal Y As Double)
Dim selHandle As SelectionHandle
Const ProcName As String = "MoveSelectionHandle"
Dim failpoint As String
On Error GoTo Err

Set selHandle = mSelectionHandles.Item(index)
selHandle.X = Round(X)
selHandle.Y = Y

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function SetSelected() As Boolean
Const ProcName As String = "SetSelected"
Dim failpoint As String
On Error GoTo Err

If Not mSelMode = SelectionModeCreate Then
    mSelMode = SelectionModeTool
    Set mRegionEV = mRegion
    Set mChartEV = mChart
    SetSelected = True
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub Delete()
Dim selHandle As SelectionHandle
Const ProcName As String = "Delete"
Dim failpoint As String
On Error GoTo Err

For Each selHandle In mSelectionHandles
    mRegion.RemoveGraphicObject selHandle.DataPoint
Next
mRegion.RemoveGraphicObjectSeries mSelectionHandleSeries
Set mSelectableObjects = Nothing
mChart.SetPointerModeSelection
Set mTool = Nothing
Set mRegionEV = Nothing
Set mChartEV = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub Deselected()
Dim selHandle As SelectionHandle
Const ProcName As String = "Deselected"
Dim failpoint As String
On Error GoTo Err

mSelMode = SelectionModeNone
For Each selHandle In mSelectionHandles
    mRegion.RemoveGraphicObject selHandle.DataPoint
Next
Set mSelectionHandles = New Collection
Set mRegionEV = Nothing
Set mChartEV = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function SelectionHandleHit( _
                ByVal X As Double, _
                ByVal Y As Double) As Boolean
Dim i As Long
Dim selHandle As SelectionHandle
Dim graphObj As IGraphicObject

Const ProcName As String = "SelectionHandleHit"
Dim failpoint As String
On Error GoTo Err

For i = 1 To mSelectionHandles.count
    Set selHandle = mSelectionHandles(i)
    
    Set graphObj = selHandle.DataPoint
    If graphObj.HitTest(X, Y) Then
        mSelMode = SelectionModes.SelectionModeSelectionHandle1 + i - 1
        SelectionHandleHit = True
    End If
Next

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub StartMove()
Dim selHandle As SelectionHandle

Const ProcName As String = "StartMove"
Dim failpoint As String
On Error GoTo Err

For Each selHandle In mSelectionHandles
    selHandle.StartMove
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function Undo() As Boolean
Dim selHandle As SelectionHandle

Const ProcName As String = "Undo"
Dim failpoint As String
On Error GoTo Err

Select Case mSelMode
Case SelectionModeNone
    mChart.SetPointerModeDefault
    Undo = True
Case SelectionModeCreate
    mTool.Delete
    Delete
    Undo = True
Case SelectionModeTool
    If mToolInControl Then
        mToolInControl = False
        mTool.Undo
        For Each selHandle In mSelectionHandles
            selHandle.UndoMove
        Next
        Undo = True
    End If
Case Else
    If mToolInControl Then
        mToolInControl = False
        mTool.Undo
        For Each selHandle In mSelectionHandles
            selHandle.UndoMove
        Next
        Undo = True
    End If
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

