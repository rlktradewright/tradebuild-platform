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

Private Const ProjectName                   As String = "ChartTools26"
Private Const ModuleName                    As String = "AbstractChartTool"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mControllerEV            As chartController
Attribute mControllerEV.VB_VarHelpID = -1
Private mController                         As chartController

Private mRegion                             As ChartRegion
Private WithEvents mRegionEV                As ChartRegion
Attribute mRegionEV.VB_VarHelpID = -1

Private mSelectionHandleSeries              As DataPointSeries
Private mSelectionHandles                   As Collection

Private mSelectableObjects                  As Collection

Private mSelMode                            As SelectionModes

Private mPrevX                              As Double
Private mPrevY                              As Double

Private mLayer                              As LayerNumbers

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
' mController Event Handlers
'@================================================================================

Private Sub mControllerEV_KeyDown(KeyCode As Integer, Shift As Integer)

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
End Sub

Private Sub mControllerEV_RegionSelected(ByVal region As ChartSkil26.ChartRegion)
If mRegion Is Nothing Then
    
    Set mTool = mToolRef.Target ' convert to a strong ref for efficiency, now that we know
                                ' that this tool object is actually being used. This will also
                                ' keep both objects alive
    Set mRegion = region
    Set mRegionEV = region
    mTool.region = region

    Set mSelectionHandleSeries = mRegion.AddGraphicObjectSeries(New DataPointSeries, LayerNumbers.LayerMax)
    mSelectionHandleSeries.includeInAutoscale = False
    mSelectionHandleSeries.displayMode = DataPointDisplayModePoint
    mSelectionHandleSeries.pointStyle = PointSquare
    mSelectionHandleSeries.lineThickness = 5
    mSelectionHandleSeries.color = vbBlack
    
    Set mSelectionHandles = New Collection
    
End If
End Sub

'@================================================================================
' mRegionEV Event Handlers
'@================================================================================

Private Sub mRegionEV_DblCLick()
mTool.DblClick
End Sub

Private Sub mRegionEV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim obj As IGraphicObject

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
            mController.SetPointerModeTool
            
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

End Sub

Private Sub mRegionEV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim selHandle As SelectionHandle

'If Not CBool(Button And MouseButtonConstants.vbLeftButton) Then Exit Sub
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
End Sub

Private Sub mRegionEV_MouseUp( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)

mTool.Click X, Y
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
mSelectableObjects.add obj
End Sub

Public Sub AddSelectionHandle( _
                ByVal X As Double, _
                ByVal Y As Double, _
                Optional ByVal AllowXMove As Boolean = True, _
                Optional ByVal AllowYMove As Boolean = True)
Dim dp As DataPoint
Dim selHandle As New SelectionHandle

Set dp = mSelectionHandleSeries.add(CDate(0))
dp.X = X
dp.dataValue = Y
selHandle.initialise dp, AllowXMove, AllowYMove
mSelectionHandles.add selHandle
End Sub

Public Sub finish()

mToolInControl = False

Select Case mSelMode
Case SelectionModeNone

Case SelectionModeCreate
    mSelMode = SelectionModeNone
    Set mRegionEV = Nothing
    Set mControllerEV = Nothing
Case SelectionModeTool

Case Else
    mSelMode = SelectionModeTool
End Select
mController.SetPointerModeDefault
mTool.EndResize
End Sub

Friend Sub initialise( _
                ByVal controller As chartController, _
                ByVal tool As IChartTool, _
                ByVal pointerIcon As IPictureDisp)
Set mController = controller
Set mControllerEV = controller
Set mToolRef = CreateWeakReference(tool)
Set mPointerIcon = pointerIcon
Set mSelectableObjects = New Collection
mController.SetPointerModeTool
End Sub

Public Sub MoveSelectionHandle( _
                ByVal index As Long, _
                ByVal X As Double, _
                ByVal Y As Double)
Dim selHandle As SelectionHandle
Set selHandle = mSelectionHandles.item(index)
selHandle.X = Round(X)
selHandle.Y = Y
End Sub

Public Function SetSelected() As Boolean
If Not mSelMode = SelectionModeCreate Then
    mSelMode = SelectionModeTool
    Set mRegionEV = mRegion
    Set mControllerEV = mController
    SetSelected = True
End If
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub Delete()
Dim selHandle As SelectionHandle
For Each selHandle In mSelectionHandles
    mRegion.RemoveGraphicObject selHandle.DataPoint
Next
mRegion.RemoveGraphicObjectSeries mSelectionHandleSeries
Set mSelectableObjects = Nothing
mController.SetPointerModeDefault
Set mTool = Nothing
Set mRegionEV = Nothing
Set mControllerEV = Nothing
End Sub

Private Sub Deselected()
Dim selHandle As SelectionHandle
mSelMode = SelectionModeNone
For Each selHandle In mSelectionHandles
    mRegion.RemoveGraphicObject selHandle.DataPoint
Next
Set mSelectionHandles = New Collection
Set mRegionEV = Nothing
Set mControllerEV = Nothing
End Sub

Private Function SelectionHandleHit( _
                ByVal X As Double, _
                ByVal Y As Double) As Boolean
Dim i As Long
Dim selHandle As SelectionHandle
Dim graphObj As IGraphicObject

For i = 1 To mSelectionHandles.count
    Set selHandle = mSelectionHandles(i)
    
    Set graphObj = selHandle.DataPoint
    If graphObj.HitTest(X, Y) Then
        mSelMode = SelectionModes.SelectionModeSelectionHandle1 + i - 1
        SelectionHandleHit = True
    End If
Next
End Function

Private Sub StartMove()
Dim selHandle As SelectionHandle

For Each selHandle In mSelectionHandles
    selHandle.StartMove
Next
End Sub

Private Function Undo() As Boolean
Dim selHandle As SelectionHandle

Select Case mSelMode
Case SelectionModeNone
    mController.SetPointerModeDefault
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
            selHandle.undoMove
        Next
        Undo = True
    End If
Case Else
    If mToolInControl Then
        mToolInControl = False
        mTool.Undo
        For Each selHandle In mSelectionHandles
            selHandle.undoMove
        Next
        Undo = True
    End If
End Select

mController.SetPointerModeDefault
End Function

