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

Private WithEvents mController              As ChartController
Attribute mController.VB_VarHelpID = -1

Private mRegion                             As ChartRegion
Private WithEvents mRegionEV                As ChartRegion
Attribute mRegionEV.VB_VarHelpID = -1

Private mSelectionHandleSeries              As DataPointSeries
Private mSelectionHandles                   As Collection

Private mSelectableObjects                 As Collection

Private mPrevPointerStyle                   As PointerStyles
Private mPrevAllowHorizScroll               As Boolean
Private mPrevAllowVertScroll                As Boolean
Private mPrevSnapCursor                     As Boolean

Private mSelMode                            As SelectionModes

Private mPrevX                              As Double
Private mPrevY                              As Double

Private mLayer                              As LayerNumbers

Private mToolRef                            As WeakReference
Private mTool                               As IChartTool

Private mMoving                             As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mController Event Handlers
'@================================================================================

Private Sub mController_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
Case vbKeyDelete
    If mSelMode = SelectionModeNone Or mSelMode = SelectionModeCreate Then Exit Sub
    mTool.Delete
    Delete
    KeyCode = 0
Case vbKeyEscape
    Select Case mSelMode
    Case SelectionModeNone
        restoreControllerSettings
    Case SelectionModeCreate
        mTool.Delete
        Delete
        KeyCode = 0
    Case Else
        If mMoving Then
            mTool.Undo
            Undo
'            Deselected
'            mRegion.ClearSelection
        End If
        mMoving = False
    End Select
End Select
End Sub

Private Sub mController_RegionSelected(ByVal Region As ChartSkil26.ChartRegion)
If mRegion Is Nothing Then
    
    Set mTool = mToolRef.Target ' convert to a strong ref for efficiency, now that we know
                                ' that this tool object is actually being used. This will also
                                ' keep both objects alive
    Set mRegion = Region
    Set mRegionEV = Region
    mTool.Region = Region

    mPrevSnapCursor = mRegion.SnapCursorToTickBoundaries
    mRegion.SnapCursorToTickBoundaries = False
        
    Set mSelectionHandleSeries = mRegion.AddGraphicObjectSeries(New DataPointSeries, LayerNumbers.LayerMax)
    mSelectionHandleSeries.includeInAutoscale = False
    mSelectionHandleSeries.displayMode = DataPointDisplayModePoint
    mSelectionHandleSeries.pointStyle = PointSquare
    mSelectionHandleSeries.lineThickness = 5
    mSelectionHandleSeries.Color = vbBlack
    
    Set mSelectionHandles = New Collection
    
End If
End Sub

'@================================================================================
' mRegionEV Event Handlers
'@================================================================================

Private Sub mRegionEV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim obj As IGraphicObject

Debug.Print "x=" & X & " y=" & Y

mPrevX = X
mPrevY = Y

If mSelectableObjects.Count = 0 Then
    ' the user must be creating
    mTool.Create X, Y
    mSelMode = SelectionModeCreate
Else
    For Each obj In mSelectableObjects
        If obj.Selected Then Exit For
    Next

    If obj Is Nothing Then
        ' the user has deselected
        Deselected
        mTool.Deselected
    Else
        Dim i As Long
        Dim selHandle As SelectionHandle
        Dim graphObj As IGraphicObject
        
        For i = 1 To mSelectionHandles.Count
            Set selHandle = mSelectionHandles(i)
            selHandle.StartMove
            
            Set graphObj = selHandle.DataPoint
            If graphObj.HitTest(X, Y) Then
                mSelMode = SelectionModes.SelectionModeSelectionHandle1 + i - 1
            End If
        Next
        
        mTool.StartResize selHandle, X, Y
        
        mMoving = True
        saveRegionSettings
    End If
End If

End Sub

Private Sub mRegionEV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim selHandle As SelectionHandle

If Not CBool(Button And MouseButtonConstants.vbLeftButton) Then Exit Sub

Debug.Print "x=" & X & " y=" & Y

Select Case mSelMode

Case SelectionModeNone
Case SelectionModeCreate
    mTool.ResizeCreate X, Y
Case SelectionModeTool
    If mMoving Then
        mTool.Move X, Y
        
        For Each selHandle In mSelectionHandles
            selHandle.X = selHandle.X + X - mPrevX
            selHandle.Y = selHandle.Y + Y - mPrevY
        Next
        
        mPrevX = X
        mPrevY = Y
    End If
Case Else
    If mMoving Then
        Set selHandle = mSelectionHandles(mSelMode - SelectionModes.SelectionModeSelectionHandle1 + 1)
        mTool.MoveResize mSelMode - SelectionModes.SelectionModeSelectionHandle1 + 1, X, Y
        selHandle.X = X
        selHandle.Y = Y
    End If
End Select
End Sub

Private Sub mRegionEV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case mSelMode
Case SelectionModeNone

Case SelectionModeCreate
    Set mRegionEV = Nothing
    restoreControllerSettings
Case SelectionModeTool
    restoreRegionSettings
Case Else
    mSelMode = SelectionModeTool
    restoreRegionSettings
End Select
mTool.EndResize
mMoving = False
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
mSelectableObjects.Add obj
End Sub

Public Sub AddSelectionHandle( _
                ByVal X As Double, _
                ByVal Y As Double)
Dim dp As DataPoint
Dim selHandle As New SelectionHandle

Set dp = mSelectionHandleSeries.Add(CDate(0))
dp.X = X
dp.dataValue = Y
selHandle.Initialise dp
mSelectionHandles.Add selHandle
End Sub

Friend Sub Initialise( _
                ByVal controller As ChartController, _
                ByVal tool As IChartTool)
Set mController = controller
Set mToolRef = CreateWeakReference(tool)
Set mSelectableObjects = New Collection
saveControllerSettings
End Sub

Public Sub SetSelected()
mSelMode = SelectionModeTool
Set mRegionEV = mRegion
End Sub

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
restoreRegionSettings
Set mTool = Nothing
Set mRegionEV = Nothing
End Sub

Private Sub Deselected()
Dim selHandle As SelectionHandle
mSelMode = SelectionModeNone
For Each selHandle In mSelectionHandles
    mRegion.RemoveGraphicObject selHandle.DataPoint
Next
Set mSelectionHandles = New Collection
Set mRegionEV = Nothing
End Sub

Private Sub restoreControllerSettings()
mController.PointerStyle = mPrevPointerStyle
mController.AllowHorizontalMouseScrolling = mPrevAllowHorizScroll
mController.AllowVerticalMouseScrolling = mPrevAllowVertScroll
mRegion.SnapCursorToTickBoundaries = mPrevSnapCursor
End Sub

Private Sub restoreRegionSettings()
mRegion.PointerStyle = mPrevPointerStyle
mController.AllowHorizontalMouseScrolling = mPrevAllowHorizScroll
mController.AllowVerticalMouseScrolling = mPrevAllowVertScroll
mRegion.SnapCursorToTickBoundaries = mPrevSnapCursor
End Sub

Private Sub saveControllerSettings()
mPrevPointerStyle = mController.PointerStyle
mController.PointerStyle = PointerStyles.PointerTool
mPrevAllowHorizScroll = mController.AllowHorizontalMouseScrolling
mPrevAllowVertScroll = mController.AllowVerticalMouseScrolling
mController.AllowHorizontalMouseScrolling = False
mController.AllowVerticalMouseScrolling = False
End Sub

Private Sub saveRegionSettings()
mPrevPointerStyle = mRegion.PointerStyle
mRegion.PointerStyle = PointerStyles.PointerTool
mPrevAllowHorizScroll = mController.AllowHorizontalMouseScrolling
mPrevAllowVertScroll = mController.AllowVerticalMouseScrolling
mPrevSnapCursor = mRegion.SnapCursorToTickBoundaries
mController.AllowHorizontalMouseScrolling = False
mController.AllowVerticalMouseScrolling = False
mRegion.SnapCursorToTickBoundaries = False
End Sub

Private Sub Undo()
Dim selHandle As SelectionHandle
For Each selHandle In mSelectionHandles
    selHandle.undoMove
Next
restoreRegionSettings
End Sub

