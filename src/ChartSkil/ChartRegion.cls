VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements DeferredAction
Implements Task

'================================================================================
' Events
'================================================================================

Event AutoscaleChanged()

Event Click()
Attribute Click.VB_UserMemId = -600

Event DblCLick()
Attribute DblCLick.VB_UserMemId = -601

Event MouseDown(button As Integer, _
                Shift As Integer, _
                x As Single, _
                y As Single)
Attribute MouseDown.VB_UserMemId = -605
                
Event mouseMove(button As Integer, _
                Shift As Integer, _
                x As Single, _
                y As Single)
Attribute mouseMove.VB_UserMemId = -606
                
Event MouseUp(button As Integer, _
                Shift As Integer, _
                x As Single, _
                y As Single)
Attribute MouseUp.VB_UserMemId = -607

'================================================================================
' Enums
'================================================================================

Private Enum DeferredPaintActions
    DeferredPaintRepaintNone
    DeferredPaintRepaintChangedAreas
    DeferredPaintRepaintRegion
    DeferredPaintChangeGaugeandRepaint
    DeferredPaintScaleAndRepaint
    DeferredPaintCalculateAndRepaint
End Enum

'================================================================================
' Constants
'================================================================================

Private Const ProjectName                As String = "ChartSkil26"
Private Const ModuleName                As String = "ChartRegion"

'================================================================================
' Types
'================================================================================

Private Type PeriodTableEntry
    firstIndex                  As Long
    lastindex                   As Long
    BackColor                   As Long
End Type

Private Type GraphicObjectTableEntry
    graphObj            As IGraphicObject
    seriesId            As Long
    boundingRect        As TRectangle
    PeriodNumber        As Long
    ExtendedObject      As Boolean
    includeInAutoscale  As Boolean
    ScaleDependent      As Boolean
    GaugeDependent      As Boolean
    NoDraw              As Boolean
    Capabilities        As Long
    Layer               As Long
    prevPeriodIndex     As Long
    nextPeriodIndex     As Long
    prevLayerIndex      As Long
    nextLayerIndex      As Long
    prevExtendedIndex   As Long
    nextExtendedIndex   As Long
    inScopeCycle        As Long
    prevInScopeIndex    As Long
    nextInScopeIndex    As Long
    visibleCycle        As Long
    prevVisibleIndex    As Long
    nextVisibleIndex    As Long
End Type

Private Type RedrawObjectsTableEntry
    graphObj            As IGraphicObject
    areaToRedraw        As TRectangle
End Type

Private Type VisibleTableEntry
    firstIndex          As Long
    lastindex           As Long
End Type

'================================================================================
' Member variables
'================================================================================

Private mController As ChartController

Private mCleared As Boolean

Private mName As String
Private mRegionNumber As Long

Private mIsXAxisRegion As Boolean
Private mIsYAxisRegion As Boolean

Private mStyle As ChartRegionStyle

Private mTitle As text

Private mPeriodTable() As PeriodTableEntry

Private mGraphicObjects() As GraphicObjectTableEntry
Private mFirstGraphicObjectIndex As Long    ' the index of the first table entry
                                            ' that actually points to a graphic
                                            ' object - all prior entries are
                                            ' merely list start and end entries
Private mNextGraphicObjectIndex As Long

Private mMaxPeriodNumber As Long

Private mInScopeCycle As Long
Private mFirstInScopeIndex As Long
Private mLastInScopeIndex As Long

Private mVisibleCycle As Long
Private mVisibleTable(MinLayer To MaxLayer) As VisibleTableEntry

Private mFirstExtendedIndex As Long
Private mLastExtendedIndex As Long

Private mGraphObjSeriesTable() As IGraphicObjectSeries
Private mGraphObjSeriesIndex As Long

Private WithEvents mFirstBarSeries As BarSeries
Attribute mFirstBarSeries.VB_VarHelpID = -1

Private mDefaultBarSeries As BarSeries
Private mDefaultDataPointSeries As DataPointSeries
Private mDefaultTextSeries As TextSeries
Private mDefaultLineSeries As LineSeries

Private mCanvas As Canvas
Private mSuppressDrawing As Boolean

Private WithEvents mToolbar As Toolbar
Attribute mToolbar.VB_VarHelpID = -1

'Private mGaugeX As Double
'Private mGaugeY As Double

Private mMinInViewPeriod As Long
Private mMaxInViewPeriod As Long

Private mNextLayer As Long

Private mDrawn As Boolean

Private mGraphObjServiceProvider As GraphObjServiceProvider
Attribute mGraphObjServiceProvider.VB_VarHelpID = -1

Private mYAxisRegion As ChartRegion
Private mYPositionText As text      ' used to indicate the position of the pointer
                                    ' in the Y axis region
Private mYGridTexts As TextSeries
Private mYGridLines As LineSeries

Private mXGridTextsCoarse As TextSeries
Private mXGridTexts As TextSeries
Private mXGridLines As LineSeries

' indicates whether the grid is currently
' hidden
Private mHideGrid As Boolean

Private mPointerMode As PointerModes
Private mPointerStyle As PointerStyles
Private mPointerIcon As IPictureDisp
Private mPointerToolStyle As PointerStyles
Private mPointerToolIcon As IPictureDisp

Private mSessionOffsetMinutes As Date
Private mVerticalGridTimePeriod As TimePeriod

Private mPercentheight As Double
Private mMinimumPercentHeight As Double
Private mYAxisPosition As Long

Private mScaleGridSpacingY As Double
Private mYScaleFormatStr As String
Private mYScaleSubFormatStr As String
Private mGridTextHeight As Double

Private mShowPerformanceText As Boolean

Private mPointerCrosshairsColor As Long
Private mPointerDiscColor As Long

Private mPrevCursorX As Single
Private mPrevCursorY As Single

Private mPerfText As text

'Private mCurrentTool As ToolTypes

' defaults for bar series
Private mDefaultBarStyle As BarStyle

' defaults for datapoint series
Private mDefaultDataPointStyle As DataPointStyle

' defaults for text and text series
Private mDefaultTextStyle As TextStyle

' defaults for lines and line series
Private mDefaultLineStyle As linestyle

Private mElapsedTimer As ElapsedTimer

Private mTaskContext As TaskContext

Private mCurrX As Single
Private mCurrY As Single

Private mUseDeferredPainting As Boolean
Private mDeferredPaintPending As Boolean

Private mDeferredPaintAction As DeferredPaintActions

Private mUndrawnAreas() As TRectangle
Private mUndrawnAreasIndex As Long

Private mRedrawObjects() As RedrawObjectsTableEntry
Private mRedrawObjectsIndex As Long

Private mSelectedObject As IGraphicObject

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Dim i As Long

Set mStyle = New ChartRegionStyle

mUseDeferredPainting = True
deferredPaint DeferredPaintRepaintNone

mNextLayer = LayerNumbers.LayerLowestUser

ReDim mPeriodTable(255) As PeriodTableEntry
ReDim mGraphicObjects(1023) As GraphicObjectTableEntry

mFirstInScopeIndex = allocateGraphicObjectTableEntry
mLastInScopeIndex = allocateGraphicObjectTableEntry
initialiseInScopeObjectList

mFirstExtendedIndex = allocateGraphicObjectTableEntry
mLastExtendedIndex = allocateGraphicObjectTableEntry
mGraphicObjects(mFirstExtendedIndex).nextExtendedIndex = mLastExtendedIndex
mGraphicObjects(mLastExtendedIndex).prevExtendedIndex = mFirstExtendedIndex

For i = MinLayer To MaxLayer
    mVisibleTable(i).firstIndex = allocateGraphicObjectTableEntry
    mVisibleTable(i).lastindex = allocateGraphicObjectTableEntry
Next
initialiseVisibleObjectList

Set mGraphObjServiceProvider = New GraphObjServiceProvider
mGraphObjServiceProvider.Region = Me

ReDim mGraphObjSeriesTable(3) As IGraphicObjectSeries

End Sub

Private Sub Class_Terminate()
Debug.Print "ChartRegion terminated"
End Sub

'================================================================================
' DeferredAction Interface Members
'================================================================================

Private Sub DeferredAction_run(ByVal data As Variant)
Dim et As ElapsedTimer
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

If mCleared Then Exit Sub

If gIsInDev Or gLogger.isLoggable(LogLevelHighDetail) Then
    If et Is Nothing Then Set et = New ElapsedTimer
    et.StartTiming
End If

Select Case mDeferredPaintAction
Case DeferredPaintRepaintNone

Case DeferredPaintRepaintChangedAreas
    
    For i = 0 To mUndrawnAreasIndex
        undrawAreaEx mUndrawnAreas(i)
    Next
    
    For i = 0 To mRedrawObjectsIndex
        redrawObject mRedrawObjects(i).graphObj, mRedrawObjects(i).areaToRedraw
    Next
Case DeferredPaintRepaintRegion
    paint
Case DeferredPaintChangeGaugeandRepaint
    calculateInscopeGaugeDependentObjects
    paint
    If Not mYAxisRegion Is Nothing Then mYAxisRegion.resizedY
Case DeferredPaintScaleAndRepaint
    calculateVerticalBounds
    determineVisibleObjects
    paint
    If Not mYAxisRegion Is Nothing Then
        mYAxisRegion.SetVerticalScaleInternal mCanvas.Bottom, mCanvas.Top
    End If
Case DeferredPaintCalculateAndRepaint
    determineInScopeObjects
    If Autoscale Then
        calculateVerticalBounds
    End If
    determineVisibleObjects
    paint
    If Not mYAxisRegion Is Nothing Then
        mYAxisRegion.SetVerticalScaleInternal mCanvas.Bottom, mCanvas.Top
    End If
End Select


If gIsInDev Or gLogger.isLoggable(LogLevelHighDetail) Then
    Dim s As String
    s = "Deferred paint: " & Format(et.ElapsedTimeMicroseconds, "0")
    Debug.Print s
    gLogger.Log LogLevelHighDetail, s
End If

deferredPaint DeferredPaintRepaintNone

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "DeferredAction_run" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

End Sub

'================================================================================
' Task Interface Members
'================================================================================

Private Sub Task_cancel()

End Sub

Private Sub Task_run()
Static i As Long
Dim j As Long

Do While i < mNextGraphicObjectIndex
    If Not mGraphicObjects(i).graphObj Is Nothing Then Set mGraphicObjects(i).graphObj = Nothing
    i = i + 1
    j = j + 1
    If j = 100 Then Exit Sub
Loop

Erase mGraphicObjects

mTaskContext.Finish 0, False

End Sub

Private Property Let Task_TaskContext(ByVal RHS As TWUtilities30.TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.name
End Property

'================================================================================
' GraphObjServiceProvider Methods
'================================================================================

Friend Sub GOSP_AddGraphicObject( _
                ByVal graphObj As IGraphicObject, _
                ByVal seriesId As Long)
Dim failpoint As Long
On Error GoTo Err

AddGraphicObject graphObj, seriesId

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GOSP_addGraphicObject" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub GOSP_ObjectChanged( _
                ByVal Handle As Long, _
                ByRef areaToRedraw As TRectangle)
Dim redrawArea As TRectangle

Dim failpoint As Long
On Error GoTo Err

' the following is to prevent lines from going past the rightmost period of the chart
redrawArea = rectIntersection(areaToRedraw, mCanvas.boundary)

With mGraphicObjects(Handle)
    
    ' get the new properties for the changed object
    
    .boundingRect = .graphObj.BoundingRectangle
    
    If Not .boundingRect.isValid Then
        removeEntryFromInScopeList Handle
        removeEntryFromVisibleList Handle
    End If
    
    .ExtendedObject = .graphObj.ExtendedObject
    .NoDraw = .graphObj.NoDraw
    If .Layer <> .graphObj.Layer Then removeEntryFromVisibleList Handle
    .Layer = .graphObj.Layer
    .includeInAutoscale = .graphObj.includeInAutoscale
    .ScaleDependent = .graphObj.ScaleDependent
    .Capabilities = .graphObj.Capabilities
End With

If mGraphicObjects(Handle).ExtendedObject Then
    removeEntryFromPeriodList Handle
    addEntryToExtendedList Handle
Else
    removeEntryFromExtendedList Handle
    
    ' The following call to addEntryToPeriodList must not be in a With block
    ' involving mGraphicObjects because in rare circumstances it can result
    ' in mGraphicObjects being ReDim'd which leads to an error 10
    addEntryToPeriodList Handle
End If

If graphicObjectIsInScope(Handle) Then
    addEntryToInScopeList Handle
    If graphicObjectIsVisible(Handle) Then
        addEntryToVisibleList Handle
    Else
        removeEntryFromVisibleList Handle
    End If
Else
    removeEntryFromInScopeList Handle
    Exit Sub
End If
    
If mSuppressDrawing Then Exit Sub

If Not mGraphicObjects(Handle).boundingRect.isValid Then Exit Sub
If mGraphicObjects(Handle).NoDraw Then Exit Sub

With mGraphicObjects(Handle)
    ' check if the vertical bounds have changed, and redraw if so
    If Autoscale And .includeInAutoscale Then
        If .boundingRect.Top >= mCanvas.boundary.Top Or .boundingRect.Bottom <= (mCanvas.boundary.Bottom) Then
            If mUseDeferredPainting Then
                deferredPaint DeferredPaintScaleAndRepaint
            Else
                calculateVerticalBounds
                determineVisibleObjects
                paint
            End If
            Exit Sub
        End If
    End If
    
    ' now redraw the changed object

    If redrawArea.isValid Then
        
        If mUseDeferredPainting Then
            mRedrawObjectsIndex = mRedrawObjectsIndex + 1
            If mRedrawObjectsIndex > UBound(mRedrawObjects) Then
                ReDim Preserve mRedrawObjects(2 * (UBound(mRedrawObjects) + 1) - 1) As RedrawObjectsTableEntry
            End If
            Set mRedrawObjects(mRedrawObjectsIndex).graphObj = .graphObj
            mRedrawObjects(mRedrawObjectsIndex).areaToRedraw = areaToRedraw
            deferredPaint DeferredPaintRepaintChangedAreas
        Else
            redrawObject .graphObj, areaToRedraw
        End If
    End If
        
End With

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GOSP_objectChanged" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription



End Sub

Friend Sub GOSP_ObjectUndrawn(ByVal Handle As Long, undrawnArea As TRectangle)

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

undrawArea undrawnArea

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GOSP_objectUndrawn" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub GOSP_PaintRegion( _
                ByVal recalculate As Boolean)

Dim failpoint As Long
On Error GoTo Err

If mSuppressDrawing Then Exit Sub

If mUseDeferredPainting Then
    If recalculate Then
        deferredPaint DeferredPaintCalculateAndRepaint
    Else
        deferredPaint DeferredPaintRepaintRegion
    End If
Else
    If recalculate Then
        determineInScopeObjects
        If Autoscale Then
            calculateVerticalBounds
        End If
        determineVisibleObjects
    End If
    paint
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GOSP_paintRegion" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub GOSP_RecalcBoundary( _
                ByVal seriesId As Long)
Dim i As Long
Dim index As Long

Dim failpoint As Long
On Error GoTo Err

For i = MinLayer To MaxLayer
    index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(i).lastindex
        With mGraphicObjects(index)
            If .seriesId = seriesId Then
                .boundingRect = .graphObj.BoundingRectangle
            End If
            index = .nextVisibleIndex
        End With
    Loop
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GOSP_recalcBoundary" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub GOSP_RemoveObject( _
                ByVal Handle As Long)
RemoveGraphicObjectEx mGraphicObjects(Handle).graphObj, True, False
End Sub

'================================================================================
' mFirstBarSeries Event Handlers
'================================================================================

Private Sub mFirstBarSeries_StyleChanged()
setBarModeButtons
End Sub

'================================================================================
' mToolbar Event Handlers
'================================================================================

Private Sub mToolbar_ButtonClick(ByVal button As MSComctlLib.button)

Dim failpoint As Long
On Error GoTo Err

Select Case button.Key
Case ToolbarCommandShowBars
    If mFirstBarSeries.barThickness > 1 Then
        mToolbar.Buttons(ToolbarCommandThinnerBars).Enabled = True
    Else
        mToolbar.Buttons(ToolbarCommandThinnerBars).Enabled = False
    End If
    mFirstBarSeries.displayMode = BarDisplayModeBar
Case ToolbarCommandShowCandlesticks
    mFirstBarSeries.displayMode = BarDisplayModeCandlestick
    If mFirstBarSeries.barWidth <= 0.1 Then
        mToolbar.Buttons(ToolbarCommandThinnerBars).Enabled = False
    Else
        mToolbar.Buttons(ToolbarCommandThinnerBars).Enabled = True
    End If
Case ToolbarCommandShowLine
    ' not yet implemented in ChartSkil
Case ToolbarCommandThinnerBars
    If mFirstBarSeries.displayMode = BarDisplayModeCandlestick Then
        If mFirstBarSeries.barWidth > 0.1 Then
            mFirstBarSeries.barWidth = mFirstBarSeries.barWidth - 0.1
        End If
        If mFirstBarSeries.barWidth <= 0.1 Then
            button.Enabled = False
        End If
    ElseIf mFirstBarSeries.displayMode = BarDisplayModeBar Then
        If mFirstBarSeries.barThickness > 1 Then
            mFirstBarSeries.barThickness = mFirstBarSeries.barThickness - 1
        End If
        If mFirstBarSeries.barThickness = 1 Then
            button.Enabled = False
        End If
    End If
Case ToolbarCommandThickerBars
    If mFirstBarSeries.displayMode = BarDisplayModeCandlestick Then
        mFirstBarSeries.barWidth = mFirstBarSeries.barWidth + 0.1
    ElseIf mFirstBarSeries.displayMode = BarDisplayModeBar Then
        mFirstBarSeries.barThickness = mFirstBarSeries.barThickness + 1
    End If
    mToolbar.Buttons(ToolbarCommandThinnerBars).Enabled = True
Case ToolbarCommandScaleDown
    ScaleUp -0.09091
Case ToolbarCommandScaleUp
    ScaleUp 0.1
Case ToolbarCommandScrollDown
    ScrollVerticalProportion -0.2
Case ToolbarCommandScrollUp
    ScrollVerticalProportion 0.2
Case ToolbarCommandAutoScale
    Autoscale = True
    If Not mSuppressDrawing Then
        If mUseDeferredPainting Then
            deferredPaint DeferredPaintScaleAndRepaint
        Else
            calculateVerticalBounds
            determineVisibleObjects
            paint
        End If
    End If
End Select

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "mToolbar_ButtonClick" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription



End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get Autoscale() As Boolean
Autoscale = mStyle.Autoscale
End Property

Public Property Let Autoscale(ByVal value As Boolean)
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "autoScale", _
            "Cannot set autoscale for an axis region"
End If

mStyle.Autoscale = value
If Not mToolbar Is Nothing Then
    mToolbar.Buttons("autoscale").Enabled = IIf(Autoscale, False, True)
End If
RaiseEvent AutoscaleChanged
End Property

Public Property Get BackColor() As Long
BackColor = mStyle.BackColor
End Property

Public Property Let BackColor(ByVal value As Long)
mStyle.BackColor = value
mCanvas.BackColor = value
End Property

Public Property Get BackGradientFillColors() As Long()
BackGradientFillColors = mStyle.BackGradientFillColors
End Property

Public Property Let BackGradientFillColors(ByRef value() As Long)
Dim ar() As Long
ar = value
mStyle.BackGradientFillColors = ar
mCanvas.gradientFillColors = ar
End Property

Public Property Get Bottom() As Double
Bottom = mCanvas.boundary.Bottom
End Property

Friend Property Let Bottom(ByVal value As Double)
mCanvas.Bottom = value
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.Bottom = value
End If
End Property

Friend Property Let Canvas( _
                ByVal value As Canvas)
Set mCanvas = value
End Property

Friend Property Let controller( _
                ByVal value As ChartController)
Set mController = value
End Property

Public Property Get DefaultBarStyle() As BarStyle
Set DefaultBarStyle = mDefaultBarStyle.clone
End Property

Public Property Let DefaultBarStyle(ByVal value As BarStyle)
Set mDefaultBarStyle = value.clone
End Property

Public Property Get DefaultDataPointStyle() As DataPointStyle
Set DefaultDataPointStyle = mDefaultDataPointStyle.clone
End Property

Public Property Let DefaultDataPointStyle(ByVal value As DataPointStyle)
Set mDefaultDataPointStyle = value.clone
End Property

Public Property Get DefaultLineStyle() As linestyle
Set DefaultLineStyle = mDefaultLineStyle.clone
End Property

Public Property Let DefaultLineStyle(ByVal value As linestyle)
Set mDefaultLineStyle = value.clone
End Property

Public Property Get DefaultTextStyle() As TextStyle
Set DefaultTextStyle = mDefaultTextStyle.clone
End Property

Public Property Let DefaultTextStyle(ByVal value As TextStyle)
Set mDefaultTextStyle = value.clone
End Property

Public Property Get FirstBarSeries() As BarSeries
Set FirstBarSeries = mFirstBarSeries
End Property

Public Property Get GridColor() As Long
GridColor = mStyle.GridColor
End Property

Public Property Let GridColor(ByVal val As Long)
mStyle.GridColor = val
End Property

Public Property Get GridlineSpacingY() As Double
GridlineSpacingY = mStyle.GridlineSpacingY
End Property

Public Property Let GridlineSpacingY(ByVal value As Double)
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "gridlineSpacingY", _
            "Cannot set gridline spacing for an axis region"
End If

mStyle.GridlineSpacingY = value
Debug.Print "Region " & mName & ": GridlineSpacingY=" & mStyle.GridlineSpacingY
End Property

Public Property Get GridTextColor() As Long
GridTextColor = mStyle.GridTextColor
End Property

Public Property Let GridTextColor(ByVal val As Long)
mStyle.GridTextColor = val
End Property

Public Property Get HasGrid() As Boolean
HasGrid = mStyle.HasGrid
End Property

Public Property Let HasGrid(ByVal val As Boolean)
mStyle.HasGrid = val
If HasGrid Then
    Set mXGridLines = Me.AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerGrid)
    mXGridLines.Color = GridColor
    mXGridLines.extendAfter = True
    mXGridLines.extendBefore = True
    mXGridLines.includeInAutoscale = False
    mXGridLines.linestyle = LineSolid
    mXGridLines.thickness = 1
    
    Set mYGridLines = Me.AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerGrid)
    mYGridLines.Color = GridColor
    mYGridLines.extendAfter = True
    mYGridLines.extendBefore = True
    mYGridLines.includeInAutoscale = False
    mYGridLines.linestyle = LineSolid
    mYGridLines.thickness = 1
Else
    Set mXGridLines = Nothing
    Set mYGridLines = Nothing
End If
End Property

Public Property Get HasGridText() As Boolean
HasGridText = mStyle.HasGridText
End Property

Public Property Let HasGridText(ByVal val As Boolean)
Dim aFont As StdFont

mStyle.HasGridText = val
If HasGridText Then
    Set mXGridTexts = Me.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText)
    mXGridTexts.align = AlignBottomCentre
    If Not mIsXAxisRegion Then
        mXGridTexts.box = True
        mXGridTexts.boxFillColor = BackColor
        mXGridTexts.boxStyle = LineInvisible
    End If
    mXGridTexts.extended = True
    mXGridTexts.Color = GridTextColor
    Set aFont = New StdFont
    aFont.name = "Arial"
    aFont.size = 8
    aFont.Underline = False
    aFont.Bold = False
    mXGridTexts.font = aFont

    Set mXGridTextsCoarse = Me.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText)
    mXGridTextsCoarse.align = AlignTopCentre
    If Not mIsXAxisRegion Then
        mXGridTextsCoarse.box = True
        mXGridTextsCoarse.boxFillColor = BackColor
    End If
    mXGridTextsCoarse.boxStyle = LineInvisible
    mXGridTextsCoarse.extended = True
    mXGridTextsCoarse.Color = GridTextColor
    Set aFont = New StdFont
    aFont.name = "Arial"
    aFont.size = 8
    aFont.Underline = False
    aFont.Bold = False
    mXGridTextsCoarse.font = aFont
Else
    Set mXGridTexts = Nothing
    Set mXGridTextsCoarse = Nothing
End If
End Property

Public Property Get height() As Double
height = mCanvas.height
End Property

Public Property Let IntegerYScale(ByVal value As Boolean)
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "integerYScale", _
            "Cannot set integerYScale for an axis region"
End If

mStyle.IntegerYScale = value
End Property

Public Property Get IntegerYScale() As Boolean
IntegerYScale = mStyle.IntegerYScale
End Property

Friend Property Let IsXAxisRegion(ByVal value As Boolean)
mIsXAxisRegion = value
End Property

Public Property Get IsXAxisRegion() As Boolean
IsXAxisRegion = mIsXAxisRegion
End Property

Friend Property Let IsYAxisRegion(ByVal value As Boolean)
mIsYAxisRegion = value
End Property

Public Property Get IsYAxisRegion() As Boolean
IsYAxisRegion = mIsYAxisRegion
End Property

Public Property Get Left() As Double
Left = mCanvas.Left
End Property

Friend Property Let Left(ByVal value As Double)
mCanvas.Left = value
End Property

Public Property Get MinimumHeight() As Double
MinimumHeight = mStyle.MinimumHeight
End Property

Public Property Let MinimumHeight(ByVal value As Double)
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "minimumHeight", _
            "Cannot set minimumHeight for an axis region"
End If

mStyle.MinimumHeight = value
End Property

Public Property Get minimumPercentHeight() As Double
minimumPercentHeight = mMinimumPercentHeight
End Property

Friend Property Let minimumPercentHeight(ByVal value As Double)
mMinimumPercentHeight = value
End Property

Public Property Get MousePosition() As Point
Set MousePosition = newPoint(mCurrX, mCurrY)
End Property

Public Property Get name() As String
name = mName
End Property

Friend Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get percentheight() As Double
percentheight = mPercentheight
End Property

Friend Property Let percentheight(ByVal value As Double)
mPercentheight = value
End Property

Public Property Get PointerCrosshairsColor() As Long
PointerCrosshairsColor = mPointerCrosshairsColor
End Property

Friend Property Let PointerCrosshairsColor(ByVal value As Long)
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "PointerCrosshairsColor", _
            "Cannot set PointerCrossHairsColor for an axis region"
End If

mPointerCrosshairsColor = value
End Property

Public Property Get PointerDiscColor() As Long
PointerDiscColor = mPointerDiscColor
End Property

Friend Property Let PointerDiscColor(ByVal value As Long)
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "PointerDiscColor", _
            "Cannot set PointerDiscColor for an axis region"
End If

mPointerDiscColor = value
End Property

Public Property Get PointerIcon() As IPictureDisp
Set PointerIcon = mPointerIcon
End Property

Public Property Let PointerIcon(ByVal value As IPictureDisp)

If value Is Nothing Then Exit Property
If value Is mPointerIcon Then Exit Property

Set mPointerIcon = value
End Property

Public Property Get PointerStyle() As PointerStyles
PointerStyle = mPointerStyle
End Property

Public Property Let PointerStyle(ByVal value As PointerStyles)
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "pointerStyle", _
            "Cannot set PointerStyle for an axis region"
End If


If mPointerStyle = value Then Exit Property

If mPointerMode = PointerModeTool Then
    mPointerStyle = value
    Exit Property
End If

If Not mDrawn Then
    mPointerStyle = value
    Exit Property
End If

Select Case mPointerStyle
Case PointerNone

Case PointerCrosshairs
    undrawCrosshairsPointer
Case PointerDisc
    undrawDiscPointer
Case PointerTool
    undrawToolPointer
Case PointerCustom
    undrawCustomPointer
End Select

mPointerStyle = value

Select Case mPointerStyle
Case PointerNone
    mCanvas.MouseIcon = gBlankMouseIcon
Case PointerCrosshairs
    mCanvas.MouseIcon = gBlankMouseIcon
    drawCrosshairsPointer mPrevCursorX, mPrevCursorY
Case PointerDisc
    mCanvas.MouseIcon = gBlankMouseIcon
    drawDiscPointer mPrevCursorX, mPrevCursorY
Case PointerTool
    mCanvas.MouseIcon = gBlankMouseIcon
    drawToolPointer mPrevCursorX, mPrevCursorY
Case PointerCustom
    mCanvas.MouseIcon = mPointerIcon
    drawCustomPointer mPrevCursorX, mPrevCursorY
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26.ChartRegion::(let)pointerStyle", _
            "Value must be a member of the PointerStyles enum"
End Select

End Property

Public Property Get RegionNumber() As Long
RegionNumber = mRegionNumber
End Property

Friend Property Let RegionNumber(ByVal val As Long)
mRegionNumber = val
End Property

Public Property Get ScaleGridSpacingY() As Single
ScaleGridSpacingY = mScaleGridSpacingY
End Property

Friend Property Let SessionStartTime(ByVal val As Date)
mSessionOffsetMinutes = calculateMinutesFromMidnight(val) Mod 60
End Property

Public Property Get ShowPerformanceText() As Boolean
ShowPerformanceText = mShowPerformanceText
End Property

Public Property Let ShowPerformanceText(ByVal val As Boolean)
mShowPerformanceText = val
If mShowPerformanceText Then
    If mElapsedTimer Is Nothing Then Set mElapsedTimer = New ElapsedTimer
Else
    Set mElapsedTimer = Nothing
End If
End Property

Public Property Let SnapCursorToTickBoundaries(ByVal value As Boolean)
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "snapCursorToTickBoundaries", _
            "Cannot set snapCursorToTickBoundaries for an axis region"
End If

mStyle.SnapCursorToTickBoundaries = value
End Property

Public Property Get SnapCursorToTickBoundaries() As Boolean
SnapCursorToTickBoundaries = mStyle.SnapCursorToTickBoundaries
End Property

Public Property Get style() As ChartRegionStyle
Set style = mStyle.clone
End Property

Public Property Let style( _
                ByVal value As ChartRegionStyle)
Dim lSuppress As Boolean

Dim failpoint As Long
On Error GoTo Err

If value Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "style", _
            "Style is null reference"
End If

lSuppress = mSuppressDrawing
mSuppressDrawing = True

BackColor = value.BackColor
BackGradientFillColors = value.BackGradientFillColors
If value.GridColor <> GridColor Then GridColor = value.GridColor
If value.GridTextColor <> GridTextColor Then GridTextColor = value.GridTextColor
If Not mIsXAxisRegion And Not mIsYAxisRegion Then
    If value.Autoscale <> Autoscale Then Autoscale = value.Autoscale
    If value.GridlineSpacingY <> GridlineSpacingY Then GridlineSpacingY = value.GridlineSpacingY
    If value.IntegerYScale <> IntegerYScale Then IntegerYScale = value.IntegerYScale
    If value.MinimumHeight <> MinimumHeight Then MinimumHeight = value.MinimumHeight
    If value.SnapCursorToTickBoundaries <> SnapCursorToTickBoundaries Then SnapCursorToTickBoundaries = value.SnapCursorToTickBoundaries
    If value.YScaleQuantum <> YScaleQuantum Then YScaleQuantum = value.YScaleQuantum
End If
If value.HasGrid <> HasGrid Then HasGrid = value.HasGrid
If value.HasGridText <> HasGridText Then HasGridText = value.HasGridText

mSuppressDrawing = lSuppress

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Style" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Friend Property Let SuppressDrawing(ByVal val As Boolean)
If mSuppressDrawing = val Then Exit Property
mSuppressDrawing = val
If Not mSuppressDrawing Then
    If mUseDeferredPainting Then
        deferredPaint DeferredPaintCalculateAndRepaint
    Else
        determineInScopeObjects
        If Autoscale Then
            calculateVerticalBounds
        End If
        determineVisibleObjects
        paint
    End If
End If
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.SuppressDrawing = val
End If
End Property

Public Property Get Title() As text
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "Title", _
            "An axis region does not have a title"
End If

If mTitle Is Nothing Then setupTitle
Set Title = mTitle
End Property

Friend Property Let Toolbar(ByVal value As Toolbar)
Dim btn As button

Set mToolbar = value

mToolbar.Buttons("autoscale").Enabled = Not Autoscale
mToolbar.Buttons("showline").Enabled = False
                                        ' because line charts are not implemented yet

End Property

Public Property Get Top() As Double
Top = mCanvas.Top
End Property

Friend Property Let Top(ByVal value As Double)
mCanvas.Top = value
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.Top = value
End If
End Property

Friend Property Let VerticalGridTimePeriod(ByVal value As TimePeriod)
Set mVerticalGridTimePeriod = value
End Property

Public Property Get width() As Single
width = mCanvas.width
End Property

Public Property Get YAxisPosition() As Long
YAxisPosition = mYAxisPosition
End Property

Friend Property Let YAxisRegion(ByVal value As ChartRegion)
Dim aFont As StdFont

Set mYAxisRegion = value

mYAxisRegion.SuppressDrawing = mSuppressDrawing
mYAxisRegion.PeriodsInView 0, 99

mYAxisRegion.DefaultTextStyle = mDefaultTextStyle

Set mYPositionText = mYAxisRegion.AddText(LayerNumbers.LayerPointer)
mYPositionText.align = AlignCentreLeft
mYPositionText.Color = vbBlack  ' vbWhite Xor mBackColor
mYPositionText.box = True
mYPositionText.boxFillColor = vbWhite   'mBackColor
mYPositionText.boxStyle = LineSolid
mYPositionText.boxColor = vbBlack   ' vbWhite Xor mBackColor
Set aFont = New StdFont
aFont.name = "Arial"
aFont.size = 8
aFont.Underline = False
aFont.Bold = False
mYPositionText.font = aFont

Set mYGridTexts = mYAxisRegion.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText)
mYGridTexts.includeInAutoscale = False
mYGridTexts.align = AlignCentreLeft
mYGridTexts.Color = GridTextColor
mYGridTexts.box = True
mYGridTexts.boxFillColor = mYAxisRegion.BackColor
mYGridTexts.boxStyle = LineInvisible
mYGridTexts.boxColor = mYAxisRegion.BackColor
mYGridTexts.paddingX = 0
mYGridTexts.paddingY = 0
Set aFont = New StdFont
aFont.name = "Arial"
aFont.size = 8
aFont.Underline = False
aFont.Bold = False
mYGridTexts.font = aFont

End Property

Public Property Get YAxisRegion() As ChartRegion
Set YAxisRegion = mYAxisRegion
End Property

Public Property Let YScaleQuantum(ByVal value As Double)
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "YScaleQuantum", _
            "Cannot set YScaleQuantum for an axis region"
End If

mStyle.YScaleQuantum = value
End Property

Public Property Get YScaleQuantum() As Double
YScaleQuantum = mStyle.YScaleQuantum
End Property

'================================================================================
' Methods
'================================================================================

Public Function AddBar( _
                ByVal timestamp As Date, _
                Optional ByVal Layer As Long = -1, _
                Optional ByVal IsSelectable As Boolean) As ChartSkil26.Line
Dim failpoint As Long
On Error GoTo Err

If mDefaultBarSeries Is Nothing Then
    Set mDefaultBarSeries = AddGraphicObjectSeries(New DataPointSeries, LayerDefaultBars)
    mDefaultBarSeries.style = mDefaultBarStyle
End If
Set AddBar = mDefaultBarSeries.Add(timestamp, , , IsSelectable)
If Layer <> -1 Then AddBar.Layer = Layer

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "AddBar" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Public Function AddDataPoint( _
                ByVal timestamp As Date, _
                Optional ByVal Layer As Long = -1, _
                Optional ByVal IsSelectable As Boolean) As ChartSkil26.Line
Dim failpoint As Long
On Error GoTo Err

If mDefaultDataPointSeries Is Nothing Then
    Set mDefaultDataPointSeries = AddGraphicObjectSeries(New DataPointSeries, LayerDefaultDataPoints)
    mDefaultDataPointSeries.style = mDefaultDataPointStyle
End If
Set AddDataPoint = mDefaultDataPointSeries.Add(timestamp, , , IsSelectable)
If Layer <> -1 Then AddDataPoint.Layer = Layer

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "AddDataPoint" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Public Function AddGraphicObjectSeries( _
                ByVal graphObjSeries As IGraphicObjectSeries, _
                Optional ByVal Layer As Long = -1, _
                Optional ByVal name As String) As IGraphicObjectSeries
Dim failpoint As Long
On Error GoTo Err

Set AddGraphicObjectSeries = graphObjSeries

If mGraphObjSeriesIndex > UBound(mGraphObjSeriesTable) Then
    ReDim Preserve mGraphObjSeriesTable(2 * (UBound(mGraphObjSeriesTable) + 1) - 1) As IGraphicObjectSeries
End If

Set mGraphObjSeriesTable(mGraphObjSeriesIndex) = graphObjSeries

If name = "" Then name = "Series" & mGraphObjSeriesIndex

graphObjSeries.initialise name, _
                        mCanvas, _
                        getLayer(Layer), _
                        mController, _
                        mGraphObjSeriesIndex, _
                        mGraphObjServiceProvider

mGraphObjSeriesIndex = mGraphObjSeriesIndex + 1

If TypeOf graphObjSeries Is BarSeries And mFirstBarSeries Is Nothing Then
    Set mFirstBarSeries = graphObjSeries
    setBarModeButtons
End If

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "addGraphicObjectSeries" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Public Function AddLine( _
                Optional ByVal Layer As Long = -1, _
                Optional ByVal IsSelectable As Boolean) As ChartSkil26.Line
Dim failpoint As Long
On Error GoTo Err

If mDefaultLineSeries Is Nothing Then
    Set mDefaultLineSeries = AddGraphicObjectSeries(New LineSeries, LayerDefaultLines)
    mDefaultLineSeries.style = mDefaultLineStyle
End If
Set AddLine = mDefaultLineSeries.Add(, , IsSelectable)
If Layer <> -1 Then AddLine.Layer = Layer

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "addLine" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Friend Sub addPeriod(ByVal PeriodNumber As Long, _
                    ByVal timestamp As Date)
Dim failpoint As Long
On Error GoTo Err

addNewPeriod PeriodNumber
If CDbl(timestamp) <> 0 Then setVerticalGridLine timestamp, PeriodNumber

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "addPeriod" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Function AddText( _
                Optional ByVal Layer As Long = -1, _
                Optional ByVal IsSelectable As Boolean) As text
Dim failpoint As Long
On Error GoTo Err

If mDefaultTextSeries Is Nothing Then
    Set mDefaultTextSeries = AddGraphicObjectSeries(New TextSeries, LayerDefaultTexts)
    mDefaultTextSeries.style = mDefaultTextStyle
End If
Set AddText = mDefaultTextSeries.Add(, , IsSelectable)
If Layer <> -1 Then AddText.Layer = Layer

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "addText" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Friend Sub ClearRegion()
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

Set mGraphObjServiceProvider = Nothing
mCleared = True

Set mToolbar = Nothing

' now remove all graphic object references from the various collections
Dim graphObjSeries As IGraphicObjectSeries
For i = 0 To mGraphObjSeriesIndex - 1
    If Not mGraphObjSeriesTable(i) Is Nothing Then
        mGraphObjSeriesTable(i).Finish
        Set mGraphObjSeriesTable(i) = Nothing
    End If
Next
Erase mGraphObjSeriesTable
ReDim mGraphObjSeriesTable(3) As IGraphicObjectSeries

StartTask Me, PriorityLow

'' now clear the YAxis region (if there is one)
'If Not mYAxisRegion Is Nothing Then
'    mYAxisRegion.clearRegion
'End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "clearRegion" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub ClearSelection()
deSelectObjectOrSeries mSelectedObject
Set mSelectedObject = Nothing
End Sub

Friend Sub Click()
clickSelectedObjectOrSeries
RaiseEvent Click
End Sub

Friend Sub DblCLick()
dblClickSelectedObjectOrSeries
RaiseEvent DblCLick
End Sub

Public Sub DisplayGrid()
Dim failpoint As Long
On Error GoTo Err

If Not mHideGrid Then Exit Sub

mHideGrid = False
If mUseDeferredPainting Then
    deferredPaint DeferredPaintRepaintRegion
Else
    paint
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "displayGrid" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub DrawCursor(ByVal button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        Dim yVal As Single
Dim failpoint As Long
On Error GoTo Err

If y = MinusInfinitySingle Then
    yVal = mCanvas.Top + 2 * mCanvas.height  ' make sure the pointer isn't visible
Else
    yVal = y
End If

If Round(x) >= YAxisPosition Then Exit Sub

Dim lPointerStyle As PointerStyles

lPointerStyle = IIf(mPointerMode = PointerModeDefault, mPointerStyle, mPointerToolStyle)

Select Case lPointerStyle
Case PointerNone

Case PointerCrosshairs
    undrawCrosshairsPointer
    drawCrosshairsPointer x, yVal
Case PointerDisc
    undrawDiscPointer
    drawDiscPointer x, yVal
Case PointerTool
    undrawToolPointer
    drawToolPointer x, yVal
Case PointerCustom
    drawCustomPointer x, yVal
End Select

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "drawCursor" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription



End Sub

Public Function FormatYValue(ByVal y As Double) As String
FormatYValue = generateYAxisCursorPositionText(y)
End Function

Public Sub HideGrid()
Dim failpoint As Long
On Error GoTo Err

If mHideGrid Then Exit Sub

mHideGrid = True
If mUseDeferredPainting Then
    deferredPaint DeferredPaintRepaintRegion
Else
    paint
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "hideGrid" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Function IsGridHidden() As Boolean
IsGridHidden = mHideGrid
End Function

Friend Sub MouseDown(button As Integer, _
                Shift As Integer, _
                x As Single, _
                y As Single)
mCurrX = x
mCurrY = y
setSelectedObjectOrSeries
RaiseEvent MouseDown(button, Shift, x, y)
End Sub
                
Friend Sub mouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
mCurrX = x
mCurrY = y
RaiseEvent mouseMove(button, Shift, x, y)
End Sub

Friend Sub MouseUp(button As Integer, _
                Shift As Integer, _
                x As Single, _
                y As Single)
mCurrX = x
mCurrY = y
RaiseEvent MouseUp(button, Shift, x, y)
End Sub

Public Function newDimension(ByVal x As Double, _
                        ByVal y As Double) As Dimension
Dim failpoint As Long
On Error GoTo Err

Set newDimension = mCanvas.newDimension(x, y)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "newDimension" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function newPoint(ByVal x As Double, _
                        ByVal y As Double, _
                        Optional ByVal coordSystemX As CoordinateSystems = CoordsLogical, _
                        Optional ByVal coordSystemY As CoordinateSystems = CoordinateSystems.CoordsLogical) As Point
Dim failpoint As Long
On Error GoTo Err

Set newPoint = mCanvas.newPoint(x, y, coordSystemX, coordSystemY)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "newPoint" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Friend Sub PaintRegion()
If mUseDeferredPainting Then
    deferredPaint DeferredPaintRepaintRegion
Else
    paint
End If

If Not mYAxisRegion Is Nothing Then mYAxisRegion.PaintRegion
End Sub

Friend Sub PeriodsInView( _
                ByVal minValue As Single, _
                ByVal maxValue As Single)
Dim failpoint As Long
On Error GoTo Err

mMinInViewPeriod = minValue
mMaxInViewPeriod = maxValue
mYAxisPosition = maxValue + 1
mCanvas.Left = minValue
mCanvas.right = maxValue '+ 0.5

If Not mSuppressDrawing Then
    If mUseDeferredPainting Then
        deferredPaint DeferredPaintCalculateAndRepaint
    Else
        determineInScopeObjects
        If Autoscale Then
            calculateVerticalBounds
        End If
        determineVisibleObjects
        paint
        If Not mYAxisRegion Is Nothing Then mYAxisRegion.PaintRegion
    End If
End If
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "periodsInView" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub RemoveGraphicObject( _
                ByVal graphObj As IGraphicObject)
RemoveGraphicObjectEx graphObj, True, True
End Sub

Friend Sub RemoveGraphicObjectEx( _
                ByVal graphObj As IGraphicObject, _
                ByVal undraw As Boolean, _
                ByVal removeFromSeries As Boolean)
Dim Handle As Long

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

Handle = graphObj.Handle

failpoint = 200

If Not mSelectedObject Is Nothing Then
    If graphObj Is mSelectedObject Then ClearSelection
End If

failpoint = 300

If undraw Then undrawArea mGraphicObjects(Handle).boundingRect

If Not mGraphicObjects(Handle).ExtendedObject Then

    failpoint = 400

    removeEntryFromPeriodList Handle
Else

    failpoint = 500

    removeEntryFromExtendedList Handle
End If

If graphicObjectIsInScope(Handle) Then

failpoint = 600
    
    removeEntryFromInScopeList Handle
    If graphicObjectIsVisible(Handle) Then

        failpoint = 700

        removeEntryFromVisibleList Handle
    End If
End If

failpoint = 800

If removeFromSeries Then mGraphObjSeriesTable(mGraphicObjects(Handle).seriesId).RemoveObject graphObj

Set mGraphicObjects(Handle).graphObj = Nothing
graphObj.Handle = -1

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RemoveGraphicObjectEx" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Sub RemoveGraphicObjectSeries( _
                ByVal pSeries As IGraphicObjectSeries)
If pSeries Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "RemoveGraphicObjectSeries", _
            "The supplied argument is Nothing"
End If

Dim t As New GraphObjSeriesRemovalTask
t.initialise pSeries, Me
StartTask t, PriorityLow
Set mGraphObjSeriesTable(pSeries.seriesId) = Nothing
End Sub

Friend Sub resizedY()
' notifies the region that the physical height of the canvas has changed.
' This means the Y gauge will have changed, so we need to recalculate
' the position of any in-scope gauge-dependent objects.
Dim failpoint As Long
On Error GoTo Err

mCanvas.resizedY

If mUseDeferredPainting Then
    deferredPaint DeferredPaintChangeGaugeandRepaint
Else
    calculateInscopeGaugeDependentObjects
    paint
    If Not mYAxisRegion Is Nothing Then mYAxisRegion.resizedY
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "resizedY" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub ScaleUp(ByVal factor As Single)
Dim failpoint As Long
On Error GoTo Err

If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "scaleUp", _
            "Cannot scale an axis region"
End If

With mCanvas
    SetVerticalScale .Bottom + factor * .height / 2, _
                                .Top - factor * .height / 2
End With

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "scaleUp" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub ScrollVertical(ByVal amount As Single)
Dim failpoint As Long
On Error GoTo Err

If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "scrollVertical", _
            "Cannot scroll an axis region"
End If

Autoscale = False
SetVerticalScale mCanvas.Bottom + amount, mCanvas.Top + amount

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "scrollVertical" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub ScrollVerticalProportion(ByVal proportion As Single)
Dim failpoint As Long
On Error GoTo Err

If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "scrollVerticalProportion", _
            "Cannot scroll an axis region"
End If

Autoscale = False
SetVerticalScale mCanvas.Bottom + proportion * mCanvas.height, _
                mCanvas.Top + proportion * mCanvas.height

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "scrollVerticalProportion" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub SetPeriodBackgroundColor( _
                ByVal PeriodNumber As Long, _
                ByVal Color As Long)
Dim Layer As Long
Dim index As Long
Dim periodrect As TRectangle
Dim redrawAreas(0) As TRectangle

Dim failpoint As Long
On Error GoTo Err

mPeriodTable(PeriodNumber).BackColor = Color

If mSuppressDrawing Then Exit Sub


If Color = 0 Then Exit Sub

With mCanvas
    .setPenAttributes Color, 1, LineInsideSolid, DrawModeCopyPen
    .setBrushAttributes Color, FillSolid
    .drawRectangle .newPoint(PeriodNumber - 0.5, .Bottom), .newPoint(PeriodNumber + 0.5, .Top)

    
    periodrect.Left = PeriodNumber - 0.5
    periodrect.right = PeriodNumber + 0.5
    periodrect.Bottom = .Bottom
    periodrect.Top = .Top
    periodrect.isValid = True
    
    For Layer = MinLayer To MaxLayer
        index = mGraphicObjects(mVisibleTable(Layer).firstIndex).nextVisibleIndex
        Do While index <> mVisibleTable(Layer).lastindex
            With mGraphicObjects(index)
                redrawAreas(0) = rectIntersection(periodrect, .boundingRect)
                If redrawAreas(0).isValid Then
                    .graphObj.Draw redrawAreas
                End If
                index = .nextVisibleIndex
            End With
        Loop
    Next
End With

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "setPeriodBackgroundColor" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub SetPointerModeDefault()
If mPointerMode = PointerModeDefault Then Exit Sub

mPointerMode = PointerModeDefault

Select Case mPointerToolStyle
Case PointerNone

Case PointerCrosshairs
    undrawCrosshairsPointer
Case PointerDisc
    undrawDiscPointer
Case PointerTool
    undrawToolPointer
Case PointerCustom
    undrawCustomPointer
End Select

Select Case mPointerStyle
Case PointerNone
    mCanvas.MouseIcon = gBlankMouseIcon
Case PointerCrosshairs
    mCanvas.MouseIcon = gBlankMouseIcon
    drawCrosshairsPointer mPrevCursorX, mPrevCursorY
Case PointerDisc
    mCanvas.MouseIcon = gBlankMouseIcon
    drawDiscPointer mPrevCursorX, mPrevCursorY
Case PointerTool
    mCanvas.MouseIcon = gBlankMouseIcon
    drawToolPointer mPrevCursorX, mPrevCursorY
Case PointerCustom
    mCanvas.MouseIcon = mPointerIcon
    drawCustomPointer mPrevCursorX, mPrevCursorY
End Select

End Sub

Friend Sub SetPointerModeTool( _
                ByVal toolPointerStyle As PointerStyles, _
                ByVal icon As IPictureDisp)

If mPointerMode <> PointerModeTool Then
    Select Case mPointerStyle
    Case PointerNone
    
    Case PointerCrosshairs
        undrawCrosshairsPointer
    Case PointerDisc
        undrawDiscPointer
    Case PointerTool
        undrawToolPointer
    Case PointerCustom
        undrawCustomPointer
    End Select
Else
    Select Case mPointerToolStyle
    Case PointerNone
    
    Case PointerCrosshairs
        undrawCrosshairsPointer
    Case PointerDisc
        undrawDiscPointer
    Case PointerTool
        undrawToolPointer
    Case PointerCustom
        undrawCustomPointer
    End Select
End If

mPointerMode = PointerModeTool
mPointerToolStyle = toolPointerStyle
Set mPointerToolIcon = icon

Select Case mPointerToolStyle
Case PointerNone
    mCanvas.MouseIcon = gBlankMouseIcon
Case PointerCrosshairs
    mCanvas.MouseIcon = gBlankMouseIcon
    drawCrosshairsPointer mPrevCursorX, mPrevCursorY
Case PointerDisc
    mCanvas.MouseIcon = gBlankMouseIcon
    drawDiscPointer mPrevCursorX, mPrevCursorY
Case PointerTool
    mCanvas.MouseIcon = gBlankMouseIcon
    drawToolPointer mPrevCursorX, mPrevCursorY
Case PointerCustom
    mCanvas.MouseIcon = mPointerToolIcon
    drawCustomPointer mPrevCursorX, mPrevCursorY
End Select

End Sub

Public Sub SetVerticalScale(ByVal lowValue As Single, _
                            ByVal highValue As Single)

Dim failpoint As Long
On Error GoTo Err

If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "setVerticalScale", _
            "Cannot set the scale for an axis region"
End If

Autoscale = False

SetVerticalScaleInternal lowValue, highValue

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "setVerticalScale" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription



End Sub

Friend Sub SetVerticalScaleInternal(ByVal lowValue As Single, _
                            ByVal highValue As Single)

Dim failpoint As Long
On Error GoTo Err

mCanvas.Top = highValue
mCanvas.Bottom = lowValue

If Not mSuppressDrawing Then
    If mUseDeferredPainting Then
        deferredPaint DeferredPaintCalculateAndRepaint
    Else
        determineVisibleObjects
        paint
    End If
End If
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.SetVerticalScaleInternal lowValue, highValue
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "setVerticalScaleInternal" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub addEntryToExtendedList(ByVal index As Long)
With mGraphicObjects(index)
    If .nextExtendedIndex <> 0 Then Exit Sub
    .nextExtendedIndex = mLastExtendedIndex
    .prevExtendedIndex = mGraphicObjects(mLastExtendedIndex).prevExtendedIndex
    mGraphicObjects(.prevExtendedIndex).nextExtendedIndex = index
    mGraphicObjects(mLastExtendedIndex).prevExtendedIndex = index
End With
End Sub

Private Sub addEntryToInScopeList(ByVal index As Long)
With mGraphicObjects(index)
    If .inScopeCycle <> mInScopeCycle Then
        .graphObj.InScope = True
        .inScopeCycle = mInScopeCycle
        .nextInScopeIndex = mLastInScopeIndex
        .prevInScopeIndex = mGraphicObjects(mLastInScopeIndex).prevInScopeIndex
        mGraphicObjects(.prevInScopeIndex).nextInScopeIndex = index
        mGraphicObjects(mLastInScopeIndex).prevInScopeIndex = index
    End If
End With
End Sub

Private Sub addEntryToPeriodList( _
                ByVal index As Long)
Dim lastindex As Long
Dim newPeriodNumber As Long

newPeriodNumber = mGraphicObjects(index).graphObj.PeriodNumber
If newPeriodNumber <> mGraphicObjects(index).PeriodNumber Then
    If mGraphicObjects(index).PeriodNumber >= 1 Then
        removeEntryFromPeriodList index
    End If
    If newPeriodNumber < 1 Then Exit Sub
    
    ' NB: the following call to addPeriod must not be within a 'With' block
    ' that refers to mGraphicObjects, because in rare circumstances it can
    ' cause mGraphicObjects to be ReDim'd which causes an error 10
    addPeriod newPeriodNumber, _
            mGraphicObjects(index).graphObj.timestamp ' in case the period hasn't been added yet
    
    lastindex = mPeriodTable(newPeriodNumber).lastindex
    With mGraphicObjects(index)
        .PeriodNumber = newPeriodNumber
        .nextPeriodIndex = lastindex
        .prevPeriodIndex = mGraphicObjects(lastindex).prevPeriodIndex
        mGraphicObjects(.prevPeriodIndex).nextPeriodIndex = index
        mGraphicObjects(lastindex).prevPeriodIndex = index
    End With
End If

End Sub

Private Sub addEntryToVisibleList(ByVal index As Long)
Dim lastindex As Long
With mGraphicObjects(index)
    If .visibleCycle <> mVisibleCycle Then
        .graphObj.Visible = True
        .visibleCycle = mVisibleCycle
        lastindex = mVisibleTable(.Layer).lastindex
        .nextVisibleIndex = lastindex
        .prevVisibleIndex = mGraphicObjects(lastindex).prevVisibleIndex
        mGraphicObjects(.prevVisibleIndex).nextVisibleIndex = index
        mGraphicObjects(lastindex).prevVisibleIndex = index
    End If
End With
End Sub

Private Function AddGraphicObject( _
                ByVal graphObj As IGraphicObject, _
                ByVal seriesId As Long) As Long
Dim Handle As Long

Handle = allocateGraphicObjectTableEntry
If mFirstGraphicObjectIndex = 0 Then
    mFirstGraphicObjectIndex = Handle
End If

graphObj.Handle = Handle

With mGraphicObjects(Handle)
    Set .graphObj = graphObj
    .seriesId = seriesId
    .boundingRect = graphObj.BoundingRectangle
    .ExtendedObject = graphObj.ExtendedObject
    .GaugeDependent = graphObj.GaugeDependent
    .includeInAutoscale = graphObj.includeInAutoscale
    .Layer = graphObj.Layer
    .ScaleDependent = graphObj.ScaleDependent
    .Capabilities = graphObj.Capabilities
    .NoDraw = graphObj.NoDraw
    
End With

If Not mGraphicObjects(Handle).ExtendedObject Then
    ' NB: the following call to addEntryToPeriodList must not be within a With block
    ' that refers to mGraphicObjects, because in rare circumstances it can
    ' cause mGraphicObjects to be ReDim'd which causes an error 10
    addEntryToPeriodList Handle
Else
    addEntryToExtendedList Handle
End If

If graphicObjectIsInScope(Handle) Then
    addEntryToInScopeList Handle
    If graphicObjectIsVisible(Handle) Then
        addEntryToVisibleList Handle
    End If
End If
AddGraphicObject = Handle
End Function

Private Sub addNewPeriod(ByVal PeriodNumber As Long)
Dim i As Long

If PeriodNumber > UBound(mPeriodTable) Then
    ReDim Preserve mPeriodTable(PeriodNumber + 200) As PeriodTableEntry
End If
If PeriodNumber > mMaxPeriodNumber Then
    For i = mMaxPeriodNumber + 1 To PeriodNumber
        initialisePeriodTableEntry i
    Next
    mMaxPeriodNumber = PeriodNumber
Else
    initialisePeriodTableEntry PeriodNumber
End If

End Sub

Private Sub addVerticalGridLine(ByVal PeriodNumber As Long, _
                                ByVal Key As String, _
                                ByVal label As String, _
                                ByVal coarseKey As String, _
                                ByVal coarseLabel As String)
Dim vGridLine As ChartSkil26.Line
Dim vGridText As text
Dim vGridTextCoarse As text

If HasGrid Then
    On Error Resume Next
    Set vGridLine = mXGridLines.Item(Key)
    On Error GoTo 0
    If vGridLine Is Nothing Then
        Set vGridLine = mXGridLines.Add(Key)
        vGridLine.SetPosition newPoint(PeriodNumber - 0.5, _
                                    0, _
                                    CoordsLogical, _
                                    CoordsRelative), _
                                newPoint(PeriodNumber - 0.5, _
                                    100, _
                                    CoordsLogical, _
                                    CoordsRelative)
    Else
        If vGridLine.Point1.XLogical > PeriodNumber Then
            vGridLine.SetPosition newPoint(PeriodNumber - 0.5, _
                                        0, _
                                        CoordsLogical, _
                                        CoordsRelative), _
                                    newPoint(PeriodNumber - 0.5, _
                                        100, _
                                        CoordsLogical, _
                                        CoordsRelative)
        End If
    End If
End If

If HasGridText Then
    On Error Resume Next
    Set vGridText = mXGridTexts.Item(Key)
    On Error GoTo 0
    If vGridText Is Nothing Then
        Set vGridText = mXGridTexts.Add(Key)
        vGridText.fixedX = False
        vGridText.fixedY = True
        vGridText.position = newPoint(PeriodNumber - 0.5, _
                                    0#, _
                                    CoordsLogical, _
                                    CoordsDistance)
        vGridText.text = label
    End If

    On Error Resume Next
    Set vGridTextCoarse = mXGridTextsCoarse.Item(coarseKey)
    On Error GoTo 0
    If vGridTextCoarse Is Nothing Then
        Set vGridTextCoarse = mXGridTextsCoarse.Add(coarseKey)
        vGridTextCoarse.fixedX = False
        vGridTextCoarse.fixedY = True
        vGridTextCoarse.position = newPoint(PeriodNumber - 0.5, _
                                    0#, _
                                    CoordsLogical, _
                                    CoordsCounterDistance)
        vGridTextCoarse.text = coarseLabel
    End If
End If

End Sub

Private Function allocateGraphicObjectTableEntry() As Long
On Error GoTo Err
allocateGraphicObjectTableEntry = mNextGraphicObjectIndex

If mNextGraphicObjectIndex > UBound(mGraphicObjects) Then
    ReDim Preserve mGraphicObjects(UBound(mGraphicObjects) + 500) As GraphicObjectTableEntry
End If
mNextGraphicObjectIndex = mNextGraphicObjectIndex + 1
Exit Function
Err:
Stop
End Function

Private Function calculateInscopeGaugeDependentObjects()
Dim index As Long
index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
Do While index <> mLastInScopeIndex
     With mGraphicObjects(index)
        If .GaugeDependent Then
            .boundingRect = .graphObj.BoundingRectangle
            If graphicObjectIsVisible(index) Then
                addEntryToVisibleList index
            End If
        End If
        index = .nextInScopeIndex
    End With
Loop

End Function

Private Function calculateMinutesFromMidnight( _
                            ByVal timestamp As Date) As Long
calculateMinutesFromMidnight = Int((timestamp - Int(timestamp) + 1 / 86400000) * 1440)
End Function

Private Sub calculateObjectHighAndLowInInterval( _
                            ByVal index As Long, _
                            ByRef interval As TInterval, _
                            ByRef high As Double, _
                            ByRef low As Double)

Dim j As Long
Dim k As Long
Dim rects() As TRectangle
Dim numrects As Long
Dim overlapXInterval As TInterval

high = MinusInfinityDouble
low = PlusInfinityDouble

With mGraphicObjects(index)
    If (.Capabilities And GraphicObjectCapabilities.ComplexBounding) Then
        overlapXInterval = intIntersection(interval, rectGetXInterval(.boundingRect))
        For j = Int(overlapXInterval.startValue + 0.9) To _
                Int(overlapXInterval.endValue)
            rects = .graphObj.BoundingRectanglesAt(j)
            numrects = -1
            On Error Resume Next
            numrects = UBound(rects)
            On Error GoTo 0
            If numrects >= 0 Then
                For k = 0 To numrects
                    If rects(k).Top > high Then
                        high = rects(k).Top
                    End If
                    If rects(k).Bottom < low Then
                        low = rects(k).Bottom
                    End If
                Next
            End If
        Next
    Else
        high = .boundingRect.Top
        low = .boundingRect.Bottom
    End If
End With
End Sub

Private Function calculateSecondsFromMidnight( _
                            ByVal timestamp As Date) As Long
calculateSecondsFromMidnight = Int((timestamp - Int(timestamp) + 1 / 86400000) * 86400)
End Function

Private Sub calculateVerticalBounds()
Dim objMax As Double
Dim objMin As Double
Dim scaleMax As Double
Dim scaleMin As Double
Dim index As Long
Dim someGaugeDependentObjects As Boolean

If Not Autoscale Then Exit Sub

scaleMax = MinusInfinityDouble
scaleMin = PlusInfinityDouble

index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
Do While index <> mLastInScopeIndex
    With mGraphicObjects(index)
        If Not .graphObj.NoDraw And _
            .includeInAutoscale And _
            Not .ScaleDependent _
        Then
            If .GaugeDependent Then
                someGaugeDependentObjects = True
            Else
                calculateObjectHighAndLowInInterval index, _
                                                rectGetXInterval(mCanvas.boundary), _
                                                objMax, _
                                                objMin
                If objMax > scaleMax Then
                    scaleMax = objMax
                End If
                If objMin < scaleMin Then
                    scaleMin = objMin
                End If
            End If
        End If
        
        index = .nextInScopeIndex
    End With
Loop

With mCanvas
    If scaleMin > scaleMax Then
        ' no autoscale relevant in-view graphic objects -
        ' so leave default scale
    ElseIf MinimumHeight <> 0 Then
        If scaleMax - scaleMin < MinimumHeight Then
            .Top = (scaleMax + scaleMin) / 2 + MinimumHeight / 2
            .Bottom = .Top - MinimumHeight
        Else
            .Top = scaleMax + (scaleMax - scaleMin) * 0.025
            .Bottom = scaleMin - (scaleMax - scaleMin) * 0.025
        End If
    Else
        If scaleMin = scaleMax Then
            scaleMax = scaleMax * 1.001
            scaleMin = scaleMax * 0.998
        End If
        .Top = scaleMax + (scaleMax - scaleMin) * 0.025
        .Bottom = scaleMin - (scaleMax - scaleMin) * 0.025
    End If
    If Not mYAxisRegion Is Nothing Then
        mYAxisRegion.Top = .Top
        mYAxisRegion.Bottom = .Bottom
    End If
End With

If Not someGaugeDependentObjects Then
    Exit Sub
End If

' there are gauge-dependent objects that may affect the max and min scales.
' We need to recalculate their bounds with the new gauge determined from
' the other objects to get a better estimate of their sizes.
index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
Do While index <> mLastInScopeIndex
    With mGraphicObjects(index)
        If Not .graphObj.NoDraw And _
            .includeInAutoscale And _
            Not .ScaleDependent And _
            .GaugeDependent _
        Then
            .boundingRect = .graphObj.BoundingRectangle
            calculateObjectHighAndLowInInterval index, _
                                            rectGetXInterval(mCanvas.boundary), _
                                            objMax, _
                                            objMin
            If objMax > scaleMax Then
                scaleMax = objMax
            End If
            If objMin < scaleMin Then
                scaleMin = objMin
            End If
        End If
        
        index = .nextInScopeIndex
    End With
Loop

With mCanvas
    .Top = scaleMax + (scaleMax - scaleMin) * 0.025
    .Bottom = scaleMin - (scaleMax - scaleMin) * 0.025
    If Not mYAxisRegion Is Nothing Then
        mYAxisRegion.Top = .Top
        mYAxisRegion.Bottom = .Bottom
    End If
End With
End Sub

Private Sub clearInScopeObjectList()
Dim index As Long

index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
Do While index <> mLastInScopeIndex
    With mGraphicObjects(index)
        .graphObj.InScope = False
        index = .nextInScopeIndex
    End With
Loop

initialiseInScopeObjectList
End Sub

Private Sub clearVisibleObjectList()
Dim i As Long
Dim index As Long

For i = MinLayer To MaxLayer
    index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(i).lastindex
        With mGraphicObjects(index)
            .graphObj.Visible = False
            index = .nextVisibleIndex
        End With
    Loop
Next

initialiseVisibleObjectList
End Sub

Private Function clickSelectedObjectOrSeries() As Boolean
If mSelectedObject Is Nothing Then
    clickSelectedObjectOrSeries = False
ElseIf mSelectedObject.IsSelectable Then
    mSelectedObject.Click
    clickSelectedObjectOrSeries = True
ElseIf mGraphObjSeriesTable(mSelectedObject.seriesId).IsSelectable Then
    mGraphObjSeriesTable(mSelectedObject.seriesId).Click
    clickSelectedObjectOrSeries = True
Else
    clickSelectedObjectOrSeries = False
End If
End Function

Private Function dblClickSelectedObjectOrSeries() As Boolean
If mSelectedObject Is Nothing Then
    dblClickSelectedObjectOrSeries = False
ElseIf mSelectedObject.IsSelectable Then
    mSelectedObject.DblCLick
    dblClickSelectedObjectOrSeries = True
ElseIf mGraphObjSeriesTable(mSelectedObject.seriesId).IsSelectable Then
    mGraphObjSeriesTable(mSelectedObject.seriesId).DblCLick
    dblClickSelectedObjectOrSeries = True
Else
    dblClickSelectedObjectOrSeries = False
End If
End Function

Private Function deSelectObjectOrSeries(ByVal graphObj As IGraphicObject) As Boolean
If graphObj Is Nothing Then
ElseIf graphObj.IsSelectable Then
    graphObj.Selected = False
    deSelectObjectOrSeries = True
ElseIf mGraphObjSeriesTable(graphObj.seriesId).IsSelectable Then
    mGraphObjSeriesTable(graphObj.seriesId).Selected = False
    deSelectObjectOrSeries = True
Else
    deSelectObjectOrSeries = False
End If
End Function

Private Sub deferredPaint(ByVal action As DeferredPaintActions)
If action = DeferredPaintRepaintNone Then
    mDeferredPaintAction = DeferredPaintRepaintNone
    mDeferredPaintPending = False
    ReDim mUndrawnAreas(3) As TRectangle
    mUndrawnAreasIndex = -1
    ReDim mRedrawObjects(3) As RedrawObjectsTableEntry
    mRedrawObjectsIndex = -1
Else
    Select Case mDeferredPaintAction
    Case DeferredPaintRepaintNone
        mDeferredPaintAction = action
    Case DeferredPaintRepaintChangedAreas
        If action = DeferredPaintRepaintRegion Or _
            action = DeferredPaintChangeGaugeandRepaint Or _
            action = DeferredPaintScaleAndRepaint Or _
            action = DeferredPaintCalculateAndRepaint _
        Then mDeferredPaintAction = action
    Case DeferredPaintRepaintRegion
        If action = DeferredPaintChangeGaugeandRepaint Or _
            action = DeferredPaintScaleAndRepaint Or _
            action = DeferredPaintCalculateAndRepaint _
        Then mDeferredPaintAction = action
    Case DeferredPaintChangeGaugeandRepaint
        If action = DeferredPaintScaleAndRepaint Or _
            action = DeferredPaintCalculateAndRepaint _
        Then mDeferredPaintAction = action
    Case DeferredPaintScaleAndRepaint
        If action = DeferredPaintCalculateAndRepaint Then mDeferredPaintAction = action
    Case DeferredPaintCalculateAndRepaint
        
    End Select
    If Not mDeferredPaintPending Then
        mDeferredPaintPending = True
        DeferAction Me, mRegionNumber, 20, ExpiryTimeUnitMilliseconds
    End If
End If
End Sub

Private Sub determineInScopeObjects()
Dim i As Long
Dim index As Long
Dim timerIn As Single
Dim numberExamined As Long
Dim numberExtended As Long
Dim numberInScope As Long
Dim minPeriod As Long
Dim maxPeriod As Long

timerIn = Timer

clearInScopeObjectList

minPeriod = IIf(mMinInViewPeriod < 1, 1, mMinInViewPeriod)
maxPeriod = IIf(mMaxPeriodNumber < mMaxInViewPeriod, mMaxPeriodNumber, mMaxInViewPeriod)
For i = minPeriod To maxPeriod
    index = mGraphicObjects(mPeriodTable(i).firstIndex).nextPeriodIndex
    Do While index <> 0
        numberExamined = numberExamined + 1
        With mGraphicObjects(index)
            If .boundingRect.isValid Then
                addEntryToInScopeList index
                numberInScope = numberInScope + 1
            End If
            index = .nextPeriodIndex
        End With
    Loop
Next

Debug.Print "DetermineInScopeObjects: nonextended=" & numberExamined & " time=" & 1000 * (Timer - timerIn) & "ms"
timerIn = Timer

index = mGraphicObjects(mFirstExtendedIndex).nextExtendedIndex
Do While index <> mLastExtendedIndex
    numberExamined = numberExamined + 1
    numberExtended = numberExtended + 1
    With mGraphicObjects(index)
        .boundingRect = .graphObj.BoundingRectangle
        If graphicObjectIsInScope(index) Then
            addEntryToInScopeList index
            numberInScope = numberInScope + 1
        End If
        index = .nextExtendedIndex
    End With
Loop

Debug.Print "DetermineInScopeObjects: extended=" & numberExtended & " time=" & 1000 * (Timer - timerIn) & "ms"
Debug.Print "DetermineInScopeObjects: total=" & numberExamined & " in-scope=" & numberInScope
End Sub

Private Sub determineVisibleObjects()
Dim index As Long
Dim timerIn As Single
Dim numberExamined As Long
Dim numberVisible As Long

timerIn = Timer

clearVisibleObjectList

index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
Do While index <> mLastInScopeIndex
    numberExamined = numberExamined + 1
    With mGraphicObjects(index)
        .boundingRect = .graphObj.BoundingRectangle
        If graphicObjectIsVisible(index) Then
            addEntryToVisibleList index
            numberVisible = numberVisible + 1
        End If
        index = .nextInScopeIndex
    End With
Loop

Debug.Print "DetermineVisibleObjects: total=" & numberExamined & " visible=" & numberVisible & " time=" & 1000 * (Timer - timerIn) & "ms"
End Sub

Private Sub drawCrosshairsPointer(ByVal x As Single, ByVal y As Single)
Dim lineX As Single

If mSuppressDrawing Then Exit Sub

With mCanvas
    .setPenAttributes mPointerCrosshairsColor Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .drawLine .newPoint(.Left, y), .newPoint(YAxisPosition, y)
    lineX = Round(x)
    .drawLine .newPoint(lineX, .Bottom), .newPoint(lineX, .Top)
    mPrevCursorX = x
    mPrevCursorY = y
    
End With

If Not mYAxisRegion Is Nothing Then
    mYPositionText.position = mYAxisRegion.newPoint(15, y, CoordsLogical, CoordsLogical)
    mYPositionText.text = generateYAxisCursorPositionText(y)
End If
End Sub

Private Sub drawCustomPointer(ByVal x As Single, ByVal y As Single)
Dim lineX As Single
Dim size As Dimension

If mSuppressDrawing Then Exit Sub

mPrevCursorX = x
mPrevCursorY = y
    
If Not mYAxisRegion Is Nothing Then
    mYPositionText.position = mYAxisRegion.newPoint(15, y, CoordsLogical, CoordsLogical)
    mYPositionText.text = generateYAxisCursorPositionText(y)
End If
End Sub

Private Sub drawDiscPointer(ByVal x As Single, ByVal y As Single)
Dim lineX As Single
Dim size As Dimension

If mSuppressDrawing Then Exit Sub

With mCanvas
    '.MousePointer = vbCrosshair
    lineX = Round(x)
    Set size = .newDimension(0.4, 0.4)
    
    .setPenAttributes vbBlack Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .drawLine .newPoint(lineX - size.XLogical, y), .newPoint(lineX + size.XLogical, y)
    .drawLine .newPoint(lineX, y - size.YLogical), .newPoint(lineX, y + size.YLogical)
    
    .setPenAttributes mPointerDiscColor Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .setBrushAttributes mPointerDiscColor Xor .BackColor, FillStyles.FillSolid
    .drawCircle .newPoint(lineX, y), .newDimension(0.8, 0.8).XLogical
    
    mPrevCursorX = x
    mPrevCursorY = y
    
End With

If Not mYAxisRegion Is Nothing Then
    mYPositionText.position = mYAxisRegion.newPoint(15, y, CoordsLogical, CoordsLogical)
    mYPositionText.text = generateYAxisCursorPositionText(y)
End If
End Sub

Private Sub drawToolPointer(ByVal x As Single, ByVal y As Single)
Dim lineX As Single
Dim size As Dimension

If mSuppressDrawing Then Exit Sub

With mCanvas
    lineX = Round(x)
    Set size = .newDimension(0.4, 0.4)
    
    .setPenAttributes vbBlack Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .drawLine .newPoint(lineX - size.XLogical, y), .newPoint(lineX + size.XLogical, y)
    .drawLine .newPoint(lineX, y - size.YLogical), .newPoint(lineX, y + size.YLogical)
    
    mPrevCursorX = x
    mPrevCursorY = y
    
End With

If Not mYAxisRegion Is Nothing Then
    mYPositionText.position = mYAxisRegion.newPoint(15, y, CoordsLogical, CoordsLogical)
    mYPositionText.text = generateYAxisCursorPositionText(y)
End If
End Sub

Private Sub drawObject( _
                ByVal graphObj As IGraphicObject, _
                ByRef areaToRedraw As TRectangle)
Dim drawAreas(0) As TRectangle
drawAreas(0) = mCanvas.boundary

mCanvas.setClippingRegion areaToRedraw
graphObj.Draw drawAreas
mCanvas.clearClippingRegion
End Sub

Private Function generateYAxisCursorPositionText(ByVal y As Single) As String
Dim fract As Single
Dim num32nds As Long

If YScaleQuantum = 0.03125 Then
    fract = y - Int(y)
    num32nds = Int(fract / 0.03125)
    If mScaleGridSpacingY > 4 * 0.03125 Then
        generateYAxisCursorPositionText = Int(y) & " ' " & Format(num32nds, "00")
    Else
        generateYAxisCursorPositionText = Int(y) & " ' " & Format(num32nds, "00") & "." & Format(100 * (fract - num32nds * 0.03125) / 0.03125, "00")
    End If
Else
    generateYAxisCursorPositionText = Format(y, mYScaleSubFormatStr)
End If
End Function

Private Function getLayer(Optional ByVal Layer As Long = -1) As Long
If Layer = -1 Then
    getLayer = mNextLayer
    If mNextLayer <> LayerNumbers.LayerHighestUser Then
        mNextLayer = mNextLayer + 1
    End If
Else
    If Layer < MinLayer Or Layer > MaxLayer Then
        getLayer = getLayer()
    Else
        getLayer = Layer
    End If
End If
End Function

Private Function getObjectOverlaps(ByVal index1 As Long, _
                                ByVal index2 As Long, _
                                ByRef clippingRect As TRectangle, _
                                ByRef useComplexBounding1 As Boolean, _
                                ByRef useComplexBounding2 As Boolean, _
                                ByRef overlapRects() As TRectangle) As Long
Dim overlaps() As TRectangle
Dim graphobj1 As IGraphicObject
Dim graphobj2 As IGraphicObject
Dim obj1Rect As TRectangle
Dim obj2Rect As TRectangle
Dim XInterval As TInterval
Dim complexBounding1 As Boolean
Dim complexBounding2 As Boolean
Dim boundingRects1() As TRectangle
Dim numRects1 As Long
Dim boundingRects2() As TRectangle
Dim numRects2 As Long
Dim i As Long
Dim j As Long
Dim k As Long

With mGraphicObjects(index1)
    Set graphobj1 = .graphObj
    obj1Rect = rectIntersection(.boundingRect, clippingRect)
    complexBounding1 = (.Capabilities And GraphicObjectCapabilities.ComplexBounding)
End With

With mGraphicObjects(index2)
    Set graphobj2 = .graphObj
    obj2Rect = rectIntersection(.boundingRect, clippingRect)
    complexBounding2 = (.Capabilities And GraphicObjectCapabilities.ComplexBounding)
End With

ReDim overlaps(0) As TRectangle
overlaps(0) = rectIntersection(obj1Rect, obj2Rect)
If Not overlaps(0).isValid Then
    getObjectOverlaps = 0
    Exit Function
End If

If Not (useComplexBounding1 Or useComplexBounding2) Or _
    Not (complexBounding1 Or complexBounding2) Or _
    (complexBounding1 And Not useComplexBounding1) Or _
    (complexBounding2 And Not useComplexBounding2) _
Then
    overlapRects = overlaps
    getObjectOverlaps = 1
    Exit Function
End If

XInterval = rectGetXInterval(rectIntersection(obj1Rect, obj2Rect))

If complexBounding1 And complexBounding2 And useComplexBounding1 And useComplexBounding2 Then
    ReDim overlaps(5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
            ' size the array so it is unlikely to need to be extended
            
    For i = Int(XInterval.startValue) To _
            Int(XInterval.endValue)
        
        boundingRects1 = graphobj1.BoundingRectanglesAt(i)
        numRects1 = -1
        On Error Resume Next
        numRects1 = UBound(boundingRects1) + 1
        On Error GoTo 0
        
        boundingRects2 = graphobj2.BoundingRectanglesAt(i)
        numRects2 = -1
        On Error Resume Next
        numRects2 = UBound(boundingRects2) + 1
        On Error GoTo 0
        
        If numRects1 > 0 And numRects2 > 0 Then
            For j = 0 To UBound(boundingRects1)
                If boundingRects1(j).isValid Then
                    For k = 0 To UBound(boundingRects2)
                        If boundingRects2(k).isValid Then
                            If getObjectOverlaps > UBound(overlaps) Then
                                ReDim Preserve overlaps(UBound(overlaps) + 5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
                            End If
                            overlaps(getObjectOverlaps) = _
                                rectIntersection(boundingRects1(j), boundingRects2(k))
                            If overlaps(getObjectOverlaps).isValid Then
                                getObjectOverlaps = getObjectOverlaps + 1
                            End If
                        End If
                    Next
                End If
            Next
        End If
    If getObjectOverlaps > 0 Then
        ReDim Preserve overlaps(getObjectOverlaps - 1) As TRectangle
        overlapRects = overlaps
    End If
    Next
ElseIf complexBounding1 Then
    ReDim overlaps(5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
            ' size the array so it is unlikely to need to be extended
    
    For i = Int(XInterval.startValue) To _
            Int(XInterval.endValue)
        
        boundingRects1 = graphobj1.BoundingRectanglesAt(i)
        numRects1 = -1
        On Error Resume Next
        numRects1 = UBound(boundingRects1) + 1
        On Error GoTo 0
        
        If numRects1 > 0 Then
            For j = 0 To UBound(boundingRects1)
                If getObjectOverlaps > UBound(overlaps) Then
                    ReDim Preserve overlaps(UBound(overlaps) + 5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
                End If
                overlaps(getObjectOverlaps) = rectIntersection(boundingRects1(j), obj2Rect)
                If overlaps(getObjectOverlaps).isValid Then
                    getObjectOverlaps = getObjectOverlaps + 1
                End If
            Next
        End If
    Next
    If getObjectOverlaps > 0 Then
        ReDim Preserve overlaps(getObjectOverlaps - 1) As TRectangle
        overlapRects = overlaps
    End If
Else
    ReDim overlaps(5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
            ' size the array so it is unlikely to need to be extended
    
    For i = Int(XInterval.startValue) To _
            Int(XInterval.endValue)
        
        boundingRects2 = graphobj2.BoundingRectanglesAt(i)
        numRects2 = -1
        On Error Resume Next
        numRects2 = UBound(boundingRects2) + 1
        On Error GoTo 0
        
        If numRects2 > 0 Then
            For j = 0 To UBound(boundingRects2)
                If getObjectOverlaps > UBound(overlaps) Then
                    ReDim Preserve overlaps(UBound(overlaps) + 5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
                End If
                overlaps(getObjectOverlaps) = rectIntersection(boundingRects2(j), obj1Rect)
                If overlaps(getObjectOverlaps).isValid Then
                    getObjectOverlaps = getObjectOverlaps + 1
                End If
            Next
        End If
    Next
    If getObjectOverlaps > 0 Then
        ReDim Preserve overlaps(getObjectOverlaps - 1) As TRectangle
        overlapRects = overlaps
    End If
End If

End Function

Private Function graphicObjectIsInScope(ByVal index As Long) As Boolean
With mGraphicObjects(index)
    If Not .boundingRect.isValid Then Exit Function
    If .ExtendedObject Then
        graphicObjectIsInScope = intOverlaps(rectGetXInterval(mCanvas.boundary), rectGetXInterval(.boundingRect))
    Else
        graphicObjectIsInScope = (.PeriodNumber >= mMinInViewPeriod And _
                                .PeriodNumber <= mMaxInViewPeriod)
    End If
End With
End Function

Private Function graphicObjectIsVisible(ByVal index As Long) As Boolean
Dim boundingRects() As TRectangle
Dim numrects As Long
Dim overlapXInterval As TInterval
Dim i As Long
Dim j As Long

With mGraphicObjects(index)
    If Not .boundingRect.isValid Or .NoDraw Then Exit Function
    If .ExtendedObject And (.Capabilities And GraphicObjectCapabilities.ComplexBounding) Then
        overlapXInterval = intIntersection(rectGetXInterval(mCanvas.boundary), rectGetXInterval(.boundingRect))
        For i = Int(overlapXInterval.startValue) To _
                Int(overlapXInterval.endValue)
            boundingRects = .graphObj.BoundingRectanglesAt(i)
            numrects = -1
            On Error Resume Next
            numrects = UBound(boundingRects)
            On Error GoTo 0
            If numrects >= 0 Then
                For j = 0 To numrects
                    If rectOverlaps(boundingRects(j), mCanvas.boundary) Then
                        graphicObjectIsVisible = True
                        Exit Function
                    End If
                Next
            End If
        Next
    Else
        graphicObjectIsVisible = rectOverlaps(.boundingRect, mCanvas.boundary)
    End If
End With
End Function

Private Sub initialiseInScopeObjectList()
mInScopeCycle = mInScopeCycle + 1
mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex = mLastInScopeIndex
mGraphicObjects(mLastInScopeIndex).prevInScopeIndex = mFirstInScopeIndex
End Sub

Private Sub initialisePeriodTableEntry( _
                ByVal index As Long)
'With mPeriodTable(index)
    If mPeriodTable(index).firstIndex = 0 Then
        mPeriodTable(index).firstIndex = allocateGraphicObjectTableEntry
        mPeriodTable(index).lastindex = allocateGraphicObjectTableEntry
        mGraphicObjects(mPeriodTable(index).firstIndex).nextPeriodIndex = mPeriodTable(index).lastindex
        mGraphicObjects(mPeriodTable(index).lastindex).prevPeriodIndex = mPeriodTable(index).firstIndex
    End If
'End With
End Sub

Private Sub initialiseVisibleObjectList()
Dim i As Long

mVisibleCycle = mVisibleCycle + 1

For i = MinLayer To MaxLayer
    mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex = mVisibleTable(i).lastindex
    mGraphicObjects(mVisibleTable(i).lastindex).prevVisibleIndex = mVisibleTable(i).firstIndex
Next

End Sub

Private Function isObjectOrSeriesSelectable(ByVal graphObj As IGraphicObject) As Boolean
If graphObj.IsSelectable Then
    isObjectOrSeriesSelectable = True
ElseIf mGraphObjSeriesTable(graphObj.seriesId).IsSelectable Then
    isObjectOrSeriesSelectable = True
Else
    isObjectOrSeriesSelectable = False
End If
End Function

Private Sub paint()
Dim numGridDivisions As Double
Dim gridSpacing As Double
Dim i As Long
Dim z As Double
Dim perfFont As StdFont
Dim counter As Long
Dim index As Long

Dim failpoint As Long
On Error GoTo Err

Static repaintCount As Long
repaintCount = repaintCount + 1
Debug.Print "Region " & RegionNumber & " repaint count = " & repaintCount

With mCanvas

If mShowPerformanceText Then mElapsedTimer.StartTiming

For i = IIf(mMinInViewPeriod < 0, 0, mMinInViewPeriod) To IIf(mMaxPeriodNumber < mMaxInViewPeriod, mMaxPeriodNumber, mMaxInViewPeriod)
    If mPeriodTable(i).BackColor <> 0 Then
        .setPenAttributes mPeriodTable(i).BackColor, 1, LineInsideSolid, DrawModeCopyPen
        .setBrushAttributes mPeriodTable(i).BackColor, FillSolid
        .drawRectangle .newPoint(i - 0.5, .Bottom), .newPoint(i + 0.5, .Top)
    End If
Next

numGridDivisions = .heightCm / GridlineSpacingY
gridSpacing = .height / numGridDivisions

mGridTextHeight = .TextHeight("123")

Dim log10 As Double
Dim logSpacing As Double
Dim exp As Long
Dim mantissa As Double

If HasGrid And Not mHideGrid Then
    log10 = Log(10#)
    logSpacing = Log(gridSpacing) / log10
    
    exp = Fix(logSpacing)
    mantissa = Abs(logSpacing) - Abs(exp)
    
    If logSpacing < 0 Then
        exp = exp - 1
        mantissa = 1 - mantissa
    End If
    
    If YScaleQuantum = 0.03125 And _
        exp < 0 _
    Then
        ' tick size is 1/32, which influences where we put the horizontal grid lines
        If exp >= -1 Then
            If mantissa >= Log(5) / log10 Then
                mScaleGridSpacingY = 16 * 0.03125
            ElseIf mantissa >= Log(2.5) / log10 Then
                mScaleGridSpacingY = 8 * 0.03125
            ElseIf mantissa >= Log(1.25) / log10 Then
                mScaleGridSpacingY = 4 * 0.03125
            ElseIf mantissa >= Log(0.625) / log10 Then
                mScaleGridSpacingY = 2 * 0.03125
            Else
                mScaleGridSpacingY = 0.03125
            End If
        Else
            mScaleGridSpacingY = 0.03125
        End If
    Else
        If mantissa >= Log(5) / log10 Then
            mScaleGridSpacingY = 5 * 10 ^ exp
        ElseIf mantissa >= Log(2.5) / log10 Then
            mScaleGridSpacingY = 2.5 * 10 ^ exp
        ElseIf mantissa >= Log(2#) / log10 Then
            mScaleGridSpacingY = 2# * 10 ^ exp
        Else
            mScaleGridSpacingY = 10 ^ exp
        End If
    End If
    If IntegerYScale Then
            mYScaleFormatStr = "0"
            mYScaleSubFormatStr = "0"
    Else
        If mScaleGridSpacingY < 1 Then
            mYScaleFormatStr = "0." & String(Abs(exp) + 1, "0")
            mYScaleSubFormatStr = "0." & String(Abs(exp) + 2, "0")
        Else
            mYScaleFormatStr = "0.0"
            mYScaleSubFormatStr = "0.00"
        End If
    End If
End If

mCanvas.paintBackground

If HasGrid And Not mHideGrid Then
    paintHorizontalGridLines
End If

setHorizontalGridTexts

For i = MinLayer To MaxLayer
    index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(i).lastindex
        With mGraphicObjects(index)
            Dim drawAreas(0) As TRectangle
            drawAreas(0) = rectIntersection(.boundingRect, mCanvas.boundary)
            .graphObj.Draw drawAreas
            index = .nextVisibleIndex
            counter = counter + 1
        End With
    Loop
Next

redrawCursor

If mShowPerformanceText Then
    If mPerfText Is Nothing Then
        Set perfFont = New StdFont
        perfFont.name = "Lucida Console"
        perfFont.size = 8
        perfFont.Underline = False
        perfFont.Bold = False
        Set mPerfText = AddText()
        mPerfText.Color = vbActiveTitleBarText
        mPerfText.font = perfFont
        mPerfText.box = True
        mPerfText.boxColor = vbBlack
        mPerfText.boxStyle = LineStyles.LineInsideSolid
        mPerfText.boxThickness = 1
        mPerfText.boxFillColor = vbActiveTitleBar
        mPerfText.boxFillStyle = FillStyles.FillSolid
        mPerfText.position = newPoint(5, 2, CoordsRelative, CoordsRelative)
        mPerfText.fixedX = True
        mPerfText.fixedY = True
        mPerfText.align = TextAlignModes.AlignBottomLeft
        mPerfText.includeInAutoscale = False
        mPerfText.paddingX = 0.5
        mPerfText.Layer = LayerTitle
        
    End If
    mPerfText.text = Format(mElapsedTimer.ElapsedTimeMicroseconds / 1000, "0.0") & "ms" & vbCrLf & _
                    counter & " visible objects" & vbCrLf & _
                    mNextGraphicObjectIndex & " total objects"
    ' need to force the text to draw now
    mPerfText.Draw
End If

End With

mDrawn = True

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "paint" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Private Sub paintHorizontalGridLines()
Dim y As Double
Dim i As Long
Dim aLine As ChartSkil26.Line

If mYGridLines Is Nothing Then Exit Sub

y = Int(mCanvas.Bottom / mScaleGridSpacingY) * mScaleGridSpacingY
Do While y < mCanvas.Top
    i = i + 1
    If i > mYGridLines.Count Then
        Set aLine = mYGridLines.Add
    Else
        Set aLine = mYGridLines.Item(i)
    End If
    aLine.SetPosition Me.newPoint(0, y, CoordsRelative, CoordsLogical), _
                        Me.newPoint(100, y, CoordsRelative, CoordsLogical)
    mGraphicObjects(aLine.Handle).boundingRect = mGraphicObjects(aLine.Handle).graphObj.BoundingRectangle
    addEntryToVisibleList aLine.Handle
    y = y + mScaleGridSpacingY
    drawObject aLine, mCanvas.boundary
Loop
For i = i + 1 To mYGridLines.Count
    Set aLine = mYGridLines.Item(i)
    aLine.SetPosition Me.newPoint(-1000, 0, CoordsLogical, CoordsLogical), _
                        Me.newPoint(-1000, 0, CoordsLogical, CoordsLogical)
    removeEntryFromVisibleList aLine.Handle
Next
End Sub

Private Sub redrawArea(areaToRedraw As TRectangle, ByVal startLayer As Long)
Dim Layer As Long
Dim index As Long

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

' redraw objects in the undrawn area, preserving layering
If Not mSuppressDrawing And areaToRedraw.isValid Then
    
    failpoint = 200

    mCanvas.setClippingRegion areaToRedraw
    
    failpoint = 300
    
    For Layer = startLayer To MaxLayer
        
        failpoint = 400
        
        index = mGraphicObjects(mVisibleTable(Layer).firstIndex).nextVisibleIndex
        Do While index <> mVisibleTable(Layer).lastindex
            With mGraphicObjects(index)
                
                failpoint = 500
                
                If rectOverlaps(.boundingRect, areaToRedraw) Then
                    
                    failpoint = 600
                    
                    Dim rects(0) As TRectangle
                    rects(0) = mCanvas.boundary
                    .graphObj.Draw rects
                End If
                index = .nextVisibleIndex
            End With
        Loop
    Next
    
    
    failpoint = 700
    
    redrawCursor
    
    failpoint = 800
    
    mCanvas.clearClippingRegion
    
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "redrawArea" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Private Sub redrawCursor()
Dim lPointerStyle As PointerStyles

lPointerStyle = IIf(mPointerMode = PointerModeDefault, mPointerStyle, mPointerToolStyle)

Select Case lPointerStyle
Case PointerNone

Case PointerCrosshairs
    drawCrosshairsPointer mPrevCursorX, mPrevCursorY
Case PointerDisc
    drawDiscPointer mPrevCursorX, mPrevCursorY
Case PointerTool
    drawToolPointer mPrevCursorX, mPrevCursorY
Case PointerCustom
    drawCustomPointer mPrevCursorX, mPrevCursorY
End Select
End Sub

Private Sub redrawObject( _
                ByVal graphObj As IGraphicObject, _
                ByRef areaToRedraw As TRectangle)
unDrawCursor areaToRedraw
drawObject graphObj, areaToRedraw
redrawArea areaToRedraw, graphObj.Layer + 1
End Sub

Private Sub removeEntryFromExtendedList(ByVal index As Long)
With mGraphicObjects(index)
    If .nextExtendedIndex = 0 Then Exit Sub
    mGraphicObjects(.nextExtendedIndex).prevExtendedIndex = .prevExtendedIndex
    mGraphicObjects(.prevExtendedIndex).nextExtendedIndex = .nextExtendedIndex
    .nextExtendedIndex = 0
    .prevExtendedIndex = 0
End With
End Sub

Private Sub removeEntryFromInScopeList(ByVal index As Long)
With mGraphicObjects(index)
    If .inScopeCycle = mInScopeCycle Then
        .inScopeCycle = 0
        mGraphicObjects(.nextInScopeIndex).prevInScopeIndex = .prevInScopeIndex
        mGraphicObjects(.prevInScopeIndex).nextInScopeIndex = .nextInScopeIndex
        .nextInScopeIndex = 0
        .prevInScopeIndex = 0
    End If
End With
End Sub

Private Sub removeEntryFromPeriodList(ByVal index As Long)
With mGraphicObjects(index)
    If .nextPeriodIndex = 0 Then Exit Sub
    mGraphicObjects(.nextPeriodIndex).prevPeriodIndex = .prevPeriodIndex
    mGraphicObjects(.prevPeriodIndex).nextPeriodIndex = .nextPeriodIndex
    .nextPeriodIndex = 0
    .prevPeriodIndex = 0
End With
End Sub

Private Sub removeEntryFromVisibleList(ByVal index As Long)
With mGraphicObjects(index)
    If .visibleCycle = mVisibleCycle Then
        .visibleCycle = 0
        mGraphicObjects(.nextVisibleIndex).prevVisibleIndex = .prevVisibleIndex
        mGraphicObjects(.prevVisibleIndex).nextVisibleIndex = .nextVisibleIndex
        .nextVisibleIndex = 0
        .prevVisibleIndex = 0
    End If
End With
End Sub

Private Function selectObjectOrSeries(ByVal graphObj As IGraphicObject) As Boolean
If graphObj.IsSelectable Then
    graphObj.Selected = True
    selectObjectOrSeries = True
ElseIf mGraphObjSeriesTable(graphObj.seriesId).IsSelectable Then
    mGraphObjSeriesTable(graphObj.seriesId).Selected = True
    selectObjectOrSeries = True
Else
    selectObjectOrSeries = False
End If
End Function

Private Sub setBarModeButtons()
If mToolbar Is Nothing Then Exit Sub

Select Case mFirstBarSeries.style.displayMode
Case BarDisplayModeBar
    mToolbar.Buttons("showbars").value = tbrPressed
    mToolbar.Buttons("showline").value = tbrUnpressed
    mToolbar.Buttons("showcandlesticks").value = tbrUnpressed
Case BarDisplayModeCandlestick
    mToolbar.Buttons("showbars").value = tbrUnpressed
    mToolbar.Buttons("showline").value = tbrUnpressed
    mToolbar.Buttons("showcandlesticks").value = tbrPressed
Case BarDisplayModeLine
    mToolbar.Buttons("showbars").value = tbrUnpressed
    mToolbar.Buttons("showline").value = tbrPressed
    mToolbar.Buttons("showcandlesticks").value = tbrUnpressed
End Select
End Sub

Private Sub setHorizontalGridTexts()
Dim y As Double
Dim i As Long
Dim aText As text

If mYGridTexts Is Nothing Then Exit Sub

y = Int(mCanvas.Bottom / mScaleGridSpacingY) * mScaleGridSpacingY
Do While y < mCanvas.Top
    i = i + 1
    If i > mYGridTexts.Count Then
        Set aText = mYGridTexts.Add
    Else
        Set aText = mYGridTexts.Item(i)
    End If
    If YScaleQuantum = 0.03125 Then
        ' tick size is 1/32 - label accordingly
        aText.text = Int(y) & " ' " & Format((y - Int(y)) / 0.03125, "00")
    Else
        aText.text = Format(y, mYScaleFormatStr)
    End If
    aText.position = mYAxisRegion.newPoint(15, y, CoordsLogical, CoordsLogical)
    y = y + mScaleGridSpacingY
Loop
For i = i + 1 To mYGridTexts.Count
    Set aText = mYGridTexts.Item(i)
    aText.text = ""
Next
End Sub

Private Sub setSelectedObjectOrSeries()
Dim startLayer As Long
Dim currLayer As Long
Dim i As Long
Dim index As Long
Dim currObjStillSelectable As Boolean
Dim newSelectedObject As IGraphicObject

If mSelectedObject Is Nothing Then
    startLayer = MaxLayer
Else
    startLayer = mSelectedObject.Layer
End If

For i = MinLayer To MaxLayer
    currLayer = (startLayer - i + MaxLayer + 1) Mod (MaxLayer + 1)
    index = mGraphicObjects(mVisibleTable(currLayer).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(currLayer).lastindex
        With mGraphicObjects(index)
            If isObjectOrSeriesSelectable(.graphObj) Then
                If .graphObj.HitTest(mCurrX, mCurrY) Then
                    If .graphObj Is mSelectedObject Then
                        currObjStillSelectable = True
                    Else
                        Set newSelectedObject = .graphObj
                        Exit For    ' we've found an object/series to select
                        
                    End If
                End If
            End If
            index = .nextVisibleIndex
        End With
    Loop
Next

If newSelectedObject Is Nothing Then
    If Not currObjStillSelectable Then
        deSelectObjectOrSeries mSelectedObject
        Set mSelectedObject = Nothing
    End If
Else
    deSelectObjectOrSeries mSelectedObject
    Set mSelectedObject = newSelectedObject
    selectObjectOrSeries newSelectedObject
End If
End Sub

Private Sub setupTitle()
Dim failpoint As Long
On Error GoTo Err

Set mTitle = AddText(LayerTitle)
mTitle.box = True
mTitle.boxColor = vbBlack
mTitle.boxStyle = LineStyles.LineInvisible
mTitle.boxThickness = 1
mTitle.boxFillColor = vbWhite
mTitle.boxFillStyle = FillStyles.FillSolid
mTitle.position = newPoint(0.2, 0.2, CoordsDistance, CoordsCounterDistance)
mTitle.fixedX = True
mTitle.fixedY = True
mTitle.align = TextAlignModes.AlignTopLeft
mTitle.includeInAutoscale = False
mTitle.paddingX = 0.5

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "setupTitle" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource


End Sub

Private Sub setVerticalGridLine(ByVal timestamp As Date, _
                            ByVal PeriodNumber As Long)
Dim vertGridInterval As Long
Dim gridLineTime As Date
Dim mins As Long
Dim secs As Long

If CDbl(timestamp) = 0 Then Exit Sub

If mVerticalGridTimePeriod Is Nothing Then
    ' just draw a vertical gridline every 10 bars
    If PeriodNumber Mod 10 = 1 Then
        addVerticalGridLine PeriodNumber, _
                    Format(timestamp, "yyyymmddhhnnss"), _
                    FormatDateTime(timestamp, vbShortTime), _
                    Format(timestamp, "yyyymmdd"), _
                    Format(timestamp, "d Mmm yy")
    End If
    Exit Sub
End If

gridLineTime = BarStartTime(timestamp, mVerticalGridTimePeriod)
    
Select Case mVerticalGridTimePeriod.units
Case TimePeriodSecond
    addVerticalGridLine PeriodNumber, _
                    Format(gridLineTime, "yyyymmddhhnnss"), _
                    FormatDateTime(gridLineTime, vbLongTime), _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d Mmm yy")
Case TimePeriodMinute
    addVerticalGridLine PeriodNumber, _
                    Format(gridLineTime, "yyyymmddhhnn"), _
                    FormatDateTime(gridLineTime, vbShortTime), _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d Mmm yy")
Case TimePeriodHour
    addVerticalGridLine PeriodNumber, _
                    Format(gridLineTime, "yyyymmddhh"), _
                    FormatDateTime(gridLineTime, vbShortTime), _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d Mmm yy")
Case TimePeriodDay
    addVerticalGridLine PeriodNumber, _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d"), _
                    Format(gridLineTime, "yyyymm"), _
                    Format(gridLineTime, "Mmm yy")
Case TimePeriodWeek
    addVerticalGridLine PeriodNumber, _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d Mmm"), _
                    Format(gridLineTime, "yyyy"), _
                    Format(gridLineTime, "yyyy")
Case TimePeriodMonth
    addVerticalGridLine PeriodNumber, _
                    Format(gridLineTime, "yyyymm"), _
                    Format(gridLineTime, "Mmm"), _
                    Format(gridLineTime, "yyyy"), _
                    Format(gridLineTime, "yyyy")
Case TimePeriodYear
    addVerticalGridLine PeriodNumber, _
                    Format(gridLineTime, "yyyy"), _
                    Format(gridLineTime, "YYYY"), _
                    "$", _
                    ""
Case TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    If PeriodNumber Mod 10 = 1 Then
        addVerticalGridLine PeriodNumber, _
                        Format(gridLineTime, "yyyymmddhhnnss"), _
                        FormatDateTime(gridLineTime, vbLongTime), _
                        Format(gridLineTime, "yyyymmdd"), _
                        Format(gridLineTime, "d Mmm yy")
    End If
End Select
End Sub

Private Sub undrawArea(undrawnArea As TRectangle)
If Not mSuppressDrawing And undrawnArea.isValid Then
    If mUseDeferredPainting Then
        Dim i As Long
        Dim handled As Boolean
        For i = 0 To mUndrawnAreasIndex
            If rectOverlaps(undrawnArea, mUndrawnAreas(i)) Then
                 mUndrawnAreas(i) = rectUnion(undrawnArea, mUndrawnAreas(i))
                 handled = True
            End If
        Next
        If Not handled Then
            mUndrawnAreasIndex = mUndrawnAreasIndex + 1
            If mUndrawnAreasIndex > UBound(mUndrawnAreas) Then
                ReDim Preserve mUndrawnAreas(2 * (UBound(mUndrawnAreas) + 1) - 1) As TRectangle
            End If
            mUndrawnAreas(mUndrawnAreasIndex) = undrawnArea
        End If
        deferredPaint DeferredPaintRepaintChangedAreas
    Else
        undrawAreaEx undrawnArea
    End If
End If

End Sub

Private Sub undrawAreaEx(undrawnArea As TRectangle)

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

' redraw objects in the undrawn area, preserving layering
If Not mSuppressDrawing And undrawnArea.isValid Then
    
    failpoint = 200

    mCanvas.setClippingRegion undrawnArea
    mCanvas.paintBackground
    mCanvas.clearClippingRegion
    
    failpoint = 400
    
    redrawArea undrawnArea, MinLayer
    
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "undrawAreaEx" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Private Sub unDrawCursor( _
                ByRef areaToRedraw As TRectangle)
Dim lPointerStyle As PointerStyles

lPointerStyle = IIf(mPointerMode = PointerModeDefault, mPointerStyle, mPointerToolStyle)

mCanvas.setClippingRegion areaToRedraw

Select Case lPointerStyle
Case PointerNone

Case PointerCrosshairs
    undrawCrosshairsPointer
Case PointerDisc
    undrawDiscPointer
Case PointerTool
    undrawToolPointer
Case PointerCustom
    undrawCustomPointer
End Select
mCanvas.clearClippingRegion
End Sub

Private Sub undrawCrosshairsPointer()
Dim prevLineX As Single

If mSuppressDrawing Then Exit Sub

With mCanvas
    .setPenAttributes mPointerCrosshairsColor Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .drawLine .newPoint(.Left, mPrevCursorY), .newPoint(YAxisPosition, mPrevCursorY)
    prevLineX = Round(mPrevCursorX)
    .drawLine .newPoint(prevLineX, .Bottom), .newPoint(prevLineX, .Top)
End With

End Sub

Private Sub undrawCustomPointer()

' nothing to do

End Sub

Private Sub undrawDiscPointer()
Dim size As Dimension
Dim prevLineX As Single

If mSuppressDrawing Then Exit Sub

With mCanvas
    prevLineX = Round(mPrevCursorX)
    Set size = .newDimension(0.4, 0.4)
    
    .setPenAttributes vbBlack Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .drawLine .newPoint(prevLineX - size.XLogical, mPrevCursorY), .newPoint(prevLineX + size.XLogical, mPrevCursorY)
    .drawLine .newPoint(prevLineX, mPrevCursorY - size.YLogical), .newPoint(prevLineX, mPrevCursorY + size.YLogical)
    
    .setPenAttributes mPointerDiscColor Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .setBrushAttributes mPointerDiscColor Xor .BackColor, FillStyles.FillSolid
    .drawCircle .newPoint(prevLineX, mPrevCursorY), .newDimension(0.8, 0.8).XLogical
End With

End Sub

Private Sub undrawToolPointer()
Dim size As Dimension
Dim prevLineX As Single

If mSuppressDrawing Then Exit Sub

With mCanvas
    prevLineX = Round(mPrevCursorX)
    Set size = .newDimension(0.4, 0.4)
    
    .setPenAttributes vbBlack Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .drawLine .newPoint(prevLineX - size.XLogical, mPrevCursorY), .newPoint(prevLineX + size.XLogical, mPrevCursorY)
    .drawLine .newPoint(prevLineX, mPrevCursorY - size.YLogical), .newPoint(prevLineX, mPrevCursorY + size.YLogical)
    
End With

End Sub


