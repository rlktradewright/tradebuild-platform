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

Event AutoscalingChanged()

Event Click()
Attribute Click.VB_UserMemId = -600

Event DblCLick()
Attribute DblCLick.VB_UserMemId = -601

Event MouseDown(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Attribute MouseDown.VB_UserMemId = -605
                
Event mouseMove(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Attribute mouseMove.VB_UserMemId = -606
                
Event MouseUp(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
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

Private Enum PropertyOverrideFlags
    AutoscalingIsSet = 1
    GridTextStyleIsSet = 2
    HasGridIsSet = 4
    HasGridTextIsSet = 8
    IntegerYScaleIsSet = &H10&
    YScaleQuantumIsSet = &H20&
    GridlineSpacingYIsSet = &H40&
    MinimumHeightIsSet = &H80&
    CursorSnapsToTickBoundariesIsSet = &H100&
    BackGradientFillColorsIsSet = &H200&
    GridLineStyleIsSet = &H400&
    YCursorTextStyleIsSet = &H800&
    YAxisTextStyleIsSet = &H1000&
End Enum

'================================================================================
' Constants
'================================================================================


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
    SeriesID            As Long
    boundingRect        As TRectangle
    PeriodNumber        As Long
    ExtendedObject      As Boolean
    IncludeInAutoscale  As Boolean
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

Private mChart As Chart

Private mCleared As Boolean

Private mName As String
Private mRegionNumber As Long

Private mIsXAxisRegion As Boolean
Private mIsYAxisRegion As Boolean

Private WithEvents mStyle As ChartRegionStyle
Attribute mStyle.VB_VarHelpID = -1

Private mTitle As Text

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

Private mCanvas As Canvas
Private mIsDrawingSuppressed As Boolean

Private WithEvents mToolbar As Toolbar
Attribute mToolbar.VB_VarHelpID = -1

Private mMinInViewPeriod As Long
Private mMaxInViewPeriod As Long

Private mNextLayer As Long

Private mDrawn As Boolean

Private mGraphObjServiceProvider As GraphObjServiceProvider
Attribute mGraphObjServiceProvider.VB_VarHelpID = -1

Private mCrosshairLineHoriz As Line
Private mCrosshairLineVert As Line

Private mYAxisRegion As ChartRegion
Private mYCursorText As Text      ' used to indicate the position of the pointer
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

Private mPercentHeight As Double
Private mMinimumPercentHeight As Double
Private mYAxisPosition As Long

Private mScaleGridSpacingY As Double
Private mYScaleFormatStr As String
Private mYScaleSubFormatStr As String
Private mGridTextHeight As Double

Private mPerformanceTextVisible As Boolean

Private mPointerCrosshairsColor As Long
Private mPointerDiscColor As Long

Private mPrevCursorX As Single
Private mPrevCursorY As Single

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

' these fields override corresponding items in the Style
Private mAutoscaling As Boolean
Private mGridTextColor  As Long
Private mHasGrid As Boolean
Private mHasGridText As Boolean
Private mIntegerYScale As Boolean
Private mYScaleQuantum As Double
Private mGridlineSpacingY As Double
Private mMinimumHeight As Double
Private mCursorSnapsToTickBoundaries As Boolean
Private mBackGradientFillColors() As Long
Private WithEvents mGridLineStyle As LineStyle
Attribute mGridLineStyle.VB_VarHelpID = -1
Private WithEvents mGridTextStyle As TextStyle
Attribute mGridTextStyle.VB_VarHelpID = -1
Private WithEvents mYAxisTextStyle As TextStyle
Attribute mYAxisTextStyle.VB_VarHelpID = -1
Private WithEvents mYCursorTextStyle As TextStyle
Attribute mYCursorTextStyle.VB_VarHelpID = -1

Private mPropertyOverrideFlags As PropertyOverrideFlags

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

'If gIsInDev Or gLogger.isLoggable(LogLevelHighDetail) Then
If gIsInDev Then
    If et Is Nothing Then Set et = New ElapsedTimer
    et.StartTiming
End If

Select Case mDeferredPaintAction
Case DeferredPaintRepaintNone

Case DeferredPaintRepaintChangedAreas
    

    failpoint = 100

    For i = 0 To mUndrawnAreasIndex
        failpoint = failpoint + i
        undrawAreaEx mUndrawnAreas(i)
    Next
    
    
    failpoint = 200
    
    For i = 0 To mRedrawObjectsIndex
        failpoint = failpoint + i
        redrawObject mRedrawObjects(i).graphObj, mRedrawObjects(i).areaToRedraw
    Next
Case DeferredPaintRepaintRegion
    
    failpoint = 300
    
    paint
Case DeferredPaintChangeGaugeandRepaint
    
    failpoint = 400
    
    calculateInscopeGaugeDependentObjects
    
    failpoint = 500
    
    
    failpoint = 600
    
    paint
    
    
    failpoint = 700
    
    If Not mYAxisRegion Is Nothing Then mYAxisRegion.resizedY
Case DeferredPaintScaleAndRepaint
    
    failpoint = 800
    
    calculateVerticalBounds
    
    failpoint = 900
    
    determineVisibleObjects
    
    failpoint = 1000
    
    paint
    
    failpoint = 1100
    
    If Not mYAxisRegion Is Nothing Then
        mYAxisRegion.SetVerticalScaleInternal mCanvas.Bottom, mCanvas.Top
    End If
Case DeferredPaintCalculateAndRepaint
    
    failpoint = 1200
    
    determineInScopeObjects
    
    failpoint = 1300
    
    If Autoscaling Then
        calculateVerticalBounds
    End If
    
    failpoint = 1400
    
    determineVisibleObjects
    
    failpoint = 1500
    
    paint
    
    failpoint = 1600
    
    If Not mYAxisRegion Is Nothing Then
        mYAxisRegion.SetVerticalScaleInternal mCanvas.Bottom, mCanvas.Top
    End If
End Select


'If gIsInDev Or gLogger.isLoggable(LogLevelHighDetail) Then
If gIsInDev Then
    Dim s As String
    s = "Deferred paint: " & Format(et.ElapsedTimeMicroseconds, "0")
    Debug.Print s
    'gLogger.Log LogLevelHighDetail, s
End If

deferredPaint DeferredPaintRepaintNone

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "DeferredAction_run" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

deferredPaint DeferredPaintRepaintNone

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
Task_TaskName = mTaskContext.Name
End Property

'================================================================================
' GraphObjServiceProvider Methods
'================================================================================

Friend Sub GOSP_AddGraphicObject( _
                ByVal graphObj As IGraphicObject, _
                ByVal SeriesID As Long)
Dim failpoint As Long
On Error GoTo Err

AddGraphicObject graphObj, SeriesID

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GOSP_addGraphicObject" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub GOSP_ObjectChanged( _
                ByVal Handle As Long, _
                ByRef areaToRedraw As TRectangle)
Dim redrawArea As TRectangle

Dim failpoint As Long
On Error GoTo Err

' the following is to prevent lines from going past the rightmost period of the chart
redrawArea = rectIntersection(areaToRedraw, mCanvas.Boundary)

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
    .IncludeInAutoscale = .graphObj.IncludeInAutoscale
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
    
If mIsDrawingSuppressed Then Exit Sub

If Not mGraphicObjects(Handle).boundingRect.isValid Then Exit Sub
If mGraphicObjects(Handle).NoDraw Then Exit Sub

With mGraphicObjects(Handle)
    ' check if the vertical bounds have changed, and redraw if so
    If Autoscaling And .IncludeInAutoscale Then
        If .boundingRect.Top >= mCanvas.Boundary.Top Or .boundingRect.Bottom <= (mCanvas.Boundary.Bottom) Then
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
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GOSP_objectChanged" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription



End Sub

Friend Sub GOSP_ObjectUndrawn(ByVal Handle As Long, undrawnArea As TRectangle)

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

undrawArea undrawnArea, mGraphicObjects(Handle).graphObj

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GOSP_objectUndrawn" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub GOSP_PaintRegion( _
                ByVal recalculate As Boolean)

Dim failpoint As Long
On Error GoTo Err

If mIsDrawingSuppressed Then Exit Sub

If mUseDeferredPainting Then
    If recalculate Then
        deferredPaint DeferredPaintCalculateAndRepaint
    Else
        deferredPaint DeferredPaintRepaintRegion
    End If
Else
    If recalculate Then
        determineInScopeObjects
        If Autoscaling Then
            calculateVerticalBounds
        End If
        determineVisibleObjects
    End If
    paint
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GOSP_paintRegion" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub GOSP_RecalcBoundary( _
                ByVal SeriesID As Long)
Dim i As Long
Dim index As Long

Dim failpoint As Long
On Error GoTo Err

For i = MinLayer To MaxLayer
    index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(i).lastindex
        With mGraphicObjects(index)
            If .SeriesID = SeriesID Then
                .boundingRect = .graphObj.BoundingRectangle
            End If
            index = .nextVisibleIndex
        End With
    Loop
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GOSP_recalcBoundary" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub GOSP_RemoveObject( _
                ByVal Handle As Long)
RemoveGraphicObjectEx mGraphicObjects(Handle).graphObj, True, False
End Sub

'================================================================================
' mFirstBarSeries Event Handlers
'================================================================================

Private Sub mFirstBarSeries_PropertyChanged(ev As PropertyChangedEvent)
If UCase$(ev.PropertyName) = UCase$("DisplayMode") Then setBarModeButtons
End Sub

'================================================================================
' mGridLineStyle Event Handlers
'================================================================================

Private Sub mGridLineStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
setGridLineStyle
End Sub

'================================================================================
' mGridTextStyle Event Handlers
'================================================================================

Private Sub mGridTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
setGridTextStyle
End Sub

'================================================================================
' mStyle Event Handlers
'================================================================================

Private Sub mStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
Dim failpoint As Long
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case UCase$("Autoscaling")
    setAutoscaling
Case UCase$("BackColor")
    setBackColor
Case UCase$("BackGradientFillColors")
    setBackGradientFillColors
Case UCase$("GridlineSpacingY")
    setGridlineSpacingY
Case UCase$("GridLineStyle")
    setGridLineStyle
Case UCase$("GridTextStyle")
    setGridTextStyle
Case UCase$("HasGrid")
    setHasGrid
Case UCase$("HasGridText")
    setHasGridText
Case UCase$("IntegerYScale")
    setIntegerYScale
Case UCase$("MinimumHeight")
    setMinimumHeight
Case UCase$("CursorSnapsToTickBoundaries")
    setCursorSnapsToTickBoundaries
Case UCase$("YAxisTextStyle")
    setYAxisTextStyle
Case UCase$("YCursorTextStyle")
    setYCursorTextStyle
Case UCase$("YScaleQuantum")
    setYScaleQuantum
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & "mStyle_PropertyChanged", _
            "Unhandled property change"
End Select

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "mStyle_PropertyChanged" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' mToolbar Event Handlers
'================================================================================

Private Sub mToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim failpoint As Long
On Error GoTo Err

Select Case Button.Key
Case ToolbarCommandShowBars
    If mFirstBarSeries.BarThickness > 1 Then
        mToolbar.Buttons(ToolbarCommandThinnerBars).Enabled = True
    Else
        mToolbar.Buttons(ToolbarCommandThinnerBars).Enabled = False
    End If
    mFirstBarSeries.DisplayMode = BarDisplayModeBar
Case ToolbarCommandShowCandlesticks
    mFirstBarSeries.DisplayMode = BarDisplayModeCandlestick
    If mFirstBarSeries.BarWidth <= 0.1 Then
        mToolbar.Buttons(ToolbarCommandThinnerBars).Enabled = False
    Else
        mToolbar.Buttons(ToolbarCommandThinnerBars).Enabled = True
    End If
Case ToolbarCommandShowLine
    ' not yet implemented in ChartSkil
Case ToolbarCommandThinnerBars
    If mFirstBarSeries.DisplayMode = BarDisplayModeCandlestick Then
        If mFirstBarSeries.BarWidth > 0.1 Then
            mFirstBarSeries.BarWidth = mFirstBarSeries.BarWidth - 0.1
        End If
        If mFirstBarSeries.BarWidth <= 0.1 Then
            Button.Enabled = False
        End If
    ElseIf mFirstBarSeries.DisplayMode = BarDisplayModeBar Then
        If mFirstBarSeries.BarThickness > 1 Then
            mFirstBarSeries.BarThickness = mFirstBarSeries.BarThickness - 1
        End If
        If mFirstBarSeries.BarThickness = 1 Then
            Button.Enabled = False
        End If
    End If
Case ToolbarCommandThickerBars
    If mFirstBarSeries.DisplayMode = BarDisplayModeCandlestick Then
        mFirstBarSeries.BarWidth = mFirstBarSeries.BarWidth + 0.1
    ElseIf mFirstBarSeries.DisplayMode = BarDisplayModeBar Then
        mFirstBarSeries.BarThickness = mFirstBarSeries.BarThickness + 1
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
    Autoscaling = True
End Select

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "mToolbar_ButtonClick" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

'================================================================================
' mYAxisTextStyle Event Handlers
'================================================================================

Private Sub mYAxisTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
setYAxisTextStyle
End Sub

'================================================================================
' mYCursorTextStyle Event Handlers
'================================================================================

Private Sub mYCursorTextStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
setYCursorTextStyle
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get Autoscaling() As Boolean
Autoscaling = mStyle.Autoscaling
If isPropertyOverrideFlagSet(AutoscalingIsSet) Then Autoscaling = mAutoscaling
End Property

Public Property Let Autoscaling(ByVal value As Boolean)
If mAutoscaling = value And isPropertyOverrideFlagSet(AutoscalingIsSet) Then Exit Property
mAutoscaling = value
setPropertyOverrideFlag AutoscalingIsSet
setAutoscaling
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = mStyle.BackColor
If isPropertyOverrideFlagSet(BackGradientFillColorsIsSet) Then BackColor = mBackGradientFillColors(0)
End Property

Public Property Let BackColor(ByVal value As OLE_COLOR)
ReDim mBackGradientFillColors(0) As Long
mBackGradientFillColors(0) = value
setPropertyOverrideFlag AutoscalingIsSet
setBackColor
End Property

Public Property Get BackGradientFillColors() As Long()
BackGradientFillColors = mStyle.BackGradientFillColors
If isPropertyOverrideFlagSet(BackGradientFillColorsIsSet) Then BackGradientFillColors = mBackGradientFillColors
End Property

Public Property Let BackGradientFillColors(ByRef value() As Long)
Dim ar() As Long
ar = value
mBackGradientFillColors = ar
setPropertyOverrideFlag BackGradientFillColorsIsSet
setBackGradientFillColors
End Property

Public Property Get Bottom() As Double
Bottom = mCanvas.Boundary.Bottom
End Property

Friend Property Let Bottom(ByVal value As Double)
mCanvas.Bottom = value
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.Bottom = value
End If
End Property

Friend Property Get Canvas() As Canvas
Set Canvas = mCanvas
End Property

Public Property Let CursorSnapsToTickBoundaries(ByVal value As Boolean)
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "snapCursorToTickBoundaries", _
            "Cannot set snapCursorToTickBoundaries for an axis region"
End If
If mCursorSnapsToTickBoundaries = value And isPropertyOverrideFlagSet(CursorSnapsToTickBoundariesIsSet) Then Exit Property
mCursorSnapsToTickBoundaries = value
setPropertyOverrideFlag CursorSnapsToTickBoundariesIsSet
setCursorSnapsToTickBoundaries
End Property

Public Property Get CursorSnapsToTickBoundaries() As Boolean
CursorSnapsToTickBoundaries = mStyle.CursorSnapsToTickBoundaries
If isPropertyOverrideFlagSet(CursorSnapsToTickBoundariesIsSet) Then CursorSnapsToTickBoundaries = mCursorSnapsToTickBoundaries
End Property

Public Property Get FirstBarSeries() As BarSeries
Set FirstBarSeries = mFirstBarSeries
End Property

Public Property Get GridlineSpacingY() As Double
GridlineSpacingY = mStyle.GridlineSpacingY
If isPropertyOverrideFlagSet(GridlineSpacingYIsSet) Then GridlineSpacingY = mGridlineSpacingY
End Property

Public Property Let GridlineSpacingY(ByVal value As Double)
If mGridlineSpacingY = value And isPropertyOverrideFlagSet(GridlineSpacingYIsSet) Then Exit Property
mGridlineSpacingY = value
setPropertyOverrideFlag GridlineSpacingYIsSet
setGridlineSpacingY
End Property

Public Property Get GridLineStyle() As LineStyle
Set GridLineStyle = mStyle.GridLineStyle
If isPropertyOverrideFlagSet(GridLineStyleIsSet) Then Set GridLineStyle = mGridLineStyle
End Property

Public Property Let GridLineStyle(ByVal value As LineStyle)
If mGridLineStyle Is value And isPropertyOverrideFlagSet(GridLineStyleIsSet) Then Exit Property
Set mGridLineStyle = value
If mGridLineStyle Is Nothing Then
    clearPropertyOverrideFlag GridLineStyleIsSet
Else
    setPropertyOverrideFlag GridLineStyleIsSet
End If
setGridLineStyle
End Property

Public Property Get GridTextStyle() As TextStyle
Set GridTextStyle = mStyle.GridTextStyle
If isPropertyOverrideFlagSet(GridTextStyleIsSet) Then Set GridTextStyle = mGridTextStyle
End Property

Public Property Let GridTextStyle(ByVal value As TextStyle)
If mGridTextStyle Is value And isPropertyOverrideFlagSet(GridTextStyleIsSet) Then Exit Property
Set mGridTextStyle = value
If mGridTextStyle Is Nothing Then
    clearPropertyOverrideFlag GridTextStyleIsSet
Else
    setPropertyOverrideFlag GridTextStyleIsSet
End If
setGridTextStyle
End Property

Public Property Get HasGrid() As Boolean
HasGrid = mStyle.HasGrid
If isPropertyOverrideFlagSet(HasGridIsSet) Then HasGrid = mHasGrid
End Property

Public Property Let HasGrid(ByVal value As Boolean)
If mHasGrid = value And isPropertyOverrideFlagSet(HasGridIsSet) Then Exit Property
mHasGrid = value
setPropertyOverrideFlag HasGridIsSet
setHasGrid
End Property

Public Property Get HasGridText() As Boolean
HasGridText = mStyle.HasGridText
If isPropertyOverrideFlagSet(HasGridTextIsSet) Then HasGridText = mHasGridText
End Property

Public Property Let HasGridText(ByVal value As Boolean)
If mHasGridText = value And isPropertyOverrideFlagSet(HasGridTextIsSet) Then Exit Property
mHasGridText = value
setPropertyOverrideFlag HasGridTextIsSet
setHasGridText
End Property

Public Property Get Height() As Double
Height = mCanvas.Height
End Property

Public Property Let IntegerYScale(ByVal value As Boolean)
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "integerYScale", _
            "Cannot set integerYScale for an axis region"
End If

If mIntegerYScale = value And isPropertyOverrideFlagSet(IntegerYScaleIsSet) Then Exit Property
mIntegerYScale = value
setPropertyOverrideFlag IntegerYScaleIsSet
setIntegerYScale
End Property

Public Property Get IntegerYScale() As Boolean
IntegerYScale = mStyle.IntegerYScale
If isPropertyOverrideFlagSet(IntegerYScaleIsSet) Then IntegerYScale = mIntegerYScale
End Property

Friend Property Let IsDrawingEnabled(ByVal value As Boolean)
If mIsDrawingSuppressed = value Then Exit Property
mIsDrawingSuppressed = value
If mIsDrawingSuppressed Then
    mCanvas.MousePointer = vbHourglass
Else
    If mIsXAxisRegion Or mIsYAxisRegion Then
        mCanvas.MousePointer = vbDefault
    Else
        Select Case mPointerMode
        Case PointerModeDefault
            setMouseIcon mPointerStyle
        Case PointerModeTool
            setMouseIcon PointerStyles.PointerTool
        Case PointerModeSelection
            setMouseIcon mPointerToolStyle
        End Select
        showCursor mPrevCursorX, mPrevCursorY
    End If
    
    If mUseDeferredPainting Then
        deferredPaint DeferredPaintCalculateAndRepaint
    Else
        determineInScopeObjects
        If Autoscaling Then
            calculateVerticalBounds
        End If
        determineVisibleObjects
        paint
    End If
End If
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.IsDrawingEnabled = mIsDrawingSuppressed
End If
End Property

Friend Property Let IsXAxisRegion(ByVal value As Boolean)
mIsXAxisRegion = value
mCanvas.MousePointer = vbDefault
End Property

Public Property Get IsXAxisRegion() As Boolean
IsXAxisRegion = mIsXAxisRegion
End Property

Friend Property Let IsYAxisRegion(ByVal value As Boolean)
mIsYAxisRegion = value
mCanvas.MousePointer = vbDefault
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
If isPropertyOverrideFlagSet(MinimumHeightIsSet) Then MinimumHeight = mMinimumHeight
End Property

Public Property Let MinimumHeight(ByVal value As Double)
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "minimumHeight", _
            "Cannot set minimumHeight for an axis region"
End If

If mMinimumHeight = value And isPropertyOverrideFlagSet(MinimumHeightIsSet) Then Exit Property
mMinimumHeight = value
setPropertyOverrideFlag MinimumHeightIsSet
setMinimumHeight
End Property

Public Property Get MinimumPercentHeight() As Double
MinimumPercentHeight = mMinimumPercentHeight
End Property

Friend Property Let MinimumPercentHeight(ByVal value As Double)
mMinimumPercentHeight = value
End Property

Public Property Get MousePosition() As Point
Set MousePosition = NewPoint(mCurrX, mCurrY)
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Get PercentHeight() As Double
PercentHeight = mPercentHeight
End Property

Friend Property Let PercentHeight(ByVal value As Double)
mPercentHeight = value
End Property

Public Property Get PerformanceTextVisible() As Boolean
PerformanceTextVisible = mPerformanceTextVisible
End Property

Public Property Let PerformanceTextVisible(ByVal value As Boolean)
mPerformanceTextVisible = value
If mPerformanceTextVisible Then
    If mElapsedTimer Is Nothing Then Set mElapsedTimer = New ElapsedTimer
Else
    Set mElapsedTimer = Nothing
End If
End Property

Public Property Get PointerCrosshairsColor() As OLE_COLOR
PointerCrosshairsColor = mPointerCrosshairsColor
End Property

Friend Property Let PointerCrosshairsColor(ByVal value As OLE_COLOR)
Dim lUseDeferredPainting As Boolean

If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "PointerCrosshairsColor", _
            "Cannot set PointerCrossHairsColor for an axis region"
End If

mPointerCrosshairsColor = value

lUseDeferredPainting = mUseDeferredPainting
mUseDeferredPainting = False
mCrosshairLineHoriz.Color = mPointerCrosshairsColor
mCrosshairLineVert.Color = mPointerCrosshairsColor
mUseDeferredPainting = lUseDeferredPainting
End Property

Public Property Get PointerDiscColor() As OLE_COLOR
PointerDiscColor = mPointerDiscColor
End Property

Friend Property Let PointerDiscColor(ByVal value As OLE_COLOR)
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

If mPointerMode <> PointerModeDefault Then
    Set mPointerIcon = value
    Exit Property
End If

If Not mDrawn Then
    Set mPointerIcon = value
    Exit Property
End If

If Not mIsDrawingSuppressed Then hideCursor

Set mPointerIcon = value

If Not mIsDrawingSuppressed Then
    setMouseIcon mPointerStyle
    showCursor mPrevCursorX, mPrevCursorY
End If

End Property

Public Property Get PointerStyle() As PointerStyles
PointerStyle = mPointerStyle
End Property

Public Property Let PointerStyle(ByVal value As PointerStyles)
Dim failpoint As Long
On Error GoTo Err

If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "pointerStyle", _
            "Cannot set PointerStyle for an axis region"
End If


If mPointerStyle = value Then Exit Property

If mPointerMode <> PointerModeDefault Then
    mPointerStyle = value
    Exit Property
End If

If Not mDrawn Then
    mPointerStyle = value
    Exit Property
End If

If Not mIsDrawingSuppressed Then hideCursor

mPointerStyle = value

If Not mIsDrawingSuppressed Then
    setMouseIcon mPointerStyle
    showCursor mPrevCursorX, mPrevCursorY
End If

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "PointerStyle" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get RegionNumber() As Long
RegionNumber = mRegionNumber
End Property

Friend Property Let RegionNumber(ByVal value As Long)
mRegionNumber = value
End Property

Public Property Get ScaleGridSpacingY() As Single
ScaleGridSpacingY = mScaleGridSpacingY
End Property

Friend Property Let SessionStartTime(ByVal value As Date)
mSessionOffsetMinutes = calculateMinutesFromMidnight(value) Mod 60
End Property

Public Property Get Style() As ChartRegionStyle
Set Style = mStyle
End Property

Public Property Let Style( _
                ByVal value As ChartRegionStyle)

Dim failpoint As Long
On Error GoTo Err

If value Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "Style", _
            "Value must not be Nothing"
End If

Set mStyle = value

setBackColor
setBackGradientFillColors
setGridTextStyle
If Not mIsXAxisRegion And Not mIsYAxisRegion Then
    setAutoscaling
    setGridlineSpacingY
    setIntegerYScale
    setMinimumHeight
    setCursorSnapsToTickBoundaries
    setYAxisTextStyle
    setYCursorTextStyle
    setYScaleQuantum
End If
setHasGrid
setHasGridText

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Style" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get Title() As Text
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "Title", _
            "An axis region does not have a title"
End If

If mTitle Is Nothing Then setupTitle
Set Title = mTitle
End Property

Friend Property Let Toolbar(ByVal value As Toolbar)
Dim btn As Button

Set mToolbar = value

mToolbar.Buttons("autoscale").Enabled = Not Autoscaling
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

Public Property Get Width() As Single
Width = mCanvas.Width
End Property

Public Property Get YAxisPosition() As Long
YAxisPosition = mYAxisPosition
End Property

Friend Property Let YAxisRegion(ByVal value As ChartRegion)
Dim aFont As StdFont

Set mYAxisRegion = value

mYAxisRegion.HasGrid = False
mYAxisRegion.HasGridText = False

mYAxisRegion.IsDrawingEnabled = mIsDrawingSuppressed
mYAxisRegion.Bottom = 0
mYAxisRegion.Top = 1
mYAxisRegion.PeriodsInView 0, 99

Dim aTextSeries As New TextSeries
mYAxisRegion.AddGraphicObjectSeries aTextSeries, LayerNumbers.LayerPointer
Set mYCursorText = aTextSeries.Add
mYCursorText.Align = AlignBoxCentreLeft
setYCursorTextStyle

Set mYGridTexts = mYAxisRegion.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText)
mYGridTexts.Align = AlignCentreLeft
setYAxisTextStyle

End Property

Public Property Get YAxisRegion() As ChartRegion
Set YAxisRegion = mYAxisRegion
End Property

Public Property Get YAxisTextStyle() As TextStyle
Set YAxisTextStyle = mStyle.YAxisTextStyle
If isPropertyOverrideFlagSet(YAxisTextStyleIsSet) Then Set YAxisTextStyle = mYAxisTextStyle
End Property

Public Property Let YAxisTextStyle(ByVal value As TextStyle)
If mYAxisTextStyle Is value And isPropertyOverrideFlagSet(YAxisTextStyleIsSet) Then Exit Property
Set mYAxisTextStyle = value
If mYAxisTextStyle Is Nothing Then
    clearPropertyOverrideFlag YAxisTextStyleIsSet
Else
    setPropertyOverrideFlag YAxisTextStyleIsSet
End If
setYAxisTextStyle
End Property

Public Property Get YCursorTextStyle() As TextStyle
Set YCursorTextStyle = mStyle.YCursorTextStyle
If isPropertyOverrideFlagSet(YCursorTextStyleIsSet) Then Set YCursorTextStyle = mYCursorTextStyle
End Property

Public Property Let YCursorTextStyle(ByVal value As TextStyle)
If mYCursorTextStyle Is value And isPropertyOverrideFlagSet(YCursorTextStyleIsSet) Then Exit Property
Set mYCursorTextStyle = value
If mYCursorTextStyle Is Nothing Then
    clearPropertyOverrideFlag YCursorTextStyleIsSet
Else
    setPropertyOverrideFlag YCursorTextStyleIsSet
End If
setYCursorTextStyle
End Property

Public Property Let YScaleQuantum(ByVal value As Double)
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "YScaleQuantum", _
            "Cannot set YScaleQuantum for an axis region"
End If

If mYScaleQuantum = value And isPropertyOverrideFlagSet(YScaleQuantumIsSet) Then Exit Property
mYScaleQuantum = value
setPropertyOverrideFlag YScaleQuantumIsSet
setYScaleQuantum
End Property

Public Property Get YScaleQuantum() As Double
YScaleQuantum = mStyle.YScaleQuantum
If isPropertyOverrideFlagSet(YScaleQuantumIsSet) Then YScaleQuantum = mYScaleQuantum
End Property

'================================================================================
' Methods
'================================================================================

Public Function AddGraphicObjectSeries( _
                ByVal graphObjSeries As IGraphicObjectSeries, _
                Optional ByVal Layer As Long = -1, _
                Optional ByVal Name As String) As IGraphicObjectSeries
Dim failpoint As Long
On Error GoTo Err

Set AddGraphicObjectSeries = graphObjSeries

If mGraphObjSeriesIndex > UBound(mGraphObjSeriesTable) Then
    ReDim Preserve mGraphObjSeriesTable(2 * (UBound(mGraphObjSeriesTable) + 1) - 1) As IGraphicObjectSeries
End If

Set mGraphObjSeriesTable(mGraphObjSeriesIndex) = graphObjSeries

If Name = "" Then Name = "Series" & mGraphObjSeriesIndex

graphObjSeries.Initialise Name, _
                        mCanvas, _
                        getLayer(Layer), _
                        mChart, _
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
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "addGraphicObjectSeries" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Public Function AddLine( _
                Optional ByVal pLayer As LayerNumbers = -1, _
                Optional ByVal Style As LineStyle) As Line
Static miscLines As LineSeries
If miscLines Is Nothing Then
    Set miscLines = AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerHighestUser - 1)
End If
Set AddLine = miscLines.Add
If pLayer <> -1 Then AddLine.Layer = pLayer
If Not Style Is Nothing Then AddLine.LocalStyle = Style
End Function

Friend Sub addPeriod(ByVal PeriodNumber As Long, _
                    ByVal Timestamp As Date)
Dim failpoint As Long
On Error GoTo Err

addNewPeriod PeriodNumber
If CDbl(Timestamp) <> 0 Then setVerticalGridLine Timestamp, PeriodNumber

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "addPeriod" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Function AddText( _
                Optional ByVal value As String, _
                Optional ByVal pLayer As LayerNumbers = -1, _
                Optional ByVal Style As TextStyle) As Text
Static miscTexts As TextSeries
If miscTexts Is Nothing Then
    Set miscTexts = AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerHighestUser)
End If
Set AddText = miscTexts.Add
If value <> "" Then AddText.Text = value
If pLayer <> -1 Then AddText.Layer = pLayer
If Not Style Is Nothing Then AddText.LocalStyle = Style
End Function

Friend Sub ClearRegion()
Dim i As Long
Dim t As GraphObjSeriesRemovalTask

Dim failpoint As Long
On Error GoTo Err

Set mTitle = Nothing

Set mFirstBarSeries = Nothing

Set mCrosshairLineHoriz = Nothing
Set mCrosshairLineVert = Nothing

Set mYAxisRegion = Nothing
Set mYCursorText = Nothing

Set mYGridTexts = Nothing
Set mYGridLines = Nothing

Set mXGridTextsCoarse = Nothing
Set mXGridTexts = Nothing
Set mXGridLines = Nothing

mCleared = True

Set mToolbar = Nothing

' now remove all graphic object references from the various collections
Dim graphObjSeries As IGraphicObjectSeries
For i = 0 To mGraphObjSeriesIndex - 1
    If Not mGraphObjSeriesTable(i) Is Nothing Then
        mGraphObjSeriesTable(i).Finish
        
        Set t = New GraphObjSeriesRemovalTask
        t.Initialise mGraphObjSeriesTable(i), Me, False
        StartTask t, PriorityLow

        Set mGraphObjSeriesTable(i) = Nothing
    End If
Next
Erase mGraphObjSeriesTable

Set mGraphObjServiceProvider = Nothing
Set mCanvas = Nothing

StartTask Me, PriorityLow

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "clearRegion" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub ClearSelection()
deSelectObjectOrSeries mSelectedObject
Set mSelectedObject = Nothing
End Sub

Public Sub ClearOverrides()
mPropertyOverrideFlags = 0
doRepaint
End Sub

Friend Sub Click()
clickSelectedObjectOrSeries
RaiseEvent Click
End Sub

Friend Sub DblCLick()
dblClickSelectedObjectOrSeries
RaiseEvent DblCLick
End Sub

Friend Sub DrawCursor(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Dim yVal As Single
Dim failpoint As Long
On Error GoTo Err

If Y = MinusInfinitySingle Then
    yVal = mCanvas.Top + 2 * mCanvas.Height  ' make sure the pointer isn't visible
Else
    yVal = Y
End If

If Round(X) >= YAxisPosition Then Exit Sub

hideCursor
showCursor Round(X), yVal

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "drawCursor" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Function FormatYValue(ByVal Y As Double) As String
FormatYValue = generateYAxisCursorPositionText(Y)
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
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "hideGrid" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub Initialise( _
                ByVal Name As String, _
                ByVal pChart As Chart, _
                ByVal pCanvas As Canvas)
mName = Name
Set mChart = pChart
Set mCanvas = pCanvas

setupCrosshairs

End Sub

Public Function IsGridHidden() As Boolean
IsGridHidden = mHideGrid
End Function

Friend Sub MouseDown(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
mCurrX = X
mCurrY = Y
If mPointerMode = PointerModeSelection Then setSelectedObjectOrSeries
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
                
Friend Sub mouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
mCurrX = X
mCurrY = Y
RaiseEvent mouseMove(Button, Shift, X, Y)
End Sub

Friend Sub MouseUp(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
mCurrX = X
mCurrY = Y
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Function NewDimension(ByVal X As Double, _
                        ByVal Y As Double) As Dimension
Dim failpoint As Long
On Error GoTo Err

Set NewDimension = mCanvas.NewDimension(X, Y)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "NewDimension" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function NewPoint(ByVal X As Double, _
                        ByVal Y As Double, _
                        Optional ByVal coordSystemX As CoordinateSystems = CoordsLogical, _
                        Optional ByVal coordSystemY As CoordinateSystems = CoordinateSystems.CoordsLogical) As Point
Dim failpoint As Long
On Error GoTo Err

Set NewPoint = mCanvas.NewPoint(X, Y, coordSystemX, coordSystemY)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "NewPoint" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Friend Sub PaintRegion()
doRepaint
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
mCanvas.Right = maxValue '+ 0.5

If Not mIsDrawingSuppressed Then
    If mUseDeferredPainting Then
        deferredPaint DeferredPaintCalculateAndRepaint
    Else
        determineInScopeObjects
        If Autoscaling Then
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
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "periodsInView" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
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

If undraw Then undrawArea mGraphicObjects(Handle).boundingRect, mGraphicObjects(Handle).graphObj

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

If removeFromSeries Then mGraphObjSeriesTable(mGraphicObjects(Handle).SeriesID).RemoveObject graphObj

Set mGraphicObjects(Handle).graphObj = Nothing
graphObj.Handle = -1

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RemoveGraphicObjectEx" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
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
t.Initialise pSeries, Me, True
StartTask t, PriorityLow
Set mGraphObjSeriesTable(pSeries.SeriesID) = Nothing
End Sub

Friend Sub resizedY()
' notifies the region that the physical Height of the canvas has changed.
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
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "resizedY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
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
    SetVerticalScale .Bottom + factor * .Height / 2, _
                                .Top - factor * .Height / 2
End With

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "scaleUp" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
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

Autoscaling = False
SetVerticalScale mCanvas.Bottom + amount, mCanvas.Top + amount

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "scrollVertical" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
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

Autoscaling = False
SetVerticalScale mCanvas.Bottom + proportion * mCanvas.Height, _
                mCanvas.Top + proportion * mCanvas.Height

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "scrollVerticalProportion" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
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

If mIsDrawingSuppressed Then Exit Sub


If Color = 0 Then Exit Sub

With mCanvas
    .SetPenAttributes Color, 1, LineInsideSolid, DrawModeCopyPen
    .SetBrushAttributes Color, FillSolid
    .DrawRectangle .NewPoint(PeriodNumber - 0.5, .Bottom), .NewPoint(PeriodNumber + 0.5, .Top)

    
    periodrect.Left = PeriodNumber - 0.5
    periodrect.Right = PeriodNumber + 0.5
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
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "setPeriodBackgroundColor" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub SetPointerModeDefault()
If mPointerMode = PointerModeDefault Then Exit Sub

If Not mIsDrawingSuppressed Then
    ClearSelection
    hideCursor
End If

mPointerMode = PointerModeDefault

If Not mIsDrawingSuppressed Then
    setMouseIcon mPointerStyle
    showCursor mPrevCursorX, mPrevCursorY
End If
End Sub

Friend Sub SetPointerModeSelection()
If mPointerMode = PointerModeSelection Then Exit Sub

If Not mIsDrawingSuppressed Then hideCursor

mPointerMode = PointerModeSelection

If Not mIsDrawingSuppressed Then
    setMouseIcon PointerStyles.PointerSelection
    showCursor mPrevCursorX, mPrevCursorY
End If
End Sub

Friend Sub SetPointerModeTool( _
                ByVal toolPointerStyle As PointerStyles, _
                ByVal icon As IPictureDisp)

If Not mIsDrawingSuppressed Then hideCursor

mPointerMode = PointerModeTool
mPointerToolStyle = toolPointerStyle
Set mPointerToolIcon = icon

If Not mIsDrawingSuppressed Then
    setMouseIcon mPointerToolStyle
    showCursor mPrevCursorX, mPrevCursorY
End If
End Sub

' Note that the arguments are ByRef because they may be adjusted
' if they are the same when expressed as Singles.
Public Sub SetVerticalScale(ByRef lowValue As Double, _
                            ByRef highValue As Double)

Dim failpoint As Long
On Error GoTo Err

If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "setVerticalScale", _
            "Cannot set the scale for an axis region"
End If

Autoscaling = False

SetVerticalScaleInternal lowValue, highValue

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "setVerticalScale" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription



End Sub

Friend Sub SetVerticalScaleInternal(ByRef lowValue As Double, _
                            ByRef highValue As Double)

Dim failpoint As Long
On Error GoTo Err

setCanvasVerticalBounds lowValue, highValue

If Not mIsDrawingSuppressed Then
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
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "setVerticalScaleInternal" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Sub ShowGrid()
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
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ShowGrid" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
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
            mGraphicObjects(index).graphObj.Timestamp ' in case the period hasn't been added yet
    
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
                ByVal SeriesID As Long) As Long
Dim Handle As Long

Handle = allocateGraphicObjectTableEntry
If mFirstGraphicObjectIndex = 0 Then
    mFirstGraphicObjectIndex = Handle
End If

graphObj.Handle = Handle

With mGraphicObjects(Handle)
    Set .graphObj = graphObj
    .SeriesID = SeriesID
    .boundingRect = graphObj.BoundingRectangle
    .ExtendedObject = graphObj.ExtendedObject
    .GaugeDependent = graphObj.GaugeDependent
    .IncludeInAutoscale = graphObj.IncludeInAutoscale
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
Dim vGridText As Text
Dim vGridTextCoarse As Text

If HasGrid Then
    On Error Resume Next
    Set vGridLine = mXGridLines.Item(Key)
    On Error GoTo 0
    If vGridLine Is Nothing Then
        Set vGridLine = mXGridLines.Add(Key)
        vGridLine.SetPosition NewPoint(PeriodNumber - 0.5, _
                                    0, _
                                    CoordsLogical, _
                                    CoordsRelative), _
                                NewPoint(PeriodNumber - 0.5, _
                                    100, _
                                    CoordsLogical, _
                                    CoordsRelative)
    Else
        If vGridLine.Point1.XLogical > PeriodNumber Then
            vGridLine.SetPosition NewPoint(PeriodNumber - 0.5, _
                                        0, _
                                        CoordsLogical, _
                                        CoordsRelative), _
                                    NewPoint(PeriodNumber - 0.5, _
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
        vGridText.FixedX = False
        vGridText.FixedY = True
        vGridText.position = NewPoint(PeriodNumber - 0.5, _
                                    0#, _
                                    CoordsLogical, _
                                    CoordsDistance)
        vGridText.Text = label
    End If

    On Error Resume Next
    Set vGridTextCoarse = mXGridTextsCoarse.Item(coarseKey)
    On Error GoTo 0
    If vGridTextCoarse Is Nothing Then
        Set vGridTextCoarse = mXGridTextsCoarse.Add(coarseKey)
        vGridTextCoarse.FixedX = False
        vGridTextCoarse.FixedY = True
        vGridTextCoarse.position = NewPoint(PeriodNumber - 0.5, _
                                    0#, _
                                    CoordsLogical, _
                                    CoordsCounterDistance)
        vGridTextCoarse.Text = coarseLabel
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
                            ByVal Timestamp As Date) As Long
calculateMinutesFromMidnight = Int((Timestamp - Int(Timestamp) + 1 / 86400000) * 1440)
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
                            ByVal Timestamp As Date) As Long
calculateSecondsFromMidnight = Int((Timestamp - Int(Timestamp) + 1 / 86400000) * 86400)
End Function

Private Sub calculateVerticalBounds()
Dim scaleMax As Double
Dim scaleMin As Double
Dim someGaugeDependentObjects As Boolean

Dim failpoint As Long
On Error GoTo Err

If Not Autoscaling Then Exit Sub

scaleMax = MinusInfinityDouble
scaleMin = PlusInfinityDouble

If calcVerticalBoundsForNonGaugeDependentObjects(scaleMin, scaleMax) Then
    ' there are gauge-dependent objects that may affect the max and min scales.
    ' We need to recalculate their bounds with the new gauge determined from
    ' the other objects to get a better estimate of their sizes.
    setCanvasVerticalBounds scaleMin, scaleMax
    calcVerticalBoundsForGaugeDependentObjects scaleMin, scaleMax
End If

If scaleMin = PlusInfinityDouble Or scaleMax = MinusInfinityDouble Then Exit Sub

setCanvasVerticalBounds scaleMin - (scaleMax - scaleMin) * 0.025, _
                        scaleMax + (scaleMax - scaleMin) * 0.025

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "calculateVerticalBounds" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource & vbCrLf & ".Bottom=" & CStr(scaleMin - (scaleMax - scaleMin) * 0.025)
Err.Raise errNumber, errSource, errDescription
End Sub

Private Sub calcVerticalBoundsForGaugeDependentObjects( _
                ByRef scaleMin As Double, _
                ByRef scaleMax As Double)
Dim objMax As Double
Dim objMin As Double
Dim index As Long

index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex

Do While index <> mLastInScopeIndex
    With mGraphicObjects(index)
        If Not .graphObj.NoDraw And _
            .IncludeInAutoscale And _
            .GaugeDependent _
        Then
            .boundingRect = .graphObj.BoundingRectangle
            
            calculateObjectHighAndLowInInterval index, _
                                            rectGetXInterval(mCanvas.Boundary), _
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
End Sub

Private Function calcVerticalBoundsForNonGaugeDependentObjects( _
                ByRef scaleMin As Double, _
                ByRef scaleMax As Double) As Boolean
Dim objMax As Double
Dim objMin As Double
Dim index As Long

index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex

Do While index <> mLastInScopeIndex
    With mGraphicObjects(index)
        If Not .graphObj.NoDraw And _
            .IncludeInAutoscale _
        Then
            If .GaugeDependent Then
                calcVerticalBoundsForNonGaugeDependentObjects = True
            Else
                
                calculateObjectHighAndLowInInterval index, _
                                                rectGetXInterval(mCanvas.Boundary), _
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
End Function

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

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As PropertyOverrideFlags)
mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)
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
ElseIf mGraphObjSeriesTable(mSelectedObject.SeriesID).IsSelectable Then
    mGraphObjSeriesTable(mSelectedObject.SeriesID).Click
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
ElseIf mGraphObjSeriesTable(mSelectedObject.SeriesID).IsSelectable Then
    mGraphObjSeriesTable(mSelectedObject.SeriesID).DblCLick
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
ElseIf mGraphObjSeriesTable(graphObj.SeriesID).IsSelectable Then
    mGraphObjSeriesTable(graphObj.SeriesID).Selected = False
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

Private Sub doRepaint()
If Not mIsDrawingSuppressed Then
    If mUseDeferredPainting Then
        deferredPaint DeferredPaintRepaintRegion
    Else
        paint
    End If
    If Not mYAxisRegion Is Nothing Then mYAxisRegion.PaintRegion
End If
End Sub

Private Sub doScaleAndRepaint()
If Not mIsDrawingSuppressed Then
    If mUseDeferredPainting Then
        deferredPaint DeferredPaintScaleAndRepaint
    Else
        calculateVerticalBounds
        determineVisibleObjects
        paint
    End If
End If
If Not mYAxisRegion Is Nothing Then mYAxisRegion.PaintRegion
End Sub

Private Sub drawCrosshairsPointer(ByVal X As Single, ByVal Y As Single)
Dim lineX As Single

If mIsDrawingSuppressed Then Exit Sub

With mCanvas
'    .SetPenAttributes mPointerCrosshairsColor Xor .BackColor, _
'                        1, _
'                        LineStyles.LineSolid, _
'                        DrawModes.DrawModeXorPen
'    .DrawLine .NewPoint(.Left, Y), .NewPoint(YAxisPosition, Y)
    lineX = Round(X)
'    .DrawLine .NewPoint(lineX, .Bottom), .NewPoint(lineX, .Top)
'
    mCrosshairLineHoriz.SetPosition NewPoint(.Left, Y), _
                                    NewPoint(YAxisPosition, Y)
    mCrosshairLineVert.SetPosition NewPoint(lineX, .Bottom), _
                                    NewPoint(lineX, .Top)
    
End With

mPrevCursorX = X
mPrevCursorY = Y

If Not mYAxisRegion Is Nothing Then
    mYCursorText.position = mYAxisRegion.NewPoint(15, Y, CoordsLogical, CoordsLogical)
    mYCursorText.Text = generateYAxisCursorPositionText(Y)
End If
End Sub

Private Sub drawCustomPointer(ByVal X As Single, ByVal Y As Single)
Dim lineX As Single
Dim size As Dimension

If mIsDrawingSuppressed Then Exit Sub

mPrevCursorX = X
mPrevCursorY = Y
    
If Not mYAxisRegion Is Nothing Then
    mYCursorText.position = mYAxisRegion.NewPoint(15, Y, CoordsLogical, CoordsLogical)
    mYCursorText.Text = generateYAxisCursorPositionText(Y)
End If
End Sub

Private Sub drawDiscPointer(ByVal X As Single, ByVal Y As Single)
Dim lineX As Single
Dim size As Dimension

If mIsDrawingSuppressed Then Exit Sub

With mCanvas
    '.MousePointer = vbCrosshair
    lineX = Round(X)
    Set size = .NewDimension(0.4, 0.4)
    
    .SetPenAttributes vbBlack Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .DrawLine .NewPoint(lineX - size.XLogical, Y), .NewPoint(lineX + size.XLogical, Y)
    .DrawLine .NewPoint(lineX, Y - size.YLogical), .NewPoint(lineX, Y + size.YLogical)
    
    .SetPenAttributes mPointerDiscColor Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .SetBrushAttributes mPointerDiscColor Xor .BackColor, FillStyles.FillSolid
    .DrawCircle .NewPoint(lineX, Y), .NewDimension(0.8, 0.8).XLogical
    
    mPrevCursorX = X
    mPrevCursorY = Y
    
End With

If Not mYAxisRegion Is Nothing Then
    mYCursorText.position = mYAxisRegion.NewPoint(15, Y, CoordsLogical, CoordsLogical)
    mYCursorText.Text = generateYAxisCursorPositionText(Y)
End If
End Sub

Private Sub drawSelectionPointer(ByVal X As Single, ByVal Y As Single)
Dim lineX As Single
Dim size As Dimension

If mIsDrawingSuppressed Then Exit Sub

' nothing to draw since the pointer is a cursor and Windows draws it
mPrevCursorX = X
mPrevCursorY = Y
    
If Not mYAxisRegion Is Nothing Then
    mYCursorText.position = mYAxisRegion.NewPoint(15, Y, CoordsLogical, CoordsLogical)
    mYCursorText.Text = generateYAxisCursorPositionText(Y)
End If
End Sub

Private Sub drawToolPointer(ByVal X As Single, ByVal Y As Single)
Dim lineX As Single
Dim size As Dimension

If mIsDrawingSuppressed Then Exit Sub

With mCanvas
    lineX = Round(X)
    Set size = .NewDimension(0.4, 0.4)
    
'    .SetPenAttributes vbBlack Xor .BackColor, _
'                        1, _
'                        LineStyles.LineSolid, _
'                        DrawModes.DrawModeXorPen
'    .DrawLine .NewPoint(lineX - size.XLogical, Y), .NewPoint(lineX + size.XLogical, Y)
'    .DrawLine .NewPoint(lineX, Y - size.YLogical), .NewPoint(lineX, Y + size.YLogical)
    
mCrosshairLineHoriz.SetPosition .NewPoint(lineX - size.XLogical, Y), .NewPoint(lineX + size.XLogical, Y)
mCrosshairLineVert.SetPosition .NewPoint(lineX, Y - size.YLogical), .NewPoint(lineX, Y + size.YLogical)
End With

mPrevCursorX = X
mPrevCursorY = Y

If Not mYAxisRegion Is Nothing Then
    mYCursorText.position = mYAxisRegion.NewPoint(15, Y, CoordsLogical, CoordsLogical)
    mYCursorText.Text = generateYAxisCursorPositionText(Y)
End If
End Sub

Private Sub drawObject( _
                ByVal graphObj As IGraphicObject, _
                ByRef areaToRedraw As TRectangle)
Dim drawAreas(0) As TRectangle
drawAreas(0) = mCanvas.Boundary
drawAreas(0).Right = drawAreas(0).Right + 1

mCanvas.SetClippingRegion areaToRedraw
graphObj.Draw drawAreas
mCanvas.ClearClippingRegion
End Sub

Private Function generateYAxisCursorPositionText(ByVal Y As Single) As String
Dim fract As Single
Dim num32nds As Long

If YScaleQuantum = 0.03125 Then
    fract = Y - Int(Y)
    num32nds = Int(fract / 0.03125)
    If mScaleGridSpacingY > 4 * 0.03125 Then
        generateYAxisCursorPositionText = Int(Y) & " ' " & Format(num32nds, "00")
    Else
        generateYAxisCursorPositionText = Int(Y) & " ' " & Format(num32nds, "00") & "." & Format(100 * (fract - num32nds * 0.03125) / 0.03125, "00")
    End If
Else
    generateYAxisCursorPositionText = Format(Y, mYScaleSubFormatStr)
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
            ' size the array so it is unlikely to need to be Extended
            
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
            ' size the array so it is unlikely to need to be Extended
    
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
            ' size the array so it is unlikely to need to be Extended
    
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

Private Function getPerformanceText() As Text
Static perfText As Text
Dim perffont As StdFont

If perfText Is Nothing Then
    Set perffont = New StdFont
    perffont.Name = "Lucida Console"
    perffont.size = 8
    perffont.Underline = False
    perffont.Bold = False
    
    Set perfText = AddText()
    perfText.Color = vbActiveTitleBarText
    perfText.Font = perffont
    perfText.Box = True
    perfText.BoxColor = vbBlack
    perfText.BoxStyle = LineStyles.LineInsideSolid
    perfText.BoxThickness = 1
    perfText.BoxFillColor = vbActiveTitleBar
    perfText.BoxFillStyle = FillStyles.FillSolid
    perfText.position = NewPoint(5, 2, CoordsRelative, CoordsRelative)
    perfText.FixedX = True
    perfText.FixedY = True
    perfText.Align = TextAlignModes.AlignBottomLeft
    perfText.IncludeInAutoscale = False
    perfText.PaddingX = 0.5
    perfText.Layer = LayerTitle
End If
Set getPerformanceText = perfText
End Function

Private Function graphicObjectIsInScope(ByVal index As Long) As Boolean
With mGraphicObjects(index)
    If Not .boundingRect.isValid Then Exit Function
    If .ExtendedObject Then
        graphicObjectIsInScope = intOverlaps(rectGetXInterval(mCanvas.Boundary), rectGetXInterval(.boundingRect))
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
        overlapXInterval = intIntersection(rectGetXInterval(mCanvas.Boundary), rectGetXInterval(.boundingRect))
        For i = Int(overlapXInterval.startValue) To _
                Int(overlapXInterval.endValue)
            boundingRects = .graphObj.BoundingRectanglesAt(i)
            numrects = -1
            On Error Resume Next
            numrects = UBound(boundingRects)
            On Error GoTo 0
            If numrects >= 0 Then
                For j = 0 To numrects
                    If rectOverlaps(boundingRects(j), mCanvas.Boundary) Then
                        graphicObjectIsVisible = True
                        Exit Function
                    End If
                Next
            End If
        Next
    Else
        graphicObjectIsVisible = rectOverlaps(.boundingRect, mCanvas.Boundary)
    End If
End With
End Function

Private Sub hideCursor()
Dim lPointerStyle As PointerStyles

mUseDeferredPainting = False

If mPointerMode = PointerModeTool Then
    lPointerStyle = mPointerToolStyle
ElseIf mPointerMode = PointerModeSelection Then
    lPointerStyle = PointerSelection
Else
    lPointerStyle = mPointerStyle
End If

Select Case lPointerStyle
Case PointerNone

Case PointerCrosshairs
    undrawCrosshairsPointer
Case PointerDisc
    undrawDiscPointer
Case PointerSelection
    undrawSelectionPointer
Case PointerTool
    undrawToolPointer
Case PointerCustom
    undrawCustomPointer
End Select

mUseDeferredPainting = True
End Sub

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
ElseIf mGraphObjSeriesTable(graphObj.SeriesID).IsSelectable Then
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
Dim counter As Long
Dim index As Long

Dim failpoint As Long
On Error GoTo Err

Static repaintCount As Long
repaintCount = repaintCount + 1
Debug.Print "Region " & RegionNumber & " repaint count = " & repaintCount

With mCanvas

    If mPerformanceTextVisible Then mElapsedTimer.StartTiming
    
    For i = IIf(mMinInViewPeriod < 0, 0, mMinInViewPeriod) To IIf(mMaxPeriodNumber < mMaxInViewPeriod, mMaxPeriodNumber, mMaxInViewPeriod)
        If mPeriodTable(i).BackColor <> 0 Then
            .SetPenAttributes mPeriodTable(i).BackColor, 1, LineInsideSolid, DrawModeCopyPen
            .SetBrushAttributes mPeriodTable(i).BackColor, FillSolid
            .DrawRectangle .NewPoint(i - 0.5, .Bottom), .NewPoint(i + 0.5, .Top)
        End If
    Next
    
    numGridDivisions = .HeightCm / GridlineSpacingY
    gridSpacing = .Height / numGridDivisions
    
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
            ' Tick size is 1/32, which influences where we put the horizontal grid lines
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
    
    .PaintBackground
    
    If HasGrid And Not mHideGrid Then
        paintHorizontalGridLines
    End If
    
    setHorizontalGridTexts
    
    Dim drawAreas(0) As TRectangle
    drawAreas(0) = .Boundary
    drawAreas(0).Right = drawAreas(0).Right + 1
    For i = MinLayer To MaxLayer
        index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
        Do While index <> mVisibleTable(i).lastindex
            With mGraphicObjects(index)
                mCanvas.SetClippingRegion .boundingRect
                .graphObj.Draw drawAreas
                mCanvas.ClearClippingRegion
                index = .nextVisibleIndex
                counter = counter + 1
            End With
        Loop
    Next
    
    showCursor mPrevCursorX, mPrevCursorY
    
    paintPerformanceText counter
    
    .ZOrder 0
    
End With

mDrawn = True

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "paint" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Private Sub paintHorizontalGridLines()
Dim Y As Double
Dim i As Long
Dim aLine As ChartSkil26.Line

If mYGridLines Is Nothing Then Exit Sub

Y = Int(mCanvas.Bottom / mScaleGridSpacingY) * mScaleGridSpacingY
Do While Y < mCanvas.Top
    i = i + 1
    If i > mYGridLines.Count Then
        Set aLine = mYGridLines.Add
    Else
        Set aLine = mYGridLines.Item(i)
    End If
    aLine.SetPosition Me.NewPoint(0, Y, CoordsRelative, CoordsLogical), _
                        Me.NewPoint(100, Y, CoordsRelative, CoordsLogical)
    mGraphicObjects(aLine.Handle).boundingRect = mGraphicObjects(aLine.Handle).graphObj.BoundingRectangle
    addEntryToVisibleList aLine.Handle
    Y = Y + mScaleGridSpacingY
    drawObject aLine, mCanvas.Boundary
Loop
For i = i + 1 To mYGridLines.Count
    Set aLine = mYGridLines.Item(i)
    aLine.SetPosition Me.NewPoint(-1000, 0, CoordsLogical, CoordsLogical), _
                        Me.NewPoint(-1000, 0, CoordsLogical, CoordsLogical)
    removeEntryFromVisibleList aLine.Handle
Next
End Sub

Private Sub paintPerformanceText( _
                ByVal numberOfObjectsRepainted As Long)
Dim perfText As Text
If mPerformanceTextVisible Then
    Set perfText = getPerformanceText
    perfText.Text = "Repaint time: " & Format(mElapsedTimer.ElapsedTimeMicroseconds / 1000, "0.0") & "ms" & vbCrLf & _
                    "Visible objects: " & numberOfObjectsRepainted & vbCrLf & _
                    "Total objects: " & mNextGraphicObjectIndex
    ' need to force the Text to draw now
    perfText.Draw
End If
End Sub

Private Sub redrawArea( _
                ByRef areaToRedraw As TRectangle, _
                ByVal startLayer As Long, _
                Optional ByVal objToExclude As IGraphicObject)
Dim Layer As Long
Dim index As Long

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

' redraw objects in the undrawn area, preserving layering
If Not mIsDrawingSuppressed And areaToRedraw.isValid Then
    
    Dim rects(0) As TRectangle
    rects(0) = mCanvas.Boundary
    rects(0).Right = rects(0).Right + 1
    failpoint = 200

    mCanvas.SetClippingRegion areaToRedraw
    
    failpoint = 300
    
    For Layer = startLayer To MaxLayer
        
        failpoint = 400
        
        index = mGraphicObjects(mVisibleTable(Layer).firstIndex).nextVisibleIndex
        Do While index <> mVisibleTable(Layer).lastindex
            With mGraphicObjects(index)
                
                failpoint = 500
                
                If Not .graphObj Is objToExclude Then
                    If rectOverlaps(.boundingRect, areaToRedraw) Then
                        
                        failpoint = 600
                        
                        'Dim rects(0) As TRectangle
                        'rects(0) = mCanvas.Boundary
                        .graphObj.Draw rects
                    End If
                End If
                index = .nextVisibleIndex
            End With
        Loop
    Next
    
    
    failpoint = 700
    
    'showCursor mPrevCursorX, mPrevCursorY
    
    failpoint = 800
    
    mCanvas.ClearClippingRegion
    
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "redrawArea" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Private Sub redrawObject( _
                ByVal graphObj As IGraphicObject, _
                ByRef areaToRedraw As TRectangle)
'unDrawCursor areaToRedraw
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
ElseIf mGraphObjSeriesTable(graphObj.SeriesID).IsSelectable Then
    mGraphObjSeriesTable(graphObj.SeriesID).Selected = True
    selectObjectOrSeries = True
Else
    selectObjectOrSeries = False
End If
End Function

Private Sub setAutoscaling()
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "autoScale", _
            "Cannot set autoscale for an axis region"
End If

If Not mToolbar Is Nothing Then
    mToolbar.Buttons("autoscale").Enabled = IIf(Autoscaling, False, True)
End If
    
If Autoscaling Then doScaleAndRepaint

RaiseEvent AutoscalingChanged
End Sub

Private Sub setBackColor()
mCanvas.BackColor = BackColor
doRepaint
End Sub

Private Sub setBackGradientFillColors()
mCanvas.GradientFillColors = BackGradientFillColors
doRepaint
End Sub

Private Sub setBarModeButtons()
If mToolbar Is Nothing Then Exit Sub

Select Case mFirstBarSeries.Style.DisplayMode
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

Private Sub setCanvasVerticalBounds( _
                ByRef Bottom As Double, _
                ByRef Top As Double)
    
If Bottom > Top Then Exit Sub
    
With mCanvas
    If MinimumHeight <> 0 Then
        If Top - Bottom < MinimumHeight Then
            Top = (Top + Bottom) / 2 + MinimumHeight / 2
            Bottom = Top - MinimumHeight
        End If
    ElseIf Bottom = Top Then
        Top = Top * 1.001
        Bottom = Top * 0.998
    End If
    .Top = Top
    .Bottom = Bottom
    If Not mYAxisRegion Is Nothing Then
        mYAxisRegion.Top = .Top
        mYAxisRegion.Bottom = .Bottom
    End If
    ' return possibly adjusted values
    Top = .Top
    Bottom = .Bottom
End With
End Sub

Private Sub setGridlineSpacingY()
If mIsXAxisRegion Or mIsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "gridlineSpacingY", _
            "Cannot set gridline spacing for an axis region"
End If

doRepaint
End Sub

Private Sub setGridLineStyle()
setHasGrid
End Sub

Private Sub setGridTextStyle()
setHasGridText
End Sub

Private Sub setHasGrid()
If HasGrid Then
    Set mXGridLines = Me.AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerGrid)
    mXGridLines.Style = GridLineStyle
    
    Set mYGridLines = Me.AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerGrid)
    mYGridLines.Style = GridLineStyle
Else
    Set mXGridLines = Nothing
    Set mYGridLines = Nothing
End If
doRepaint
End Sub

Private Sub setHasGridText()
Dim aFont As StdFont

If HasGridText Then
    Set mXGridTexts = Me.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText)
    mXGridTexts.Style = GridTextStyle
    ' override the Style for the following, but don't change the Style itself
    
    mXGridTexts.Align = AlignBottomCentre
    If Not mIsXAxisRegion Then
        mXGridTexts.Box = True
        mXGridTexts.BoxFillColor = BackColor
        mXGridTexts.BoxStyle = LineInvisible
        mXGridTexts.BoxFillStyle = FillSolid
    End If
    mXGridTexts.Extended = True

    Set mXGridTextsCoarse = Me.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText)
    mXGridTextsCoarse.Style = GridTextStyle
    
    ' override the Style for the following, but don't change the Style itself
    mXGridTextsCoarse.Align = AlignTopCentre
    If Not mIsXAxisRegion Then
        mXGridTextsCoarse.Box = True
        mXGridTextsCoarse.BoxFillColor = BackColor
        mXGridTextsCoarse.BoxStyle = LineInvisible
        mXGridTextsCoarse.BoxFillStyle = FillSolid
    End If
    mXGridTextsCoarse.Extended = True
Else
    Set mXGridTexts = Nothing
    Set mXGridTextsCoarse = Nothing
End If
doRepaint
End Sub

Private Sub setHorizontalGridTexts()
Dim Y As Double
Dim i As Long
Dim aText As Text

If mYGridTexts Is Nothing Then Exit Sub

Y = Int(mCanvas.Bottom / mScaleGridSpacingY) * mScaleGridSpacingY
Do While Y < mCanvas.Top
    i = i + 1
    If i > mYGridTexts.Count Then
        Set aText = mYGridTexts.Add
    Else
        Set aText = mYGridTexts.Item(i)
    End If
    If YScaleQuantum = 0.03125 Then
        ' Tick size is 1/32 - label accordingly
        aText.Text = Int(Y) & " ' " & Format((Y - Int(Y)) / 0.03125, "00")
    Else
        aText.Text = Format(Y, mYScaleFormatStr)
    End If
    aText.position = mYAxisRegion.NewPoint(15, Y, CoordsLogical, CoordsLogical)
    Y = Y + mScaleGridSpacingY
Loop
For i = i + 1 To mYGridTexts.Count
    Set aText = mYGridTexts.Item(i)
    aText.Text = ""
Next
End Sub

Private Sub setIntegerYScale()
doRepaint
End Sub

Private Sub setMinimumHeight()
doScaleAndRepaint
End Sub

Private Sub setMouseIcon( _
                ByVal PointerStyle As PointerStyles)
mCanvas.MousePointer = vbCustom
Select Case PointerStyle
Case PointerNone
    mCanvas.MouseIcon = gBlankCursor
Case PointerCrosshairs
    mCanvas.MouseIcon = gBlankCursor
Case PointerDisc
    mCanvas.MouseIcon = gBlankCursor
Case PointerSelection
    mCanvas.MouseIcon = gSelectorCursor
Case PointerTool
    mCanvas.MouseIcon = gBlankCursor
Case PointerCustom
    mCanvas.MouseIcon = mPointerIcon
Case PointerWindowsDefault
    mCanvas.MousePointer = vbDefault
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26.ChartRegion::(let)pointerStyle", _
            "Value must be a member of the PointerStyles enum"
End Select
End Sub

Private Sub setSelectedObjectOrSeries()
Dim startLayer As Long
Dim currLayer As Long
Dim i As Long
Dim index As Long
Dim currObjStillSelectable As Boolean
Dim newSelectedObject As IGraphicObject
Dim found As Boolean

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
                        ' we've found an object/series to select
                        
                        ' NB we can't use Exit For because, because we are
                        ' within a With statement, and it causes the table lock
                        ' not to be released - see http://support.microsoft.com/kb/187553
                        GoTo ExitFor
                        
                    End If
                End If
            End If
            index = .nextVisibleIndex
        End With
    Loop
Next

ExitFor:

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

Private Sub setCursorSnapsToTickBoundaries()
' nothing to do
End Sub

Private Sub setPropertyOverrideFlag( _
                ByVal flag As PropertyOverrideFlags)
mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)
End Sub

Private Sub setupCrosshairs()
Set mCrosshairLineHoriz = AddLine(LayerNumbers.LayerPointer)
mCrosshairLineHoriz.Color = vbRed
mCrosshairLineHoriz.Extended = True
mCrosshairLineHoriz.FixedX = True
mCrosshairLineHoriz.FixedY = True
mCrosshairLineHoriz.IncludeInAutoscale = False
mCrosshairLineHoriz.LineStyle = LineSolid
mCrosshairLineHoriz.Thickness = 1

Set mCrosshairLineVert = AddLine(LayerNumbers.LayerPointer)
mCrosshairLineVert.Color = vbRed
mCrosshairLineVert.Extended = True
mCrosshairLineVert.FixedX = True
mCrosshairLineVert.FixedY = True
mCrosshairLineVert.IncludeInAutoscale = False
mCrosshairLineVert.LineStyle = LineSolid
mCrosshairLineVert.Thickness = 1
End Sub

Private Sub setupTitle()
Dim failpoint As Long
On Error GoTo Err

Set mTitle = AddText(, LayerTitle)
mTitle.Box = True
mTitle.BoxColor = vbBlack
mTitle.BoxStyle = LineStyles.LineInvisible
mTitle.BoxThickness = 1
mTitle.BoxFillColor = vbWhite
mTitle.BoxFillStyle = FillStyles.FillSolid
mTitle.position = NewPoint(0.2, 0.2, CoordsDistance, CoordsCounterDistance)
mTitle.FixedX = True
mTitle.FixedY = True
mTitle.Align = TextAlignModes.AlignTopLeft
mTitle.IncludeInAutoscale = False
mTitle.PaddingX = 0.5

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "setupTitle" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource


End Sub

Private Sub setVerticalGridLine(ByVal Timestamp As Date, _
                            ByVal PeriodNumber As Long)
Dim vertGridInterval As Long
Dim gridLineTime As Date
Dim mins As Long
Dim secs As Long

If CDbl(Timestamp) = 0 Then Exit Sub

If mVerticalGridTimePeriod Is Nothing Then
    ' just draw a vertical gridline every 10 bars
    If PeriodNumber Mod 10 = 1 Then
        addVerticalGridLine PeriodNumber, _
                    Format(Timestamp, "yyyymmddhhnnss"), _
                    FormatDateTime(Timestamp, vbShortTime), _
                    Format(Timestamp, "yyyymmdd"), _
                    Format(Timestamp, "d Mmm yy")
    End If
    Exit Sub
End If

gridLineTime = BarStartTime(Timestamp, mVerticalGridTimePeriod)
    
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

Private Sub setYAxisTextStyle()
If Not mYGridTexts Is Nothing Then mYGridTexts.Style = YAxisTextStyle
End Sub

Private Sub setYCursorTextStyle()
If Not mYCursorText Is Nothing Then mYCursorText.LocalStyle = YCursorTextStyle
End Sub

Private Sub setYScaleQuantum()
doRepaint
End Sub

Private Sub showCursor( _
                ByVal X As Single, _
                ByVal Y As Single)
Dim PointerStyle As PointerStyles

mUseDeferredPainting = False

If mPointerMode = PointerModeTool Then
    PointerStyle = mPointerToolStyle
ElseIf mPointerMode = PointerModeSelection Then
    PointerStyle = PointerSelection
Else
    PointerStyle = mPointerStyle
End If

Select Case PointerStyle
Case PointerNone

Case PointerCrosshairs
    drawCrosshairsPointer X, Y
Case PointerDisc
    drawDiscPointer X, Y
Case PointerSelection
    drawSelectionPointer X, Y
Case PointerTool
    drawToolPointer X, Y
Case PointerCustom
    drawCustomPointer X, Y
End Select

mUseDeferredPainting = True
End Sub

Private Function isPropertyOverrideFlagSet( _
                ByVal flag As PropertyOverrideFlags) As Boolean
isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)
End Function

Private Sub undrawArea( _
                ByRef undrawnArea As TRectangle, _
                ByVal objToExclude As IGraphicObject)
If Not mIsDrawingSuppressed And undrawnArea.isValid Then
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
        undrawAreaEx undrawnArea, objToExclude
    End If
End If

End Sub

Private Sub undrawAreaEx( _
                ByRef undrawnArea As TRectangle, _
                Optional ByVal objToExclude As IGraphicObject)

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

' redraw objects in the undrawn area, preserving layering
If Not mIsDrawingSuppressed And undrawnArea.isValid Then
    
    failpoint = 200

    mCanvas.SetClippingRegion undrawnArea
    mCanvas.PaintBackground
    mCanvas.ClearClippingRegion
    
    failpoint = 400
    
    redrawArea undrawnArea, MinLayer, objToExclude
    
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "undrawAreaEx" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Private Sub unDrawCursor( _
                ByRef areaToRedraw As TRectangle)

mCanvas.SetClippingRegion areaToRedraw
hideCursor
mCanvas.ClearClippingRegion
End Sub

Private Sub undrawCrosshairsPointer()
Dim prevLineX As Single

If mIsDrawingSuppressed Then Exit Sub

mCrosshairLineHoriz.SetPosition NewPoint(-1000, 0), _
                                NewPoint(-1000, 0)
mCrosshairLineVert.SetPosition NewPoint(-1000, 0), _
                                NewPoint(-1000, 0)

'With mCanvas
'    .SetPenAttributes mPointerCrosshairsColor Xor .BackColor, _
'                        1, _
'                        LineStyles.LineSolid, _
'                        DrawModes.DrawModeXorPen
'    .DrawLine .NewPoint(.Left, mPrevCursorY), .NewPoint(YAxisPosition, mPrevCursorY)
'    prevLineX = Round(mPrevCursorX)
'    .DrawLine .NewPoint(prevLineX, .Bottom), .NewPoint(prevLineX, .Top)
'End With

End Sub

Private Sub undrawCustomPointer()

' nothing to do

End Sub

Private Sub undrawDiscPointer()
Dim size As Dimension
Dim prevLineX As Single

If mIsDrawingSuppressed Then Exit Sub

With mCanvas
    prevLineX = Round(mPrevCursorX)
    Set size = .NewDimension(0.4, 0.4)
    
    .SetPenAttributes vbBlack Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .DrawLine .NewPoint(prevLineX - size.XLogical, mPrevCursorY), .NewPoint(prevLineX + size.XLogical, mPrevCursorY)
    .DrawLine .NewPoint(prevLineX, mPrevCursorY - size.YLogical), .NewPoint(prevLineX, mPrevCursorY + size.YLogical)
    
    .SetPenAttributes mPointerDiscColor Xor .BackColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .SetBrushAttributes mPointerDiscColor Xor .BackColor, FillStyles.FillSolid
    .DrawCircle .NewPoint(prevLineX, mPrevCursorY), .NewDimension(0.8, 0.8).XLogical
End With

End Sub

Private Sub undrawSelectionPointer()

If mIsDrawingSuppressed Then Exit Sub

' nothing to do since the pointer is actually a proper cursor
End Sub

Private Sub undrawToolPointer()
Dim size As Dimension
Dim prevLineX As Single

If mIsDrawingSuppressed Then Exit Sub

mCrosshairLineHoriz.SetPosition NewPoint(-1000, 0), _
                                NewPoint(-1000, 0)
mCrosshairLineVert.SetPosition NewPoint(-1000, 0), _
                                NewPoint(-1000, 0)
'With mCanvas
'    prevLineX = Round(mPrevCursorX)
'    Set size = .NewDimension(0.4, 0.4)
'
'    .SetPenAttributes vbBlack Xor .BackColor, _
'                        1, _
'                        LineStyles.LineSolid, _
'                        DrawModes.DrawModeXorPen
'    .DrawLine .NewPoint(prevLineX - size.XLogical, mPrevCursorY), .NewPoint(prevLineX + size.XLogical, mPrevCursorY)
'    .DrawLine .NewPoint(prevLineX, mPrevCursorY - size.YLogical), .NewPoint(prevLineX, mPrevCursorY + size.YLogical)
'
'End With

End Sub


