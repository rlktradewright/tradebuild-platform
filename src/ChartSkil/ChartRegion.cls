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
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type PeriodTableEntry
    firstIndex                  As Long
    lastindex                   As Long
    backColor                   As Long
End Type

Private Type GraphicObjectTableEntry
    graphObj            As IGraphicObject
    seriesId            As Long
    boundingRect        As TRectangle
    periodNumber        As Long
    extendedObject      As Boolean
    includeInAutoscale  As Boolean
    scaleDependent      As Boolean
    gaugeDependent      As Boolean
    noDraw              As Boolean
    capabilities        As Long
    layer               As Long
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

Private Type VisibleTableEntry
    firstIndex          As Long
    lastindex           As Long
End Type

Public Enum LayerNumbers
    LayerBackground = MinLayer
    LayerGrid = MinLayer + 5
    LayerGridText = LayerGrid + 1
    LayerLowestUser = MinLayer + 10
    LayerHIghestUser = MaxLayer - 10
    LayerDefaultLines
    LayerDefaultTexts
    LayerTitle = MaxLayer - 5
    LayerPointer = MaxLayer
End Enum


'================================================================================
' Member variables
'================================================================================

Private mController As ChartController

Private mName As String
Private mRegionNumber As Long

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

Private mBarSeriesCol As Collection
Private mDataPointSeriesCol As Collection
Private mLineSeriesCol As Collection
Private mTextSeriesCol As Collection

Private mFirstBarSeries As BarSeries

Private mDefaultTextSeries As TextSeries
Private mDefaultLineSeries As LineSeries

'Private mSurface As PictureBox
Private mCanvas As canvas
Private mSuppressDrawing As Boolean

Private WithEvents mToolbar As toolbar
Attribute mToolbar.VB_VarHelpID = -1

'Private mGaugeX As Double
'Private mGaugeY As Double

Private mMinInViewPeriod As Long
Private mMaxInViewPeriod As Long

Private mNextLayer As Long

Private mDrawn As Boolean

Private WithEvents mEventProxy As ChartRegionEventProxy
Attribute mEventProxy.VB_VarHelpID = -1

Private mYAxisRegion As ChartRegion
Private mYPositionText As text      ' used to indicate the position of the pointer
                                    ' in the Y axis region
Private mYGridTexts As TextSeries
Private mYGridLines As LineSeries

Private mXGridTextsCoarse As TextSeries
Private mXGridTexts As TextSeries
Private mXGridLines As LineSeries

' indicates whether there is a grid, not whether it
' is currently visible
Private mhasGridText As Boolean

' indicates whether the grid is currently
' hidden
Private mHideGrid As Boolean


Private mSessionOffsetMinutes As Date
Private mVerticalGridSpacing As Long
Private mVerticalGridUnits As TimePeriodUnits

Private mAutoscale As Boolean
Private mMinimumHeight As Double
'Private mPrevHeight As Single
Private mPercentheight As Double
Private mMinimumPercentHeight As Double
Private mYAxisPosition As Long
Private mIntegerYScale As Boolean
Private mYScaleQuantum As Double

'Private mRegionRect As TRectangle

Private mGridlineSpacingY As Double
Private mScaleGridSpacingY As Double
Private mYScaleFormatStr As String
Private mYScaleSubFormatStr As String
Private mGridTextHeight As Double

Private mBackColor As Long
Private mGridColor As Long
Private mGridTextColor As Long
Private mhasGrid As Boolean
Private mShowPerformanceText As Boolean

Private mPointerStyle As PointerStyles
Private mPointerCrosshairsColor As Long
Private mPointerDiscColor As Long

Private mPrevCursorX As Single
Private mPrevCursorY As Single

Private mPerfText As text

Private mCurrentTool As ToolTypes

' defaults for bar series
Private mDefaultBarStyle As BarStyle

' defaults for datapoint series
Private mDefaultDataPointStyle As DataPointStyle

' defaults for text and text series
Private mDefaultTextStyle As TextStyle

' defaults for lines and line series
Private mDefaultLineStyle As linestyle

Private mElapsedTimer As ElapsedTimer

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Dim i As Long

Set mCanvas = New canvas
mNextLayer = LayerNumbers.LayerLowestUser

ReDim mPeriodTable(200) As PeriodTableEntry
ReDim mGraphicObjects(500) As GraphicObjectTableEntry

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

Set mEventProxy = New ChartRegionEventProxy

mGridlineSpacingY = GridlineSpacingCm
Debug.Print "Region " & mName & ": GridlineSpacingY=" & mGridlineSpacingY

Set mBarSeriesCol = New Collection
Set mDataPointSeriesCol = New Collection
Set mLineSeriesCol = New Collection
Set mTextSeriesCol = New Collection

End Sub

Private Sub Class_Terminate()
Debug.Print "ChartRegion terminated"
End Sub

'================================================================================
' mEventProxy Event Handlers
'================================================================================

Private Sub mEventProxy_addGraphicObject( _
                ByVal graphObj As IGraphicObject, _
                ByVal seriesId As Long)
addGraphicObject graphObj, seriesId
End Sub

Private Sub mEventProxy_objectChanged( _
                ByVal handle As Long, _
                ByRef areaToRedraw As TRectangle)
Dim layer As Long
Dim index As Long
Dim overlaps() As TRectangle
Dim redrawArea As TRectangle

redrawArea = rectIntersection(areaToRedraw, mCanvas.boundary)

With mGraphicObjects(handle)
    
    ' get the new properties for the changed object
    
    .boundingRect = .graphObj.boundingRectangle(False, _
                                        False, _
                                        False, _
                                        mCanvas.gaugeX, _
                                        mCanvas.gaugeY, _
                                        mCanvas.boundary)
    
    If Not .boundingRect.isValid Then
        removeEntryFromInScopeList handle
        removeEntryFromVisibleList handle
    End If
    
    .extendedObject = .graphObj.extendedObject
    .noDraw = .graphObj.noDraw
    .layer = .graphObj.layer
    .includeInAutoscale = .graphObj.includeInAutoscale
    .scaleDependent = .graphObj.scaleDependent
    .capabilities = .graphObj.capabilities
End With

If mGraphicObjects(handle).extendedObject Then
    removeEntryFromPeriodList handle
    addEntryToExtendedList handle
Else
    removeEntryFromExtendedList handle
    
    ' The following call to addEntryToPeriodList must not be in a With block
    ' involving mGraphicObjects because in rare circumstances it can result
    ' in mGraphicObjects being ReDim'd which leads to an error 10
    addEntryToPeriodList handle
End If

If graphicObjectIsInScope(handle) Then
    addEntryToInScopeList handle
    If graphicObjectIsVisible(handle) Then
        addEntryToVisibleList handle
    End If
End If
    
If Not mGraphicObjects(handle).boundingRect.isValid Then Exit Sub

With mGraphicObjects(handle)
    ' check if the vertical bounds have changed, and redraw if so
    If mAutoscale And Not mSuppressDrawing Then
        If .includeInAutoscale Then
            If .boundingRect.top >= mCanvas.boundary.top Or .boundingRect.bottom <= (mCanvas.boundary.bottom) Then
                calculateVerticalBounds
                determineVisibleObjects
                paintRegion
                Exit Sub
            End If
        End If
    End If
    
    ' now redraw the changed object
    If Not mSuppressDrawing And Not .noDraw Then
        If graphicObjectIsInScope(handle) Then
            addEntryToInScopeList handle
            If graphicObjectIsVisible(handle) Then
                addEntryToVisibleList handle
            Else
                Exit Sub
            End If
        End If
    
        If redrawArea.isValid Then
            Dim drawAreas(0) As TRectangle
            drawAreas(0) = redrawArea
            .graphObj.draw drawAreas, mCanvas.hdc
        End If
    End If
    
    ' now redraw any overlapping objects in higher layers than the changed object
    If Not mSuppressDrawing And Not .noDraw And redrawArea.isValid Then
        For layer = .layer + 1 To MaxLayer
            index = mGraphicObjects(mVisibleTable(layer).firstIndex).nextVisibleIndex
            Do While index <> mVisibleTable(layer).lastindex
                With mGraphicObjects(index)
                    If index <> handle Then
                        If getObjectOverlaps(handle, index, redrawArea, True, True, overlaps) <> 0 Then
                            .graphObj.draw overlaps, mCanvas.hdc
                        End If
                    End If
                    index = .nextVisibleIndex
                End With
            Loop
        Next
    End If
End With

End Sub

Private Sub mEventProxy_objectUndrawn(ByVal handle As Long, undrawnArea As ChartSkilTypes.TRectangle)
Dim layer As Long
Dim index As Long
Dim overlaps() As TRectangle
Dim blockUndraw As Boolean

With mGraphicObjects(handle)
    blockUndraw = .capabilities And capabilities.blockUndraw
    ' redraw objects in the undrawn area, preserving layering
    If Not mSuppressDrawing And undrawnArea.isValid Then
        For layer = MinLayer To MaxLayer
            index = mGraphicObjects(mVisibleTable(layer).firstIndex).nextVisibleIndex
            Do While index <> mVisibleTable(layer).lastindex
                With mGraphicObjects(index)
                    If index <> handle Then
                        If getObjectOverlaps(handle, index, undrawnArea, Not blockUndraw, True, overlaps) <> 0 Then
                            .graphObj.draw overlaps, mCanvas.hdc
                        End If
                    End If
                    index = .nextVisibleIndex
                End With
            Loop
        Next
    End If
End With
End Sub

Private Sub mEventProxy_paintRegion( _
                ByVal recalculate As Boolean)

If mSuppressDrawing Then Exit Sub

If recalculate Then
    determineInScopeObjects
    If mAutoscale Then
        calculateVerticalBounds
    End If
    determineVisibleObjects
End If
paintRegion
End Sub

Private Sub mEventProxy_recalcBoundary( _
                ByVal seriesId As Long)
Dim i As Long
Dim index As Long

For i = MinLayer To MaxLayer
    index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(i).lastindex
        With mGraphicObjects(index)
            If .seriesId = seriesId Then
                .boundingRect = .graphObj.boundingRectangle(False, _
                                                            False, _
                                                            False, _
                                                            0, _
                                                            0, _
                                                            mCanvas.boundary)
            End If
            index = .nextVisibleIndex
        End With
    Loop
Next
End Sub

'================================================================================
' mToolbar Event Handlers
'================================================================================

Private Sub mToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.key
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
            Button.Enabled = False
        End If
    ElseIf mFirstBarSeries.displayMode = BarDisplayModeBar Then
        If mFirstBarSeries.barThickness > 1 Then
            mFirstBarSeries.barThickness = mFirstBarSeries.barThickness - 1
        End If
        If mFirstBarSeries.barThickness = 1 Then
            Button.Enabled = False
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
    scaleUp -0.09091
Case ToolbarCommandScaleUp
    scaleUp 0.1
Case ToolbarCommandScrollDown
    scrollVerticalProportion -0.2
Case ToolbarCommandScrollUp
    scrollVerticalProportion 0.2
Case ToolbarCommandAutoScale
    autoscale = True
End Select

End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get autoscale() As Boolean
autoscale = mAutoscale
End Property

Public Property Let autoscale(ByVal value As Boolean)
mAutoscale = value
If Not mToolbar Is Nothing Then
    mToolbar.Buttons("autoscale").Enabled = IIf(mAutoscale, False, True)
End If
If Not mSuppressDrawing Then
    If mAutoscale Then
        calculateVerticalBounds
        determineVisibleObjects
        paintRegion
    End If
End If
End Property

'Friend Property Get barWidth() As Single
'barWidth = mBarWidth
'End Property
'
'Friend Property Let barWidth(ByVal value As Single)
'mBarWidth = value
'End Property
'
'Public Property Get barDisplayMode() As BarDisplayModes
'barDisplayMode = mDisplayMode
'End Property
'
'Public Property Let barDisplayMode(ByVal value As BarDisplayModes)
'mDisplayMode = value
'End Property
'
'Public Property Get barDownColor() As Long
'barDownColor = mBarDownColor
'End Property
'
'Public Property Let barDownColor(ByVal value As Long)
'mBarDownColor = value
'End Property
'
'Public Property Get barOutlineThickness() As Long
'barOutlineThickness = mBarOutlineThickness
'End Property
'
'Public Property Let barOutlineThickness(ByVal value As Long)
'mBarOutlineThickness = value
'End Property
'
'Public Property Get barSolidUpBody() As Boolean
'barSolidUpBody = mBarSolidUpBody
'End Property
'
'Public Property Let barSolidUpBody(ByVal value As Boolean)
'mBarSolidUpBody = value
'End Property
'
'Public Property Get barTailThickness() As Long
'barTailThickness = mBarTailThickness
'End Property
'
'Public Property Let barTailThickness(ByVal value As Long)
'mBarTailThickness = value
'End Property
'
'Public Property Get barThickness() As Long
'barThickness = mBarThickness
'End Property
'
'Public Property Let barThickness(ByVal value As Long)
'mBarThickness = value
'End Property
'
'Public Property Get barUpColor() As Long
'barUpColor = mBarUpColor
'End Property
'
'Public Property Let barUpColor(ByVal value As Long)
'mBarUpColor = value
'End Property

Friend Property Let controller( _
                ByVal value As ChartController)
Set mController = value
End Property

Public Property Get currentTool() As ToolTypes
currentTool = mCurrentTool
End Property

Friend Property Let currentTool(ByVal value As ToolTypes)
mCurrentTool = value
End Property

Public Property Get defaultBarStyle() As BarStyle
Set defaultBarStyle = mDefaultBarStyle.clone
End Property

Public Property Let defaultBarStyle(ByVal value As BarStyle)
Set mDefaultBarStyle = value.clone
End Property

Public Property Get defaultDataPointStyle() As DataPointStyle
Set defaultDataPointStyle = mDefaultDataPointStyle.clone
End Property

Public Property Let defaultDataPointStyle(ByVal value As DataPointStyle)
Set mDefaultDataPointStyle = value.clone
End Property

Public Property Get defaultLineStyle() As linestyle
Set defaultLineStyle = mDefaultLineStyle.clone
End Property

Public Property Let defaultLineStyle(ByVal value As linestyle)
Set mDefaultLineStyle = value.clone
End Property

Public Property Get defaultTextStyle() As TextStyle
Set defaultTextStyle = mDefaultTextStyle.clone
End Property

Public Property Let defaultTextStyle(ByVal value As TextStyle)
Set mDefaultTextStyle = value.clone
End Property

Public Property Get gridColor() As Long
gridColor = mGridColor
End Property

Public Property Let gridColor(ByVal val As Long)
mGridColor = val
End Property

Public Property Get gridlineSpacingY() As Double
gridlineSpacingY = mGridlineSpacingY
End Property

Public Property Let gridlineSpacingY(ByVal value As Double)
mGridlineSpacingY = value
Debug.Print "Region " & mName & ": GridlineSpacingY=" & mGridlineSpacingY
End Property

Public Property Get gridTextColor() As Long
gridTextColor = mGridTextColor
End Property

Public Property Let gridTextColor(ByVal val As Long)
mGridTextColor = val
End Property

Public Property Get hasGrid() As Boolean
hasGrid = mhasGrid
End Property

Public Property Let hasGrid(ByVal val As Boolean)
mhasGrid = val
If mhasGrid Then
    Set mXGridLines = Me.addLineSeries(LayerNumbers.LayerGrid)
    mXGridLines.Color = mGridColor
    mXGridLines.extendAfter = True
    mXGridLines.extendBefore = True
    mXGridLines.includeInAutoscale = False
    mXGridLines.linestyle = LineSolid
    mXGridLines.thickness = 1
    
    Set mYGridLines = Me.addLineSeries(LayerNumbers.LayerGrid)
    mYGridLines.Color = mGridColor
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

Public Property Get hasGridText() As Boolean
hasGridText = mhasGridText
End Property

Public Property Let hasGridText(ByVal val As Boolean)
Dim aFont As StdFont

mhasGridText = val
If mhasGridText Then
    Set mXGridTexts = Me.addTextSeries(LayerNumbers.LayerGridText)
    mXGridTexts.align = AlignBottomCentre
    mXGridTexts.box = True
    mXGridTexts.boxFillColor = mBackColor
    mXGridTexts.boxStyle = LineInvisible
    mXGridTexts.extended = True
    mXGridTexts.Color = mGridTextColor
    Set aFont = New StdFont
    aFont.name = "Arial"
    aFont.Size = 8
    aFont.Underline = False
    aFont.Bold = False
    mXGridTexts.font = aFont

    Set mXGridTextsCoarse = Me.addTextSeries(LayerNumbers.LayerGridText)
    mXGridTextsCoarse.align = AlignTopCentre
    mXGridTextsCoarse.box = True
    mXGridTextsCoarse.boxFillColor = mBackColor
    mXGridTextsCoarse.boxStyle = LineInvisible
    mXGridTextsCoarse.extended = True
    mXGridTextsCoarse.Color = mGridTextColor
    Set aFont = New StdFont
    aFont.name = "Arial"
    aFont.Size = 8
    aFont.Underline = False
    aFont.Bold = False
    mXGridTextsCoarse.font = aFont
Else
    Set mXGridTexts = Nothing
End If
End Property

Public Property Let integerYScale(ByVal value As Boolean)
mIntegerYScale = value
End Property

Public Property Get minimumHeight() As Double
minimumHeight = mMinimumHeight
End Property

Public Property Let minimumHeight(ByVal value As Double)
mMinimumHeight = value
End Property

Public Property Get minimumPercentHeight() As Double
minimumPercentHeight = mMinimumPercentHeight
End Property

Friend Property Let minimumPercentHeight(ByVal value As Double)
mMinimumPercentHeight = value
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

Public Property Let PointerCrosshairsColor(ByVal value As Long)
mPointerCrosshairsColor = value
End Property

Public Property Get PointerDiscColor() As Long
PointerDiscColor = mPointerDiscColor
End Property

Public Property Let PointerDiscColor(ByVal value As Long)
mPointerDiscColor = value
End Property

Public Property Get pointerStyle() As PointerStyles
pointerStyle = mPointerStyle
End Property

Public Property Let pointerStyle(ByVal value As PointerStyles)

If mPointerStyle = value Then Exit Property

If Not mDrawn Then
    mPointerStyle = value
    Exit Property
End If

Select Case mPointerStyle
Case PointerNone

Case PointerCrosshairs
    undrawCrosshairs
Case PointerDisc
    undrawPointer
End Select

mPointerStyle = value

Select Case mPointerStyle
Case PointerNone

Case PointerCrosshairs
    mCanvas.MousePointer = vbDefault
    drawCrosshairs mPrevCursorX, mPrevCursorY
Case PointerDisc
    mCanvas.MousePointer = vbCrosshair
    drawPointer mPrevCursorX, mPrevCursorY
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26.ChartRegion::(let)pointerStyle", _
            "Value must be a member of the PointerStyles enum"
End Select

End Property

Public Property Get regionBackColor() As Long
regionBackColor = mBackColor
End Property

Public Property Let regionBackColor(ByVal val As Long)
mBackColor = val
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.regionBackColor = val
End If
End Property

Public Property Get regionBottom() As Single
regionBottom = mCanvas.boundary.bottom
End Property

Friend Property Let regionBottom(ByVal value As Single)
mCanvas.bottom = value
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.regionBottom = value
End If
End Property

Public Property Get regionHeight() As Single
regionHeight = mCanvas.height
End Property

'Friend Property Let regionHeight(ByVal value As Single)
'If value < mMinimumHeight Then
'    mCanvas.height = mMinimumHeight
'Else
'    mCanvas.height = value
'End If
'End Property

Public Property Get regionLeft() As Single
regionLeft = mCanvas.left
End Property

Friend Property Let regionLeft(ByVal value As Single)
mCanvas.left = value
End Property

Public Property Get regionNumber() As Long
regionNumber = mRegionNumber
End Property

Public Property Let regionNumber(ByVal val As Long)
mRegionNumber = val
End Property

Public Property Get regionTop() As Single
regionTop = mCanvas.top
End Property

Friend Property Let regionTop(ByVal value As Single)
mCanvas.top = value
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.regionTop = value
End If
End Property

Public Property Get regionWidth() As Single
regionWidth = mCanvas.width
End Property

'Friend Property Let regionWidth(ByVal value As Single)
'mCanvas.width = value
'End Property

Public Property Get scaleGridSpacingY() As Single
scaleGridSpacingY = mScaleGridSpacingY
End Property

Friend Property Let sessionStartTime(ByVal val As Date)
mSessionOffsetMinutes = calculateMinutesFromMidnight(val) Mod 60
End Property

Public Property Get showPerformanceText() As Boolean
showPerformanceText = mShowPerformanceText
End Property

Public Property Let showPerformanceText(ByVal val As Boolean)
mShowPerformanceText = val
If mShowPerformanceText Then
    If mElapsedTimer Is Nothing Then Set mElapsedTimer = New ElapsedTimer
Else
    Set mElapsedTimer = Nothing
End If
End Property

Public Property Get style() As ChartRegionStyle
Set style = mStyle.clone
End Property

Public Property Let style( _
                ByVal value As ChartRegionStyle)
Dim lSuppress As Boolean

lSuppress = mSuppressDrawing
mSuppressDrawing = True

If value.autoscale <> mAutoscale Then autoscale = value.autoscale
If value.backColor <> mBackColor Then regionBackColor = value.backColor
If value.gridColor <> mGridColor Then gridColor = value.gridColor
If value.gridlineSpacingY <> mGridlineSpacingY Then gridlineSpacingY = value.gridlineSpacingY
If value.gridTextColor <> mGridTextColor Then gridTextColor = value.gridTextColor
If value.integerYScale <> mIntegerYScale Then integerYScale = value.integerYScale
If value.minimumHeight <> mMinimumHeight Then minimumHeight = value.minimumHeight
If value.pointerStyle <> mPointerStyle Then pointerStyle = value.pointerStyle
If value.hasGrid <> mhasGrid Then hasGrid = value.hasGrid
If value.hasGridText <> mhasGridText Then hasGridText = value.hasGridText
If value.YScaleQuantum <> mYScaleQuantum Then YScaleQuantum = value.YScaleQuantum

Set mStyle = value
mSuppressDrawing = lSuppress
End Property

Friend Property Let suppressDrawing(ByVal val As Boolean)
If mSuppressDrawing = val Then Exit Property
mSuppressDrawing = val
If Not mSuppressDrawing Then
    determineInScopeObjects
    If mAutoscale Then
        calculateVerticalBounds
    End If
    determineVisibleObjects
    paintRegion
End If
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.suppressDrawing = val
End If
End Property

Friend Property Let surface(ByVal value As PictureBox)
mCanvas.surface = value
End Property

Public Property Get title() As String
title = mTitle.text
End Property

Friend Property Let toolbar(ByVal value As toolbar)
Dim btn As Button

Set mToolbar = value

For Each btn In mToolbar.Buttons
    If btn.key = "autoscale" Then btn.Enabled = IIf(mAutoscale, False, True)
    If btn.key = "showbars" Then btn.value = tbrPressed
    If btn.key = "showline" Then btn.Enabled = False
                                        ' because line charts are not implemented yet
Next

End Property

Friend Property Let verticalGridSpacing(ByVal value As Long)
mVerticalGridSpacing = value
End Property

Friend Property Let verticalGridUnits(ByVal value As TimePeriodUnits)
mVerticalGridUnits = value
End Property

Public Property Get YAxisPosition() As Long
YAxisPosition = mYAxisPosition
End Property

Friend Property Let YAxisRegion(ByVal value As ChartRegion)
Dim aFont As StdFont

Set mYAxisRegion = value
mYAxisRegion.regionBackColor = regionBackColor
mYAxisRegion.hasGrid = False
mYAxisRegion.periodsInView 0, 99
mYAxisRegion.pointerStyle = PointerNone
mYAxisRegion.suppressDrawing = mSuppressDrawing

mYAxisRegion.defaultTextStyle = mDefaultTextStyle

Set mYPositionText = mYAxisRegion.addText(LayerNumbers.LayerPointer)
mYPositionText.align = AlignCentreLeft
mYPositionText.Color = vbWhite Xor mBackColor
mYPositionText.box = True
mYPositionText.boxFillColor = mBackColor
mYPositionText.boxStyle = LineSolid
mYPositionText.boxColor = vbWhite Xor mBackColor
Set aFont = New StdFont
aFont.name = "Arial"
aFont.Size = 8
aFont.Underline = False
aFont.Bold = False
mYPositionText.font = aFont

Set mYGridTexts = mYAxisRegion.addTextSeries(LayerNumbers.LayerGrid)
mYGridTexts.includeInAutoscale = False
mYGridTexts.align = AlignCentreLeft
mYGridTexts.Color = mGridTextColor
mYGridTexts.box = False
Set aFont = New StdFont
aFont.name = "Arial"
aFont.Size = 8
aFont.Underline = False
aFont.Bold = False
mYGridTexts.font = aFont

End Property

Public Property Let YScaleQuantum(ByVal value As Double)
mYScaleQuantum = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function addBarSeries( _
                Optional ByVal layer As Long = -1, _
                Optional ByVal name As String, _
                Optional ByVal style As BarStyle) As BarSeries
Static seriesNumber As Long
Set addBarSeries = New BarSeries
If mFirstBarSeries Is Nothing Then Set mFirstBarSeries = addBarSeries
If name = "" Then
    seriesNumber = seriesNumber + 1
    name = "BarSeries" & seriesNumber
End If
mBarSeriesCol.add addBarSeries, name
addBarSeries.initialise name, _
                        mController, _
                        mEventProxy, _
                        getLayer(layer), _
                        mCanvas, _
                        IIf(style Is Nothing, mDefaultBarStyle, style)
End Function

Public Function addDataPointSeries( _
                Optional ByVal layer As Long = -1, _
                Optional ByVal name As String, _
                Optional ByVal style As DataPointStyle) As DataPointSeries
Static seriesNumber As Long
Set addDataPointSeries = New DataPointSeries
If name = "" Then
    seriesNumber = seriesNumber + 1
    name = "DataPointSeries" & seriesNumber
End If
mDataPointSeriesCol.add addDataPointSeries, name
addDataPointSeries.initialise name, _
                        mController, _
                        mEventProxy, _
                        getLayer(layer), _
                        mCanvas, _
                        IIf(style Is Nothing, mDefaultDataPointStyle, style)
End Function

Public Function addLine(Optional ByVal layer As Long = -1) As ChartSkil26.Line
If mDefaultLineSeries Is Nothing Then
    Set mDefaultLineSeries = addLineSeries(LayerDefaultLines)
End If
Set addLine = mDefaultLineSeries.add
If layer <> -1 Then addLine.layer = layer
End Function

Public Function addLineSeries( _
                Optional ByVal layer As Long = -1, _
                Optional ByVal name As String, _
                Optional ByVal style As linestyle) As LineSeries
Static seriesNumber As Long
Set addLineSeries = New LineSeries
If name = "" Then
    seriesNumber = seriesNumber + 1
    name = "LineSeries" & seriesNumber
End If
mLineSeriesCol.add addLineSeries, name
addLineSeries.initialise name, _
                        mEventProxy, _
                        getLayer(layer), _
                        mCanvas, _
                        IIf(style Is Nothing, mDefaultLineStyle, style)
End Function

Friend Sub addPeriod(ByVal periodNumber As Long, _
                    ByVal timestamp As Date)
addNewPeriod periodNumber
setVerticalGridLine timestamp, periodNumber
End Sub

Public Function addText(Optional ByVal layer As Long = -1) As text
If mDefaultTextSeries Is Nothing Then
    Set mDefaultTextSeries = addTextSeries(LayerDefaultTexts)
End If
Set addText = mDefaultTextSeries.add
If layer <> -1 Then addText.layer = layer
End Function

Public Function addTextSeries( _
                Optional ByVal layer As Long = -1, _
                Optional ByVal name As String, _
                Optional ByVal style As TextStyle) As TextSeries
Static seriesNumber As Long
Set addTextSeries = New TextSeries
If name = "" Then
    seriesNumber = seriesNumber + 1
    name = "TextSeries" & seriesNumber
End If
mTextSeriesCol.add addTextSeries, name
addTextSeries.initialise name, _
                        mEventProxy, _
                        getLayer(layer), _
                        mCanvas, _
                        IIf(style Is Nothing, mDefaultTextStyle, style)
End Function

Friend Sub clearRegion()
Dim i As Long

If Not mSuppressDrawing Then mCanvas.clear
mCanvas.surface = Nothing

Set mToolbar = Nothing

' now remove all graphic object references from the various collections
For i = mBarSeriesCol.count To 1 Step -1
    Dim aBarSeries As BarSeries
    Set aBarSeries = mBarSeriesCol(i)
    aBarSeries.removeAll
    aBarSeries.finish
    mBarSeriesCol.remove i
Next

For i = mDataPointSeriesCol.count To 1 Step -1
    Dim aDataPointSeries As DataPointSeries
    Set aDataPointSeries = mDataPointSeriesCol(i)
    aDataPointSeries.removeAll
    aDataPointSeries.finish
    mDataPointSeriesCol.remove i
Next

For i = mLineSeriesCol.count To 1 Step -1
    Dim aLineSeries As LineSeries
    Set aLineSeries = mLineSeriesCol(i)
    aLineSeries.removeAll
    aLineSeries.finish
    mLineSeriesCol.remove i
Next

For i = mTextSeriesCol.count To 1 Step -1
    Dim aTextSeries As TextSeries
    Set aTextSeries = mTextSeriesCol(i)
    aTextSeries.removeAll
    aTextSeries.finish
    mTextSeriesCol.remove i
Next

Erase mGraphicObjects
' now notify each object in the graphic object table that it's finished with
'For i = mFirstGraphicObjectIndex To mNextGraphicObjectIndex - 1
'    If Not mGraphicObjects(i).graphobj Is Nothing Then
'        mGraphicObjects(i).graphobj.finish
'    End If
'    Set mGraphicObjects(i).graphobj = Nothing
'Next

' now clear the YAxis region (if there is one)
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.clearRegion
End If
End Sub

Public Sub displayGrid()
If Not mHideGrid Then Exit Sub

mHideGrid = False
paintRegion
End Sub

Public Sub hideGrid()
If mHideGrid Then Exit Sub

mHideGrid = True
paintRegion
End Sub

Public Function isGridHidden() As Boolean
isGridHidden = mHideGrid
End Function

Friend Sub MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim yVal As Single
If y = MinusInfinitySingle Then
    yVal = mCanvas.top + 2 * mCanvas.height  ' make sure the pointer isn't visible
Else
    yVal = y
End If

If Round(x) >= YAxisPosition Then Exit Sub

Select Case mPointerStyle
Case PointerNone

Case PointerCrosshairs
    undrawCrosshairs
    drawCrosshairs x, yVal
Case PointerDisc
    undrawPointer
    drawPointer x, yVal
End Select

End Sub

Public Function newDimension(ByVal x As Double, _
                        ByVal y As Double) As Dimension
Set newDimension = mCanvas.newDimension(x, y)
End Function

Public Function newPoint(ByVal x As Double, _
                        ByVal y As Double, _
                        Optional ByVal coordSystemX As CoordinateSystems = CoordsLogical, _
                        Optional ByVal coordSystemY As CoordinateSystems = CoordinateSystems.CoordsLogical) As Point
Set newPoint = mCanvas.newPoint(x, y, coordSystemX, coordSystemY)
End Function

Friend Sub paintRegion()
Dim numGridDivisions As Double
Dim gridSpacing As Double
Dim i As Long
Dim z As Double
Dim perfFont As StdFont
Dim counter As Long
Dim index As Long

If mSuppressDrawing Then Exit Sub

Static repaintCount As Long
repaintCount = repaintCount + 1
'debug.print "repaint count = " & repaintCount

With mCanvas

If mShowPerformanceText Then mElapsedTimer.StartTiming

.clear
.backColor = regionBackColor

For i = IIf(mMinInViewPeriod < 0, 0, mMinInViewPeriod) To IIf(mMaxPeriodNumber < mMaxInViewPeriod, mMaxPeriodNumber, mMaxInViewPeriod)
    If mPeriodTable(i).backColor <> 0 Then
        .setPenAttributes mPeriodTable(i).backColor, 1, LineInsideSolid, DrawModeCopyPen
        .setBrushAttributes mPeriodTable(i).backColor, FillSolid
        .drawRectangle i - 0.5, .bottom, i + 0.5, .top
    End If
Next

'Dim gridTimer As New ElapsedTimer
'gridTimer.StartTiming

numGridDivisions = .heightCm / mGridlineSpacingY
gridSpacing = .height / numGridDivisions

mGridTextHeight = .TextHeight("123")

Dim log10 As Double
Dim logSpacing As Double
Dim exp As Long
Dim mantissa As Double

If mhasGrid And Not mHideGrid Then
    log10 = Log(10#)
    logSpacing = Log(gridSpacing) / log10
    
    exp = Fix(logSpacing)
    mantissa = Abs(logSpacing) - Abs(exp)
    
    If logSpacing < 0 Then
        exp = exp - 1
        mantissa = 1 - mantissa
    End If
    
    If mYScaleQuantum = 0.03125 And _
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
    If mIntegerYScale Then
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

If mhasGrid And Not mHideGrid Then
    setHorizontalGridLines
End If

setHorizontalGridTexts

'Debug.Print "Grid generation " & gridTimer.ElapsedTimeMicroseconds / 1000 & " millisecs"

For i = MinLayer To MaxLayer
    index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(i).lastindex
        With mGraphicObjects(index)
            Dim drawAreas(0) As TRectangle
            drawAreas(0) = rectIntersection(.boundingRect, mCanvas.boundary)
            .graphObj.draw drawAreas, mCanvas.hdc
            index = .nextVisibleIndex
            counter = counter + 1
        End With
    Loop
Next

Select Case mPointerStyle
Case PointerNone

Case PointerCrosshairs
    drawCrosshairs mPrevCursorX, mPrevCursorY
Case PointerDisc
    drawPointer mPrevCursorX, mPrevCursorY
End Select

If mShowPerformanceText Then
    If mPerfText Is Nothing Then
        Set perfFont = New StdFont
        perfFont.name = "Lucida Console"
        perfFont.Size = 8
        perfFont.Underline = False
        perfFont.Bold = False
        Set mPerfText = addText()
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
        mPerfText.layer = LayerTitle
        
    End If
    mPerfText.text = Format(mElapsedTimer.ElapsedTimeMicroseconds / 1000, "0.0") & "ms" & vbCrLf & _
                    counter & " visible objects" & vbCrLf & _
                    mNextGraphicObjectIndex & " total objects"
    ' need to force the text to draw now
    mPerfText.draw
End If

End With

mDrawn = True

If Not mYAxisRegion Is Nothing Then mYAxisRegion.paintRegion
End Sub

Friend Sub periodsInView( _
                ByVal minValue As Single, _
                ByVal maxValue As Single)
mMinInViewPeriod = minValue
mMaxInViewPeriod = maxValue
mYAxisPosition = maxValue + 1
mCanvas.left = minValue
mCanvas.right = maxValue '+ 0.5

If Not mSuppressDrawing Then
    determineInScopeObjects
    If mAutoscale Then
        calculateVerticalBounds
    End If
    determineVisibleObjects
End If
End Sub

Public Sub removeBarSeries( _
                ByVal pBarSeries As BarSeries)
removeGraphicObjectSeries pBarSeries
End Sub

Public Sub removeDataPointSeries( _
                ByVal pDataPointSeries As DataPointSeries)
removeGraphicObjectSeries pDataPointSeries
End Sub

Public Sub removeLineSeries( _
                ByVal pLineSeries As LineSeries)
removeGraphicObjectSeries pLineSeries
End Sub

Public Sub removeTextSeries( _
                ByVal pTextSeries As TextSeries)
removeGraphicObjectSeries pTextSeries
End Sub

Friend Sub resizedY()
' notifies the region that the physical height of the canvas has changed.
' This means the Y gauge will have changed, so we need to recalculate
' the position of any in-scope gauge-dependent objects.
mCanvas.resizedY
calculateInscopeGaugeDependentObjects
If Not mYAxisRegion Is Nothing Then mYAxisRegion.resizedY
End Sub

Public Sub scaleUp(ByVal factor As Single)
autoscale = False
With mCanvas
    setVerticalScale .bottom + factor * .height / 2, _
                                .top - factor * .height / 2
End With
End Sub

Public Sub setPeriodBackgroundColor( _
                ByVal periodNumber As Long, _
                ByVal Color As Long)
Dim layer As Long
Dim index As Long
Dim periodrect As TRectangle
Dim redrawAreas(0) As TRectangle

mPeriodTable(periodNumber).backColor = Color

If mSuppressDrawing Then Exit Sub


If Color = 0 Then Exit Sub

With mCanvas
    .setPenAttributes Color, 1, LineInsideSolid, DrawModeCopyPen
    .setBrushAttributes Color, FillSolid
    .drawRectangle periodNumber - 0.5, .bottom, periodNumber + 0.5, .top

    
    periodrect.left = periodNumber - 0.5
    periodrect.right = periodNumber + 0.5
    periodrect.bottom = .bottom
    periodrect.top = .top
    periodrect.isValid = True
    
    For layer = MinLayer To MaxLayer
        index = mGraphicObjects(mVisibleTable(layer).firstIndex).nextVisibleIndex
        Do While index <> mVisibleTable(layer).lastindex
            With mGraphicObjects(index)
                redrawAreas(0) = rectIntersection(periodrect, .boundingRect)
                If redrawAreas(0).isValid Then
                    .graphObj.draw redrawAreas, mCanvas.hdc
                End If
                index = .nextVisibleIndex
            End With
        Loop
    Next
End With
End Sub

Public Sub setTitle(ByVal val As String, _
                        ByVal Color As Long, _
                        ByVal font As StdFont)
If mTitle Is Nothing Then
    Set mTitle = addText(LayerTitle)
End If
mTitle.Color = Color
mTitle.font = font
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
mTitle.text = val
End Sub

Public Sub setVerticalScale(ByVal lowValue As Single, _
                            ByVal highValue As Single)

If mAutoscale Then
    calculateInscopeGaugeDependentObjects
    Exit Sub
End If

mCanvas.top = highValue
mCanvas.bottom = lowValue

If Not mSuppressDrawing Then
    determineVisibleObjects
    paintRegion
End If
If Not mYAxisRegion Is Nothing Then
    mYAxisRegion.setVerticalScale lowValue, highValue
End If
End Sub

Public Sub scrollVertical(ByVal amount As Single)
autoscale = False
setVerticalScale mCanvas.bottom + amount, mCanvas.top + amount
End Sub

Public Sub scrollVerticalProportion(ByVal proportion As Single)
autoscale = False
setVerticalScale mCanvas.bottom + proportion * mCanvas.height, _
                mCanvas.top + proportion * mCanvas.height
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
        .graphObj.inScope = True
        .inScopeCycle = mInScopeCycle
        .nextInScopeIndex = mLastInScopeIndex
        .prevInScopeIndex = mGraphicObjects(mLastInScopeIndex).prevInScopeIndex
        mGraphicObjects(.prevInScopeIndex).nextInScopeIndex = index
        mGraphicObjects(mLastInScopeIndex).prevInScopeIndex = index
    End If
End With
End Sub

Private Sub addEntryToPeriodList(ByVal index As Long)
Dim lastindex As Long
Dim newPeriodNumber As Long

newPeriodNumber = mGraphicObjects(index).graphObj.periodNumber
If newPeriodNumber <> mGraphicObjects(index).periodNumber Then
    If mGraphicObjects(index).periodNumber >= 1 Then
        removeEntryFromPeriodList index
    End If
    If newPeriodNumber < 1 Then Exit Sub
    
    ' NB: the following call to addNewPeriod must not be within a With block
    ' that refers to mGraphicObjects, because is rare circumstances it can
    ' cause mGraphicObjects to be ReDim'd which causes an error 10
    addNewPeriod newPeriodNumber  ' in case the period hasn't been added yet
    
    lastindex = mPeriodTable(newPeriodNumber).lastindex
    With mGraphicObjects(index)
        .periodNumber = newPeriodNumber
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
        .graphObj.visible = True
        .visibleCycle = mVisibleCycle
        lastindex = mVisibleTable(.layer).lastindex
        .nextVisibleIndex = lastindex
        .prevVisibleIndex = mGraphicObjects(lastindex).prevVisibleIndex
        mGraphicObjects(.prevVisibleIndex).nextVisibleIndex = index
        mGraphicObjects(lastindex).prevVisibleIndex = index
    End If
End With
End Sub

Private Function addGraphicObject( _
                ByVal graphObj As IGraphicObject, _
                ByVal seriesId As Long) As Long
Dim handle As Long

handle = allocateGraphicObjectTableEntry
If mFirstGraphicObjectIndex = 0 Then
    mFirstGraphicObjectIndex = handle
End If

graphObj.handle = handle

With mGraphicObjects(handle)
    Set .graphObj = graphObj
    .seriesId = seriesId
    .boundingRect = graphObj.boundingRectangle(True, True, True, 0, 0, mCanvas.boundary)
    .extendedObject = graphObj.extendedObject
    .gaugeDependent = graphObj.gaugeDependent
    .includeInAutoscale = graphObj.includeInAutoscale
    .layer = graphObj.layer
    .scaleDependent = graphObj.scaleDependent
    .capabilities = graphObj.capabilities
    .noDraw = graphObj.noDraw
    
End With

If Not mGraphicObjects(handle).extendedObject Then
    ' NB: the following call to addEntryToPeriodList must not be within a With block
    ' that refers to mGraphicObjects, because in rare circumstances it can
    ' cause mGraphicObjects to be ReDim'd which causes an error 10
    addEntryToPeriodList handle
Else
    addEntryToExtendedList handle
End If

If graphicObjectIsInScope(handle) Then
    addEntryToInScopeList handle
    If graphicObjectIsVisible(handle) Then
        addEntryToVisibleList handle
    End If
End If
addGraphicObject = handle
End Function

Private Sub addNewPeriod(ByVal periodNumber As Long)
If periodNumber > mMaxPeriodNumber Then
    mMaxPeriodNumber = periodNumber
End If
If periodNumber > UBound(mPeriodTable) Then
    ReDim Preserve mPeriodTable(periodNumber + 200) As PeriodTableEntry
End If

With mPeriodTable(periodNumber)
    If .firstIndex = 0 Then
        .firstIndex = allocateGraphicObjectTableEntry
        .lastindex = allocateGraphicObjectTableEntry
        mGraphicObjects(.firstIndex).nextPeriodIndex = .lastindex
        mGraphicObjects(.lastindex).prevPeriodIndex = .firstIndex
    End If
End With
End Sub

Private Sub addVerticalGridLine(ByVal periodNumber As Long, _
                                ByVal key As String, _
                                ByVal label As String, _
                                ByVal coarseKey As String, _
                                ByVal coarseLabel As String)
Dim vGridLine As ChartSkil26.Line
Dim vGridText As text
Dim vGridTextCoarse As text

If mhasGrid Then
    On Error Resume Next
    Set vGridLine = mXGridLines.item(key)
    On Error GoTo 0
    If vGridLine Is Nothing Then
        Set vGridLine = mXGridLines.add(key)
        vGridLine.point1 = newPoint(periodNumber - 0.5, _
                                    0, _
                                    CoordsLogical, _
                                    CoordsRelative)
        vGridLine.point2 = newPoint(periodNumber - 0.5, _
                                    100, _
                                    CoordsLogical, _
                                    CoordsRelative)
    Else
        If vGridLine.point1.XLogical > periodNumber Then
            vGridLine.point1 = newPoint(periodNumber - 0.5, _
                                        0, _
                                        CoordsLogical, _
                                        CoordsRelative)
            vGridLine.point2 = newPoint(periodNumber - 0.5, _
                                        100, _
                                        CoordsLogical, _
                                        CoordsRelative)
        End If
    End If
End If

If mhasGridText Then
    On Error Resume Next
    Set vGridText = mXGridTexts.item(key)
    On Error GoTo 0
    If vGridText Is Nothing Then
        Set vGridText = mXGridTexts.add(key)
        vGridText.fixedX = False
        vGridText.fixedY = True
        vGridText.position = newPoint(periodNumber - 0.5, _
                                    0#, _
                                    CoordsLogical, _
                                    CoordsDistance)
        vGridText.text = label
    End If

    On Error Resume Next
    Set vGridTextCoarse = mXGridTextsCoarse.item(coarseKey)
    On Error GoTo 0
    If vGridTextCoarse Is Nothing Then
        Set vGridTextCoarse = mXGridTextsCoarse.add(coarseKey)
        vGridTextCoarse.fixedX = False
        vGridTextCoarse.fixedY = True
        vGridTextCoarse.position = newPoint(periodNumber - 0.5, _
                                    0#, _
                                    CoordsLogical, _
                                    CoordsCounterDistance)
        vGridTextCoarse.text = coarseLabel
    End If
End If

End Sub

Private Function allocateGraphicObjectTableEntry() As Long
allocateGraphicObjectTableEntry = mNextGraphicObjectIndex

If mNextGraphicObjectIndex > UBound(mGraphicObjects) Then
    ReDim Preserve mGraphicObjects(UBound(mGraphicObjects) + 500) As GraphicObjectTableEntry
End If
mNextGraphicObjectIndex = mNextGraphicObjectIndex + 1
End Function

'Private Function calculateGaugeX() As Double
'calculateGaugeX = (mSurface.ScaleWidth / mSurface.Width) * Screen.TwipsPerPixelX
'End Function
'
'Private Function calculateGaugeY() As Double
'calculateGaugeY = -(mSurface.ScaleHeight / mSurface.height) * Screen.TwipsPerPixelY
'End Function
'

Private Function calculateInscopeGaugeDependentObjects()
Dim index As Long
index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
Do While index <> mLastInScopeIndex
     With mGraphicObjects(index)
        If .gaugeDependent Then
            .boundingRect = .graphObj.boundingRectangle(False, False, True, 0, 0, mCanvas.boundary)
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
    If (.capabilities And capabilities.ComplexBounding) Then
        overlapXInterval = intIntersection(interval, rectGetXInterval(.boundingRect))
        For j = Int(overlapXInterval.startValue + 0.9) To _
                Int(overlapXInterval.endValue)
            rects = .graphObj.boundingRectanglesAt(j)
            numrects = -1
            On Error Resume Next
            numrects = UBound(rects)
            On Error GoTo 0
            If numrects >= 0 Then
                For k = 0 To numrects
                    If rects(k).top > high Then
                        high = rects(k).top
                    End If
                    If rects(k).bottom < low Then
                        low = rects(k).bottom
                    End If
                Next
            End If
        Next
    Else
        high = .boundingRect.top
        low = .boundingRect.bottom
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

If Not mAutoscale Then Exit Sub

scaleMax = MinusInfinityDouble
scaleMin = PlusInfinityDouble

index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
Do While index <> mLastInScopeIndex
    With mGraphicObjects(index)
        If Not .graphObj.noDraw And _
            .includeInAutoscale And _
            Not .scaleDependent _
        Then
            If .gaugeDependent Then
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
    ElseIf mMinimumHeight <> 0 Then
        If scaleMax - scaleMin < mMinimumHeight Then
            .top = (scaleMax + scaleMin) / 2 + mMinimumHeight / 2
            .bottom = .top - mMinimumHeight
        Else
            .top = scaleMax + (scaleMax - scaleMin) * 0.025
            .bottom = scaleMin - (scaleMax - scaleMin) * 0.025
        End If
    Else
        If scaleMin = scaleMax Then
            scaleMax = scaleMax * 1.001
            scaleMin = scaleMax * 0.998
        End If
        .top = scaleMax + (scaleMax - scaleMin) * 0.025
        .bottom = scaleMin - (scaleMax - scaleMin) * 0.025
    End If
    If Not mYAxisRegion Is Nothing Then
        mYAxisRegion.regionTop = .top
        mYAxisRegion.regionBottom = .bottom
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
        If Not .graphObj.noDraw And _
            .includeInAutoscale And _
            Not .scaleDependent And _
            .gaugeDependent _
        Then
            .boundingRect = .graphObj.boundingRectangle( _
                                        False, _
                                        True, _
                                        True, _
                                        0, _
                                        0, _
                                        mCanvas.boundary)
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
    .top = scaleMax + (scaleMax - scaleMin) * 0.025
    .bottom = scaleMin - (scaleMax - scaleMin) * 0.025
    If Not mYAxisRegion Is Nothing Then
        mYAxisRegion.regionTop = .top
        mYAxisRegion.regionBottom = .bottom
    End If
End With
End Sub

Private Sub clearInScopeObjectList()
Dim index As Long

index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
Do While index <> mLastInScopeIndex
    With mGraphicObjects(index)
        .graphObj.inScope = False
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
            .graphObj.visible = False
            index = .nextVisibleIndex
        End With
    Loop
Next

initialiseVisibleObjectList
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
For i = minPeriod To IIf(UBound(mPeriodTable) < mMaxInViewPeriod, UBound(mPeriodTable), mMaxInViewPeriod)
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
        .boundingRect = .graphObj.boundingRectangle(True, False, False, 0, 0, mCanvas.boundary)
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
        .boundingRect = .graphObj.boundingRectangle( _
                                    False, _
                                    True, _
                                    False, _
                                    0, _
                                    0, _
                                    mCanvas.boundary)
        If graphicObjectIsVisible(index) Then
            addEntryToVisibleList index
            numberVisible = numberVisible + 1
        End If
        index = .nextInScopeIndex
    End With
Loop

Debug.Print "DetermineVisibleObjects: total=" & numberExamined & " visible=" & numberVisible & " time=" & 1000 * (Timer - timerIn) & "ms"
End Sub

Private Sub drawCrosshairs(ByVal x As Single, ByVal y As Single)
Dim lineX As Single

If mSuppressDrawing Then Exit Sub

With mCanvas
    .setPenAttributes mPointerCrosshairsColor Xor .backColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .drawLine .left, y, YAxisPosition, y
    lineX = Round(x)
    .drawLine lineX, .bottom, lineX, .top
    mPrevCursorX = x
    mPrevCursorY = y
    
End With

If Not mYAxisRegion Is Nothing Then
    mYPositionText.position = mYAxisRegion.newPoint(5, y, CoordsLogical, CoordsLogical)
    mYPositionText.text = generateYAxisCursorPositionText(y)
End If
End Sub

Private Sub drawPointer(ByVal x As Single, ByVal y As Single)
Dim lineX As Single

If mSuppressDrawing Then Exit Sub

With mCanvas
    .MousePointer = vbCrosshair
    .setPenAttributes mPointerDiscColor Xor .backColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .setBrushAttributes mPointerDiscColor Xor .backColor, FillStyles.FillSolid
    lineX = Round(x)
    .drawCircle lineX, y, .newDimension(0.8, 0.8).XLogical
    mPrevCursorX = x
    mPrevCursorY = y
    
End With

If Not mYAxisRegion Is Nothing Then
    mYPositionText.position = mYAxisRegion.newPoint(5, y, CoordsLogical, CoordsLogical)
    mYPositionText.text = generateYAxisCursorPositionText(y)
End If
End Sub

Private Function generateYAxisCursorPositionText(ByVal y As Single) As String
Dim fract As Single
Dim num32nds As Long

If mYScaleQuantum = 0.03125 Then
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

Private Function getLayer(Optional ByVal layer As Long = -1) As Long
If layer = -1 Then
    getLayer = mNextLayer
    If mNextLayer <> LayerNumbers.LayerHIghestUser Then
        mNextLayer = mNextLayer + 1
    End If
Else
    If layer < MinLayer Or layer > MaxLayer Then
        getLayer = getLayer()
    Else
        getLayer = layer
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
    complexBounding1 = (.capabilities And capabilities.ComplexBounding)
End With

With mGraphicObjects(index2)
    Set graphobj2 = .graphObj
    obj2Rect = rectIntersection(.boundingRect, clippingRect)
    complexBounding2 = (.capabilities And capabilities.ComplexBounding)
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
        
        boundingRects1 = graphobj1.boundingRectanglesAt(i)
        numRects1 = -1
        On Error Resume Next
        numRects1 = UBound(boundingRects1) + 1
        On Error GoTo 0
        
        boundingRects2 = graphobj2.boundingRectanglesAt(i)
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
        
        boundingRects1 = graphobj1.boundingRectanglesAt(i)
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
        
        boundingRects2 = graphobj2.boundingRectanglesAt(i)
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
    If .extendedObject Then
        graphicObjectIsInScope = intOverlaps(rectGetXInterval(mCanvas.boundary), rectGetXInterval(.boundingRect))
    ElseIf .periodNumber >= mMinInViewPeriod And _
            .periodNumber <= mMaxInViewPeriod _
    Then
        graphicObjectIsInScope = True
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
    If Not .boundingRect.isValid Or .noDraw Then Exit Function
    If .extendedObject And (.capabilities And capabilities.ComplexBounding) Then
        overlapXInterval = intIntersection(rectGetXInterval(mCanvas.boundary), rectGetXInterval(.boundingRect))
        For i = Int(overlapXInterval.startValue) To _
                Int(overlapXInterval.endValue)
            boundingRects = .graphObj.boundingRectanglesAt(i)
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

Private Sub initialiseVisibleObjectList()
Dim i As Long

mVisibleCycle = mVisibleCycle + 1

For i = MinLayer To MaxLayer
    mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex = mVisibleTable(i).lastindex
    mGraphicObjects(mVisibleTable(i).lastindex).prevVisibleIndex = mVisibleTable(i).firstIndex
Next

End Sub

Private Sub removeGraphicObjectSeries( _
                ByVal pSeries As IGraphicObjectSeries)
Dim graphObj As IGraphicObject
Dim i As Long

For i = pSeries.count To 1 Step -1
    Set graphObj = pSeries.item(i)
    removeGraphicObject graphObj
    graphObj.finish
    pSeries.remove i
Next
pSeries.finish
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

Private Sub removeGraphicObject( _
                ByVal graphObj As IGraphicObject)
Dim handle As Long

handle = graphObj.handle

If Not mGraphicObjects(handle).extendedObject Then
    removeEntryFromPeriodList handle
Else
    removeEntryFromExtendedList handle
End If

If graphicObjectIsInScope(handle) Then
    removeEntryFromInScopeList handle
    If graphicObjectIsVisible(handle) Then
        removeEntryFromVisibleList handle
    End If
End If

Set mGraphicObjects(handle).graphObj = Nothing
graphObj.handle = -1

End Sub

Private Sub setHorizontalGridLines()
Dim y As Double
Dim i As Long
Dim aLine As ChartSkil26.Line
Dim currSuppress As Boolean

If mYGridLines Is Nothing Then Exit Sub

currSuppress = mSuppressDrawing
mSuppressDrawing = True
y = Int(mCanvas.bottom / mScaleGridSpacingY) * mScaleGridSpacingY
Do While y < mCanvas.top
    i = i + 1
    If i > mYGridLines.count Then
        Set aLine = mYGridLines.add
    Else
        Set aLine = mYGridLines.item(i)
    End If
    aLine.point1 = Me.newPoint(0, y, CoordsRelative, CoordsLogical)
    aLine.point2 = Me.newPoint(100, y, CoordsRelative, CoordsLogical)
    addEntryToVisibleList aLine.handle
    y = y + mScaleGridSpacingY
Loop
For i = i + 1 To mYGridLines.count
    Set aLine = mYGridLines.item(i)
    aLine.point1 = Me.newPoint(-1000, 0, CoordsLogical, CoordsLogical)
    aLine.point2 = Me.newPoint(-1000, 0, CoordsLogical, CoordsLogical)
Next
mSuppressDrawing = currSuppress
End Sub

Private Sub setHorizontalGridTexts()
Dim y As Double
Dim i As Long
Dim aText As text
'Dim currSuppress As Boolean

If mYGridTexts Is Nothing Then Exit Sub

'currSuppress = mSuppressDrawing
'mSuppressDrawing = True
y = Int(mCanvas.bottom / mScaleGridSpacingY) * mScaleGridSpacingY
Do While y < mCanvas.top
    i = i + 1
    If i > mYGridTexts.count Then
        Set aText = mYGridTexts.add
    Else
        Set aText = mYGridTexts.item(i)
    End If
    If mYScaleQuantum = 0.03125 Then
        ' tick size is 1/32 - label accordingly
        aText.text = Int(y) & " ' " & Format((y - Int(y)) / 0.03125, "00")
    Else
        aText.text = Format(y, mYScaleFormatStr)
    End If
    aText.position = mYAxisRegion.newPoint(1, y, CoordsLogical, CoordsLogical)
    y = y + mScaleGridSpacingY
Loop
For i = i + 1 To mYGridTexts.count
    Set aText = mYGridTexts.item(i)
    aText.text = ""
Next
'mSuppressDrawing = currSuppress
End Sub

Private Sub setVerticalGridLine(ByVal timestamp As Date, _
                            ByVal periodNumber As Long)
Dim vertGridInterval As Long
Dim gridLineTime As Date
Dim mins As Long
Dim secs As Long

Debug.Assert mVerticalGridSpacing <> 0

If mVerticalGridUnits = TimePeriodNone Then
    ' just draw a vertical gridline every mVerticalGridSpacing bars
    If periodNumber Mod mVerticalGridSpacing = 1 Then
        addVerticalGridLine periodNumber, _
                    Format(timestamp, "yyyymmddhhnnss"), _
                    FormatDateTime(timestamp, vbShortTime), _
                    Format(timestamp, "yyyymmdd"), _
                    Format(timestamp, "d Mmm yy")
    End If
    Exit Sub
End If

gridLineTime = BarStartTime(timestamp, mVerticalGridSpacing, mVerticalGridUnits)
    
Select Case mVerticalGridUnits
Case TimePeriodSecond
    addVerticalGridLine periodNumber, _
                    Format(gridLineTime, "yyyymmddhhnnss"), _
                    FormatDateTime(gridLineTime, vbLongTime), _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d Mmm yy")
Case TimePeriodMinute
    addVerticalGridLine periodNumber, _
                    Format(gridLineTime, "yyyymmddhhnn"), _
                    FormatDateTime(gridLineTime, vbShortTime), _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d Mmm yy")
Case TimePeriodHour
    addVerticalGridLine periodNumber, _
                    Format(gridLineTime, "yyyymmddhh"), _
                    FormatDateTime(gridLineTime, vbShortTime), _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d Mmm yy")
Case TimePeriodDay
    addVerticalGridLine periodNumber, _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d"), _
                    Format(gridLineTime, "yyyymm"), _
                    Format(gridLineTime, "Mmm yy")
Case TimePeriodWeek
    addVerticalGridLine periodNumber, _
                    Format(gridLineTime, "yyyymmdd"), _
                    Format(gridLineTime, "d Mmm"), _
                    Format(gridLineTime, "yyyy"), _
                    Format(gridLineTime, "yyyy")
Case TimePeriodMonth
    addVerticalGridLine periodNumber, _
                    Format(gridLineTime, "yyyymm"), _
                    Format(gridLineTime, "Mmm"), _
                    Format(gridLineTime, "yyyy"), _
                    Format(gridLineTime, "yyyy")
Case TimePeriodYear
    addVerticalGridLine periodNumber, _
                    Format(gridLineTime, "yyyy"), _
                    Format(gridLineTime, "YYYY"), _
                    "$", _
                    ""
Case TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    If periodNumber Mod mVerticalGridSpacing = 1 Then
        addVerticalGridLine periodNumber, _
                        Format(gridLineTime, "yyyymmddhhnnss"), _
                        FormatDateTime(gridLineTime, vbLongTime), _
                        Format(gridLineTime, "yyyymmdd"), _
                        Format(gridLineTime, "d Mmm yy")
    End If
End Select
End Sub

Private Sub undrawCrosshairs()
Dim nearestGridLine As Single
Dim nextNearestGridLine As Single
Dim prevLineX As Single

If mSuppressDrawing Then Exit Sub

With mCanvas
    .setPenAttributes mPointerCrosshairsColor Xor .backColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .drawLine .left, mPrevCursorY, YAxisPosition, mPrevCursorY
    prevLineX = Round(mPrevCursorX)
    .drawLine prevLineX, .bottom, prevLineX, .top
End With

End Sub

Private Sub undrawPointer()
Dim nearestGridLine As Single
Dim nextNearestGridLine As Single
Dim prevLineX As Single

If mSuppressDrawing Then Exit Sub

With mCanvas
    .setPenAttributes mPointerDiscColor Xor .backColor, _
                        1, _
                        LineStyles.LineSolid, _
                        DrawModes.DrawModeXorPen
    .setBrushAttributes mPointerDiscColor Xor .backColor, FillStyles.FillSolid
    prevLineX = Round(mPrevCursorX)
    .drawCircle prevLineX, mPrevCursorY, .newDimension(0.8, 0.8).XLogical
End With

End Sub


