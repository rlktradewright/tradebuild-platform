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
    firstIndex          As Long
    lastindex           As Long
End Type

Private Type GraphicObjectTableEntry
    graphObj            As IGraphicObject
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
    LayerLowestUser = MinLayer + 10
    LayerHIghestUser = MaxLayer - 10
    LayerTitle = MaxLayer - 5
    LayerPointer = MaxLayer
End Enum


'================================================================================
' Member variables
'================================================================================

Private mRegionNumber As Long

Private mPeriodTable() As PeriodTableEntry

Private mGraphicObjects() As GraphicObjectTableEntry
Private mGraphicObjectsCount As Long

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
Private mTextObjects As Collection
Private mLineObjects As Collection

Private mSurface As PictureBox
Private mSuppressDrawing As Boolean

Private mGaugeX As Double
Private mGaugeY As Double

Private mMinInViewPeriod As Long
Private mMaxInViewPeriod As Long

Private mNextLayer As Long

Private mDrawn As Boolean

Private WithEvents mEventProxy As ChartRegionEventProxy
Attribute mEventProxy.VB_VarHelpID = -1

Private mAutoscale As Boolean
Private mMinimumHeight As Single
Private mPrevHeight As Single
Private mPercentHeight As Double
Private mMinimumPercentHeight As Double
Private mYAxisPosition As Long
Private mIntegerYScale As Boolean

Private mRegionRect As TRectangle

Private mGridlineSpacingY As Double
Private mScaleGridSpacingY As Double
Private mYScaleFormatStr As String
Private mYScaleSubFormatStr As String
Private mGridTextHeight As Double

Private mBackColour As Long
Private mGridColour As Long
Private mShowGrid As Boolean
Private mShowCrosshairs As Boolean
Private mShowPerformanceText As Boolean

Private mPrevCursorX As Single
Private mPrevCursorY As Single

Private mTitle As String

Private mPerfText As text

Private mCurrentTool As ToolTypes

' defaults for bar series
Private mBarTailThickness As Long
Private mBarOutlineThickness As Long
Private mBarUpColour As Long
Private mBarDownColour As Long
Private mBarDisplayAsCandlestick As Boolean
Private mBarSolidUpBody As Boolean
Private mBarThickness As Long
Private mBarCandleWidth As Single

' defaults for datapoint series
Private mDataPointLineThickness As Long
Private mDataPointLineColour As Long
Private mDataPointLineStyle As LineStyles
Private mDataPointDisplayMode As DisplayModes
Private mDataPointHistBarWidth As Single
Private mDataPointIncludeInAutoscale As Boolean

' defaults for text and text series
Private mTextFont As StdFont
Private mTextColour As Long
Private mTextBox As Boolean
Private mTextBoxColour As Long
Private mTextBoxStyle As LineStyles
Private mTextBoxThickness As Long
Private mTextBoxFillColour As Long
Private mTextBoxFillStyle As FillStyles
Private mTextAlign As TextAlignModes
Private mTextKeepInView As Boolean
Private mTextFixedX As Boolean
Private mTextFixedY As Boolean
Private mTextIncludeInAutoscale As Boolean
Private mTextExtended As Boolean
Private mTextPaddingX As Double
Private mTextPaddingY As Double

' defaults for lines and line series
Private mLineColour As Long
Private mLineThickness As Long
Private mLineStyle As LineStyles
Private mLineExtendBefore As Boolean
Private mLineExtendAfter As Boolean
Private mLineArrowStartStyle As ArrowStyles
Private mLineArrowStartLength As Long
Private mLineArrowStartWidth As Long
Private mLineArrowStartColor As Long
Private mLineArrowStartFillColor As Long
Private mLineArrowStartFillStyle As FillStyles
Private mLineArrowEndStyle As ArrowStyles
Private mLineArrowEndLength As Long
Private mLineArrowEndWidth As Long
Private mLineArrowEndColor As Long
Private mLineArrowEndFillColor As Long
Private mLineArrowEndFillStyle As FillStyles
Private mLineFixedX As Boolean
Private mLineFixedY As Boolean
Private mLineIncludeInAutoscale As Boolean
Private mLineExtended As Boolean

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Dim aFont As StdFont
Dim i As Long

mNextLayer = LayerNumbers.LayerLowestUser

ReDim mPeriodTable(1000) As PeriodTableEntry
ReDim mGraphicObjects(10000) As GraphicObjectTableEntry

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

Set mBarSeriesCol = New Collection
Set mDataPointSeriesCol = New Collection
Set mLineSeriesCol = New Collection
Set mTextSeriesCol = New Collection
Set mTextObjects = New Collection
Set mLineObjects = New Collection
Set mEventProxy = New ChartRegionEventProxy

rectInitialise mRegionRect

mGridlineSpacingY = GridlineSpacingCm


' defaults for bar series
mBarTailThickness = 1
mBarOutlineThickness = 1
mBarUpColour = &H1D9311
mBarDownColour = &H43FC2
mBarDisplayAsCandlestick = True
mBarSolidUpBody = True
mBarThickness = 2
mBarCandleWidth = 0.5

' defaults for datapoint series
mDataPointLineThickness = 1
mDataPointLineColour = vbBlack
mDataPointLineStyle = LineStyles.LineSolid
mDataPointDisplayMode = DisplayModes.DisplayAsLines
mDataPointHistBarWidth = 0.5
mDataPointIncludeInAutoscale = True

' defaults for text and text series
Set aFont = New StdFont
aFont.Bold = False
aFont.Italic = False
aFont.name = "Arial"
aFont.Size = 10
aFont.Strikethrough = False
aFont.Underline = False
Set mTextFont = aFont
mTextColour = vbBlack
mTextBox = False
mTextBoxColour = vbBlack
mTextBoxStyle = LineStyles.LineSolid
mTextBoxThickness = 1
mTextBoxFillColour = vbWhite
mTextBoxFillStyle = FillStyles.FillSolid
mTextAlign = TextAlignModes.AlignBottomLeft
mTextKeepInView = False
mTextIncludeInAutoscale = False
mTextExtended = False
mTextPaddingX = 0.5
mTextPaddingY = 0.5

' defaults for lines and line series

mLineColour = vbBlack
mLineThickness = 1
mLineStyle = LineStyles.LineSolid
mLineExtendBefore = False
mLineExtendAfter = False
mLineArrowStartStyle = ArrowStyles.ArrowNone
mLineArrowStartLength = 20
mLineArrowStartWidth = 20
mLineArrowStartColor = vbBlack
mLineArrowStartFillColor = vbBlack
mLineArrowStartFillStyle = FillStyles.FillSolid
mLineArrowEndStyle = ArrowStyles.ArrowNone
mLineArrowEndLength = 10
mLineArrowEndWidth = 8
mLineArrowEndColor = vbBlack
mLineArrowEndFillColor = vbBlack
mLineArrowEndFillStyle = FillStyles.FillSolid
mLineFixedX = False
mLineFixedY = False
mLineIncludeInAutoscale = False
mLineExtended = False


End Sub

'================================================================================
' mEventProxy Event Handlers
'================================================================================

Private Sub mEventProxy_addGraphicObject(ByVal graphObj As IGraphicObject)
addGraphicObject graphObj
End Sub

Private Sub mEventProxy_objectChanged( _
                ByVal handle As Long, _
                ByRef areaToRedraw As TRectangle)
Dim layer As Long
Dim index As Long
Dim overlaps() As TRectangle
Dim redrawArea As TRectangle

redrawArea = rectIntersection(areaToRedraw, mRegionRect)

With mGraphicObjects(handle)
    
    ' get the new properties for the changed object
    .boundingRect = .graphObj.boundingRectangle(False, False, False, mGaugeX, mGaugeY, mRegionRect)
    
    If Not .boundingRect.isvalid Then
        removeEntryFromInScopeList handle
        removeEntryFromVisibleList handle
        Exit Sub
    End If
    
    .extendedObject = .graphObj.extendedObject
    .noDraw = .graphObj.noDraw
    If .extendedObject Then
        removeEntryFromPeriodList handle
        addEntryToExtendedList handle
    Else
        removeEntryFromExtendedList handle
        addEntryToPeriodList handle
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
    
        If redrawArea.isvalid Then
            Dim drawAreas(0) As TRectangle
            drawAreas(0) = redrawArea
            .graphObj.draw drawAreas
        End If
    End If
    
    ' now redraw any overlapping objects in higher layers than the changed object
    If Not mSuppressDrawing And Not .noDraw And redrawArea.isvalid Then
        For layer = .layer + 1 To MaxLayer
            index = mGraphicObjects(mVisibleTable(layer).firstIndex).nextVisibleIndex
            Do While index <> mVisibleTable(layer).lastindex
                With mGraphicObjects(index)
                    If index <> handle Then
                        If getObjectOverlaps(handle, index, redrawArea, True, True, overlaps) <> 0 Then
                            .graphObj.draw overlaps
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
    'If handle = 523 Then Stop
    If Not mSuppressDrawing And undrawnArea.isvalid Then
        For layer = MinLayer To MaxLayer
            index = mGraphicObjects(mVisibleTable(layer).firstIndex).nextVisibleIndex
            Do While index <> mVisibleTable(layer).lastindex
                With mGraphicObjects(index)
                    If index <> handle Then
                        If getObjectOverlaps(handle, index, undrawnArea, Not blockUndraw, True, overlaps) <> 0 Then
                            .graphObj.draw overlaps
                        End If
                    End If
                    index = .nextVisibleIndex
                End With
            Loop
        Next
    End If
End With
End Sub

Private Sub mEventProxy_boundsExceeded()
If mAutoscale And Not mSuppressDrawing Then
    calculateVerticalBounds
    paintRegion
End If
End Sub

Private Sub mEventProxy_paintRegion()
paintRegion
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get autoscale() As Boolean
autoscale = mAutoscale
End Property

Public Property Let autoscale(ByVal value As Boolean)
mAutoscale = value
If mAutoscale Then
    calculateVerticalBounds
    determineVisibleObjects True
    paintRegion
End If
End Property

Friend Property Get barCandleWidth() As Single
barCandleWidth = mBarCandleWidth
End Property

Friend Property Let barCandleWidth(ByVal value As Single)
mBarCandleWidth = value
End Property

Public Property Get barDisplayAsCandlestick() As Boolean
barDisplayAsCandlestick = mBarDisplayAsCandlestick
End Property

Public Property Let barDisplayAsCandlestick(ByVal value As Boolean)
mBarDisplayAsCandlestick = value
End Property

Public Property Get barDownColour() As Long
barDownColour = mBarDownColour
End Property

Public Property Let barDownColour(ByVal value As Long)
mBarDownColour = value
End Property

Public Property Get barOutlineThickness() As Long
barOutlineThickness = mBarOutlineThickness
End Property

Public Property Let barOutlineThickness(ByVal value As Long)
mBarOutlineThickness = value
End Property

Public Property Get barSolidUpBody() As Boolean
barSolidUpBody = mBarSolidUpBody
End Property

Public Property Let barSolidUpBody(ByVal value As Boolean)
mBarSolidUpBody = value
End Property

Public Property Get barTailThickness() As Long
barTailThickness = mBarTailThickness
End Property

Public Property Let barTailThickness(ByVal value As Long)
mBarTailThickness = value
End Property

Public Property Get barThickness() As Long
barThickness = mBarThickness
End Property

Public Property Let barThickness(ByVal value As Long)
mBarThickness = value
End Property

Public Property Get barUpColour() As Long
barUpColour = mBarUpColour
End Property

Public Property Let barUpColour(ByVal value As Long)
mBarUpColour = value
End Property

Public Property Get currentTool() As ToolTypes
currentTool = mCurrentTool
End Property

Friend Property Let currentTool(ByVal value As ToolTypes)
mCurrentTool = value
End Property

Public Property Get dataPointDisplayMode() As DisplayModes
dataPointDisplayMode = mDataPointDisplayMode
End Property

Public Property Let dataPointDisplayMode(ByVal value As DisplayModes)
mDataPointDisplayMode = value
End Property

Public Property Get dataPointHistBarWidth() As Single
dataPointHistBarWidth = mDataPointHistBarWidth
End Property

Public Property Let dataPointHistBarWidth(ByVal value As Single)
mDataPointHistBarWidth = value
End Property

Public Property Get dataPointIncludeInAutoscale() As Boolean
dataPointIncludeInAutoscale = mDataPointIncludeInAutoscale
End Property

Public Property Let dataPointIncludeInAutoscale(ByVal newValue As Boolean)
mDataPointIncludeInAutoscale = newValue
End Property

Public Property Get dataPointLineColour() As Long
dataPointLineColour = mDataPointLineColour
End Property

Public Property Let dataPointLineColour(ByVal value As Long)
mDataPointLineColour = value
End Property

Public Property Get dataPointLineStyle() As LineStyles
dataPointLineStyle = mDataPointLineStyle
End Property

Public Property Let dataPointLineStyle(ByVal value As LineStyles)
mDataPointLineStyle = value
End Property

Public Property Get dataPointLineThickness() As Long
dataPointLineThickness = mDataPointLineThickness
End Property

Public Property Let dataPointLineThickness(ByVal value As Long)
mDataPointLineThickness = value
End Property

Public Property Get gridColor() As Long
gridColor = mGridColour
End Property

Public Property Let gridColor(ByVal val As Long)
mGridColour = val
End Property

Public Property Get gridlineSpacingY() As Double
gridlineSpacingY = mGridlineSpacingY
End Property

Public Property Let gridlineSpacingY(ByVal value As Double)
mGridlineSpacingY = value
End Property

Public Property Let integerYScale(ByVal value As Boolean)
mIntegerYScale = value
End Property

Public Property Get lineArrowEndColor() As Long
lineArrowEndColor = mLineArrowEndColor
End Property

Public Property Let lineArrowEndColor(ByVal newValue As Long)
mLineArrowEndColor = newValue
End Property

Public Property Get lineArrowEndFillColor() As Long
lineArrowEndFillColor = mLineArrowEndFillColor
End Property

Public Property Let lineArrowEndFillColor(ByVal newValue As Long)
mLineArrowEndFillColor = newValue
End Property

Public Property Get lineArrowEndFillStyle() As FillStyles
lineArrowEndFillStyle = mLineArrowEndFillStyle
End Property

Public Property Let lineArrowEndFillStyle(ByVal newValue As FillStyles)
mLineArrowEndFillStyle = newValue
End Property

Public Property Get lineArrowEndLength() As Long
lineArrowEndLength = mLineArrowEndLength
End Property

Public Property Let lineArrowEndLength(ByVal newValue As Long)
mLineArrowEndLength = newValue
End Property

Public Property Get lineArrowEndStyle() As ArrowStyles
lineArrowEndStyle = mLineArrowEndStyle
End Property

Public Property Let lineArrowEndStyle(ByVal newValue As ArrowStyles)
mLineArrowEndStyle = newValue
End Property

Public Property Get lineArrowEndWidth() As Long
lineArrowEndWidth = mLineArrowEndWidth
End Property

Public Property Let lineArrowEndWidth(ByVal newValue As Long)
mLineArrowEndWidth = newValue
End Property

Public Property Get lineArrowStartColor() As Long
lineArrowStartColor = mLineArrowStartColor
End Property

Public Property Let lineArrowStartColor(ByVal newValue As Long)
mLineArrowStartColor = newValue
End Property

Public Property Get lineArrowStartFillColor() As Long
lineArrowStartFillColor = mLineArrowStartFillColor
End Property

Public Property Let lineArrowStartFillColor(ByVal newValue As Long)
mLineArrowStartFillColor = newValue
End Property

Public Property Get lineArrowStartFillStyle() As FillStyles
lineArrowStartFillStyle = lineArrowStartFillStyle
End Property

Public Property Let lineArrowStartFillStyle(ByVal newValue As FillStyles)
mLineArrowStartFillStyle = newValue
End Property

Public Property Get lineArrowStartLength() As Long
lineArrowStartLength = mLineArrowStartLength
End Property

Public Property Let lineArrowStartLength(ByVal newValue As Long)
mLineArrowStartLength = newValue
End Property

Public Property Get lineArrowStartStyle() As ArrowStyles
lineArrowStartStyle = mLineArrowStartStyle
End Property

Public Property Let lineArrowStartStyle(ByVal newValue As ArrowStyles)
mLineArrowStartStyle = newValue
End Property

Public Property Get lineArrowStartWidth() As Long
lineArrowStartWidth = mLineArrowStartWidth
End Property

Public Property Let lineArrowStartWidth(ByVal newValue As Long)
mLineArrowStartWidth = newValue
End Property

Public Property Get lineColor() As Long
lineColor = mLineColour
End Property

Public Property Let lineColor(ByVal val As Long)
mLineColour = val
End Property

Public Property Get lineExtendAfter() As Boolean
lineExtendAfter = mLineExtendAfter
End Property

Public Property Let lineExtendAfter(ByVal val As Boolean)
mLineExtendAfter = val
End Property

Public Property Get lineExtendBefore() As Boolean
lineExtendBefore = mLineExtendBefore
End Property

Public Property Let lineExtendBefore(ByVal val As Boolean)
mLineExtendBefore = val
End Property

Public Property Get lineExtended() As Boolean
lineExtended = mLineExtended
End Property

Public Property Let lineExtended(ByVal val As Boolean)
mLineExtended = val
End Property

Public Property Get lineFixedX() As Boolean
lineFixedX = mLineFixedX
End Property

Public Property Let lineFixedX(ByVal val As Boolean)
mLineFixedX = val
End Property

Public Property Get lineFixedY() As Boolean
lineFixedY = mLineFixedY
End Property

Public Property Let lineFixedY(ByVal val As Boolean)
mLineFixedY = val
End Property

Public Property Get lineIncludeInAutoscale() As Boolean
lineIncludeInAutoscale = mLineIncludeInAutoscale
End Property

Public Property Let lineIncludeInAutoscale(ByVal val As Boolean)
mLineIncludeInAutoscale = val
End Property

Public Property Get lineStyle() As LineStyles
lineStyle = mLineStyle
End Property

Public Property Let lineStyle(ByVal val As LineStyles)
mLineStyle = val
End Property

Public Property Get lineThickness() As Long
lineThickness = mLineThickness
End Property

Public Property Let lineThickness(ByVal val As Long)
mLineThickness = val
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

Public Property Let minimumPercentHeight(ByVal value As Double)
mMinimumPercentHeight = value
End Property

Public Property Get percentheight() As Double
percentheight = mPercentHeight
End Property

Public Property Let percentheight(ByVal value As Double)
mPercentHeight = value
End Property

Public Property Get regionBackColor() As Long
regionBackColor = mBackColour
End Property

Public Property Let regionBackColor(ByVal val As Long)
mBackColour = val
End Property

Public Property Get regionBottom() As Single
regionBottom = mSurface.ScaleTop + mSurface.ScaleHeight
End Property

Friend Property Let regionBottom(ByVal value As Single)
mSurface.ScaleHeight = value - mSurface.ScaleTop
mRegionRect.bottom = value
End Property

Public Property Get regionHeight() As Single
regionHeight = -mSurface.ScaleHeight
End Property

Friend Property Let regionHeight(ByVal value As Single)
With mSurface
    If value < mMinimumHeight Then
        .ScaleHeight = -minimumHeight
    Else
        .ScaleHeight = -value
    End If
    If mGaugeY = 0 Then
        ' this is the first time regionHeight has been set, so we need to
        ' set Y gauge for use when graphic objects are added before the
        ' first call to paint region
        mGaugeY = calculateGaugeY
    End If
    mRegionRect.bottom = .ScaleTop + .ScaleHeight
End With
End Property

Public Property Get regionLeft() As Single
regionLeft = mSurface.ScaleLeft
End Property

Friend Property Let regionLeft(ByVal value As Single)
mSurface.ScaleLeft = value
mRegionRect.left = value
End Property

Public Property Get regionNumber() As Long
regionNumber = mRegionNumber
End Property

Public Property Let regionNumber(ByVal val As Long)
mRegionNumber = val
End Property

Public Property Get regionTop() As Single
regionTop = mSurface.ScaleTop
End Property

Friend Property Let regionTop(ByVal value As Single)
mRegionRect.bottom = mRegionRect.bottom + value - mRegionRect.top
mSurface.ScaleTop = value
mRegionRect.top = value
End Property

Public Property Get regionWidth() As Single
regionWidth = mSurface.ScaleWidth
End Property

Friend Property Let regionWidth(ByVal value As Single)
mSurface.ScaleWidth = value
End Property

Public Property Get showGrid() As Boolean
showGrid = mShowGrid
End Property

Public Property Let showGrid(ByVal val As Boolean)
mShowGrid = val
End Property

Friend Property Get showCrosshairs() As Boolean
showCrosshairs = mShowCrosshairs
End Property

Friend Property Let showCrosshairs(ByVal val As Boolean)
If mShowCrosshairs = val Then Exit Property
mShowCrosshairs = val
If Not mDrawn Then Exit Property
If mShowCrosshairs Then
    undrawPointer
    mSurface.MousePointer = vbDefault
Else
    undrawCrosshairs
    mSurface.MousePointer = vbCrosshair
End If

End Property

Public Property Get showPerformanceText() As Boolean
showPerformanceText = mShowPerformanceText
End Property

Public Property Let showPerformanceText(ByVal val As Boolean)
mShowPerformanceText = val
End Property

Friend Property Let suppressDrawing(ByVal val As Boolean)
If mSuppressDrawing = val Then Exit Property
mSuppressDrawing = val
If Not mSuppressDrawing Then
    determineInScopeObjects
    If mAutoscale Then
        calculateVerticalBounds
    End If
    determineVisibleObjects True
    paintRegion
End If
End Property

Friend Property Let surface(ByVal value As PictureBox)
Set mSurface = value
With mSurface
    mPrevHeight = .Height
End With
End Property

Public Property Get textAlign() As TextAlignModes
textAlign = mTextAlign
End Property

Public Property Let textAlign(ByVal newValue As TextAlignModes)
mTextAlign = newValue
End Property

Friend Property Get textBox() As Boolean
textBox = mTextBox
End Property

Public Property Let textBox(ByVal newValue As Boolean)
mTextBox = newValue
End Property

Public Property Get textBoxColor() As Long
textBoxColor = mTextBoxColour
End Property

Public Property Let textBoxColor(ByVal newValue As Long)
mTextBoxColour = newValue
End Property

Friend Property Get textBoxFillColor() As Long
textBoxFillColor = mTextBoxFillColour
End Property

Public Property Let textBoxFillColor(ByVal newValue As Long)
mTextBoxFillColour = newValue
End Property

Public Property Get textBoxFillStyle() As FillStyles
textBoxFillStyle = mTextBoxFillStyle
End Property

Public Property Let textBoxFillStyle(ByVal newValue As FillStyles)
mTextBoxFillStyle = newValue
End Property

Public Property Get textBoxStyle() As LineStyles
textBoxStyle = mTextBoxStyle
End Property

Public Property Let textBoxStyle(ByVal newValue As LineStyles)
mTextBoxStyle = newValue
End Property

Public Property Get textBoxThickness() As Long
textBoxThickness = mTextBoxThickness
End Property

Public Property Let textBoxThickness(ByVal newValue As Long)
mTextBoxThickness = newValue
End Property

Public Property Get textColor() As Long
textColor = mTextColour
End Property

Public Property Let textColor(ByVal newValue As Long)
mTextColour = newValue
End Property

Friend Property Get textExtended() As Boolean
textExtended = mTextExtended
End Property

Public Property Let textExtended(ByVal newValue As Boolean)
mTextExtended = newValue
End Property

Friend Property Get textFixedX() As Boolean
textFixedX = mTextFixedX
End Property

Public Property Let textFixedX(ByVal newValue As Boolean)
mTextFixedX = newValue
End Property

Friend Property Get textFixedY() As Boolean
textFixedY = mTextFixedY
End Property

Public Property Let textFixedY(ByVal newValue As Boolean)
mTextFixedY = newValue
End Property

Public Property Get textFont() As StdFont
Set textFont = mTextFont
End Property

Public Property Let textFont(ByVal value As StdFont)
Set mTextFont = value
End Property

Public Property Get textIncludeInAutoscale() As Boolean
textIncludeInAutoscale = mTextIncludeInAutoscale
End Property

Public Property Let textIncludeInAutoscale(ByVal newValue As Boolean)
mTextIncludeInAutoscale = newValue
End Property

Public Property Get textKeepInView() As Boolean
textKeepInView = mTextKeepInView
End Property

Public Property Let textKeepInView(ByVal newValue As Boolean)
mTextKeepInView = newValue
End Property

Public Property Get textPaddingX() As Double
textPaddingX = mTextPaddingX
End Property

Public Property Let textPaddingX(ByVal newValue As Double)
mTextPaddingX = newValue
End Property

Public Property Get textPaddingY() As Double
textPaddingY = mTextPaddingY
End Property

Public Property Let textPaddingY(ByVal newValue As Double)
mTextPaddingY = newValue
End Property

Public Property Get scaleGridSpacingY() As Single
scaleGridSpacingY = mScaleGridSpacingY
End Property

Public Property Get YAxisPosition() As Long
YAxisPosition = mYAxisPosition
End Property

'================================================================================
' Methods
'================================================================================

Public Function addBarSeries(Optional ByVal layer As Long = -1) As BarSeries
Set addBarSeries = New BarSeries
mBarSeriesCol.Add addBarSeries
addBarSeries.surface = mSurface
addBarSeries.candleWidth = barCandleWidth
addBarSeries.displayAsCandlestick = barDisplayAsCandlestick
addBarSeries.downColour = barDownColour
addBarSeries.eventProxy = mEventProxy
addBarSeries.layer = getLayer(layer)
addBarSeries.outlineThickness = barOutlineThickness
addBarSeries.solidUpBody = barSolidUpBody
addBarSeries.tailThickness = barTailThickness
addBarSeries.barThickness = barThickness
addBarSeries.upColour = barUpColour
addBarSeries.eventProxy = mEventProxy
End Function

Public Function addDataPointSeries(Optional ByVal layer As Long = -1) As DataPointSeries
Set addDataPointSeries = New DataPointSeries
mDataPointSeriesCol.Add addDataPointSeries
addDataPointSeries.surface = mSurface
addDataPointSeries.eventProxy = mEventProxy
addDataPointSeries.displayMode = dataPointDisplayMode
addDataPointSeries.histBarWidth = dataPointHistBarWidth
addDataPointSeries.includeInAutoscale = dataPointIncludeInAutoscale
addDataPointSeries.layer = getLayer(layer)
addDataPointSeries.lineColour = dataPointLineColour
addDataPointSeries.lineStyle = dataPointLineStyle
addDataPointSeries.lineThickness = dataPointLineThickness
End Function

Public Function addLine(Optional ByVal layer As Long = -1) As ChartSkil.Line
Set addLine = New ChartSkil.Line
mLineObjects.Add addLine
With addLine
    .eventProxy = mEventProxy
    .surface = mSurface
    .arrowEndColor = lineArrowEndColor
    .arrowEndFillColor = lineArrowEndFillColor
    .arrowEndFillStyle = lineArrowEndFillStyle
    .arrowEndLength = lineArrowEndLength
    .arrowEndStyle = lineArrowEndStyle
    .arrowEndWidth = lineArrowEndWidth
    .arrowStartColor = lineArrowStartColor
    .arrowStartFillColor = lineArrowStartFillColor
    .arrowStartFillStyle = lineArrowStartFillStyle
    .arrowStartLength = lineArrowStartLength
    .arrowStartStyle = lineArrowStartStyle
    .arrowStartWidth = lineArrowStartWidth
    .color = lineColor
    .extendAfter = lineExtendAfter
    .extendBefore = lineExtendBefore
    .extended = lineExtended
    .fixedX = lineFixedX
    .fixedY = lineFixedY
    .style = lineStyle
    .thickness = lineThickness
    .includeInAutoscale = lineIncludeInAutoscale
    .layer = getLayer(layer)
End With
addGraphicObject addLine
End Function

Public Function addLineSeries(Optional ByVal layer As Long = -1) As LineSeries
Set addLineSeries = New LineSeries
mLineSeriesCol.Add addLineSeries
With addLineSeries
    .surface = mSurface
    .eventProxy = mEventProxy
    .arrowEndColor = lineArrowEndColor
    .arrowEndFillColor = lineArrowEndFillColor
    .arrowEndFillStyle = lineArrowEndFillStyle
    .arrowEndLength = lineArrowEndLength
    .arrowEndStyle = lineArrowEndStyle
    .arrowEndWidth = lineArrowEndWidth
    .arrowStartColor = lineArrowStartColor
    .arrowStartFillColor = lineArrowStartFillColor
    .arrowStartFillStyle = lineArrowStartFillStyle
    .arrowStartLength = lineArrowStartLength
    .arrowStartStyle = lineArrowStartStyle
    .arrowStartWidth = lineArrowStartWidth
    .color = lineColor
    .extendAfter = lineExtendAfter
    .extendBefore = lineExtendBefore
    .extended = lineExtended
    .fixedX = lineFixedX
    .fixedY = lineFixedY
    .includeInAutoscale = lineIncludeInAutoscale
    .layer = getLayer(layer)
    .style = lineStyle
    .thickness = lineThickness
End With
End Function

Friend Sub addPeriod(ByVal periodNumber As Long)
If periodNumber > UBound(mPeriodTable) Then
    ReDim Preserve mPeriodTable(UBound(mPeriodTable) + 1000) As PeriodTableEntry
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

Public Function addText(Optional ByVal layer As Long = -1) As text
Set addText = New text
mTextObjects.Add addText
addText.eventProxy = mEventProxy
addText.surface = mSurface
addText.align = textAlign
addText.box = textBox
addText.boxColor = textBoxColor
addText.boxFillColor = textBoxFillColor
addText.boxFillStyle = textBoxFillStyle
addText.boxStyle = textBoxStyle
addText.boxThickness = textBoxThickness
addText.color = textColor
addText.extended = textExtended
addText.fixedX = textFixedX
addText.fixedY = textFixedY
addText.font = textFont
addText.includeInAutoscale = textIncludeInAutoscale
addText.keepInView = textKeepInView
addText.paddingX = textPaddingX
addText.paddingY = textPaddingY
addText.layer = getLayer(layer)
addGraphicObject addText
End Function

Public Function addTextSeries(Optional ByVal layer As Long = -1) As TextSeries
Set addTextSeries = New TextSeries
mTextSeriesCol.Add addTextSeries
addTextSeries.surface = mSurface
addTextSeries.eventProxy = mEventProxy
addTextSeries.align = textAlign
addTextSeries.box = textBox
addTextSeries.boxColor = textBoxColor
addTextSeries.boxFillColor = textBoxFillColor
addTextSeries.boxFillStyle = textBoxFillStyle
addTextSeries.boxStyle = textBoxStyle
addTextSeries.boxThickness = textBoxThickness
addTextSeries.color = textColor
addTextSeries.extended = textExtended
addTextSeries.fixedX = textFixedX
addTextSeries.fixedY = textFixedY
addTextSeries.font = textFont
addTextSeries.includeInAutoscale = textIncludeInAutoscale
addTextSeries.layer = getLayer(layer)
addTextSeries.keepInView = textKeepInView
addTextSeries.paddingX = textPaddingX
addTextSeries.paddingY = textPaddingY
End Function

Friend Sub MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim yVal As Single
If Y = MinusInfinitySingle Then
    yVal = mSurface.ScaleTop + 2 * mSurface.ScaleHeight ' make sure the pointer isn't visible
Else
    yVal = Y
End If
If Round(X) >= YAxisPosition Then Exit Sub
If mShowCrosshairs Then
    undrawCrosshairs
    drawCrosshairs X, yVal
Else
    undrawPointer
    drawPointer X, yVal
End If
End Sub

Public Function newDimension(ByVal X As Double, _
                        ByVal Y As Double) As Dimension
Set newDimension = New Dimension
newDimension.surface = mSurface
newDimension.X = X
newDimension.Y = Y
End Function

Public Function newPoint(ByVal X As Double, _
                        ByVal Y As Double, _
                        Optional ByVal posnTypeX As PositionType = PositionAbsolute, _
                        Optional ByVal posnTypeY As PositionType = PositionAbsolute) As Point
Set newPoint = New Point
newPoint.surface = mSurface
newPoint.X = X
newPoint.Y = Y
newPoint.PositionTypeX = posnTypeX
newPoint.PositionTypeY = posnTypeY
End Function

Friend Sub paintRegion()
Dim startTimer As Single
Dim heightCm As Single
Dim numGridDivisions As Double
Dim gridSpacing As Double
Dim i As Long
Dim z As Double
Dim perfFont As StdFont
Dim counter As Long
Dim index As Long
Dim gaugeX As Double
Dim gaugeY As Double

If mSuppressDrawing Then Exit Sub

Static repaintCount As Long
repaintCount = repaintCount + 1
'debug.print "repaint count = " & repaintCount

With mSurface
startTimer = Timer

.Cls
.backColor = regionBackColor

heightCm = (mSurface.Height / 1440#) * 2.54

If .Height <> mPrevHeight Then
    ' the region has been resized vertically - need to recalculate the position
    ' of any in-scope gauge-dependent objects
    
    ' Note that resizing will have changed scaleheight and scale top, so we first
    ' need to reset these to their proper values if we're autoscaling
    
    .ScaleTop = mRegionRect.top
    .ScaleHeight = mRegionRect.bottom - mRegionRect.top

    mPrevHeight = .Height
End If

' gauge may also have changed if the region has extended through auto- or
' manual scaling

gaugeX = calculateGaugeX
gaugeY = calculateGaugeY

If gaugeX <> mGaugeX Or gaugeY <> mGaugeY Then
    mGaugeX = gaugeX
    mGaugeY = gaugeY
    index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
    Do While index <> mLastInScopeIndex
         With mGraphicObjects(index)
            If .gaugeDependent Then
                .boundingRect = .graphObj.boundingRectangle(False, False, True, mGaugeX, mGaugeY, mRegionRect)
                If graphicObjectIsVisible(index) Then
                    addEntryToVisibleList index
                End If
            End If
            index = .nextInScopeIndex
        End With
    Loop
End If

numGridDivisions = heightCm / mGridlineSpacingY
gridSpacing = -.ScaleHeight / numGridDivisions

mGridTextHeight = .TextHeight("123")

Dim log10 As Double
Dim logSpacing As Double
Dim exp As Long
Dim mantissa As Double

log10 = Log(10#)
logSpacing = Log(gridSpacing) / log10

exp = Fix(logSpacing)
mantissa = Abs(logSpacing) - Abs(exp)

If logSpacing < 0 Then
    exp = exp - 1
    mantissa = 1 - mantissa
End If

If mantissa >= Log(5) / log10 Then
    mScaleGridSpacingY = 5 * 10 ^ exp
ElseIf mantissa >= Log(2.5) / log10 Then
    mScaleGridSpacingY = 2.5 * 10 ^ exp
ElseIf mantissa >= Log(2#) / log10 Then
    mScaleGridSpacingY = 2# * 10 ^ exp
Else
    mScaleGridSpacingY = 10 ^ exp
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

.DrawWidth = 1
.DrawMode = vbCopyPen
z = Int((.ScaleTop + .ScaleHeight) / mScaleGridSpacingY) * mScaleGridSpacingY
Do While z < .ScaleTop
    mSurface.Line (.ScaleLeft, z)-(YAxisPosition, z), gridColor
    .CurrentX = YAxisPosition + 0.5
    .CurrentY = .CurrentY - mGridTextHeight / 2
    .ForeColor = vbWhite Xor .backColor
    mSurface.Print Format(z, mYScaleFormatStr)
    z = z + mScaleGridSpacingY
Loop

For i = MinLayer To MaxLayer
    index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(i).lastindex
        With mGraphicObjects(index)
            Dim drawAreas(0) As TRectangle
            drawAreas(0) = rectIntersection(.boundingRect, mRegionRect)
            .graphObj.draw drawAreas
            index = .nextVisibleIndex
            counter = counter + 1
        End With
    Loop
Next

If mShowCrosshairs Then
    drawCrosshairs mPrevCursorX, mPrevCursorY
Else
    drawPointer mPrevCursorX, mPrevCursorY
End If

If mShowPerformanceText Then
    If mPerfText Is Nothing Then
        Set perfFont = New StdFont
        perfFont.name = "Lucida Console"
        perfFont.Size = 8
        perfFont.Underline = False
        perfFont.Bold = False
        Set mPerfText = addText()
        mPerfText.color = vbActiveTitleBarText
        mPerfText.font = perfFont
        mPerfText.box = True
        mPerfText.boxColor = vbBlack
        mPerfText.boxStyle = LineStyles.LineInsideSolid
        mPerfText.boxThickness = 1
        mPerfText.boxFillColor = vbActiveTitleBar
        mPerfText.boxFillStyle = FillStyles.FillSolid
        mPerfText.position = newPoint(5, 2, PositionRelative, PositionRelative)
        mPerfText.fixedX = True
        mPerfText.fixedY = True
        mPerfText.align = TextAlignModes.AlignBottomLeft
        mPerfText.includeInAutoscale = False
        mPerfText.keepInView = True
        mPerfText.paddingX = 0.5
        mPerfText.layer = LayerTitle
        
    End If
    mPerfText.text = Format(1000 * (Timer - startTimer), "0.0") & "ms" & vbCrLf & _
                    counter & " visible objects" & vbCrLf & _
                    mGraphicObjectsCount & " total objects"
    ' need to force the text to draw now
    mPerfText.draw
End If

End With

mDrawn = True
End Sub

Friend Sub periodsInView(ByVal minValue As Single, ByVal maxValue As Single)
Dim gaugeChanged As Boolean
Dim gaugeX As Double
Dim gaugeY As Double

mMinInViewPeriod = minValue
mMaxInViewPeriod = maxValue
mSurface.ScaleLeft = minValue
mYAxisPosition = maxValue + 1
mRegionRect.left = minValue
mRegionRect.right = maxValue '+ 0.5
mRegionRect.isvalid = True

gaugeX = calculateGaugeX
If gaugeX <> mGaugeX Then
    gaugeChanged = True
    mGaugeX = gaugeX
End If
If Not mSuppressDrawing Then
    determineInScopeObjects
    
    If mAutoscale Then
        calculateVerticalBounds
        gaugeY = calculateGaugeY
        If gaugeY <> mGaugeY Then
            gaugeChanged = True
            mGaugeY = gaugeY
        End If
    End If
    determineVisibleObjects gaugeChanged
End If
End Sub

Public Sub setTitle(ByVal val As String, _
                        ByVal color As Long, _
                        ByVal font As StdFont)
Dim title As text
mTitle = val
Set title = addText(LayerTitle)
title.color = color
title.font = font
title.box = True
title.boxColor = vbBlack
title.boxStyle = LineStyles.LineInvisible
title.boxThickness = 1
title.boxFillColor = vbWhite
title.boxFillStyle = FillStyles.FillSolid
title.position = newPoint(0.2, 0.2, PositionDistance, PositionCounterDistance)
title.fixedX = True
title.fixedY = True
title.align = TextAlignModes.AlignTopLeft
title.includeInAutoscale = False
title.keepInView = True
title.paddingX = 0.5
title.text = mTitle
End Sub

Public Sub setVerticalScale(ByVal lowValue As Single, _
                            ByVal highValue As Single)
Dim gaugeChanged As Boolean
Dim gaugeY As Double

If mAutoscale Then Exit Sub
regionTop = highValue
regionHeight = highValue - lowValue
gaugeY = calculateGaugeY
If gaugeY <> mGaugeY Then
    gaugeChanged = True
    mGaugeY = gaugeY
End If
If Not mSuppressDrawing Then
    determineVisibleObjects gaugeChanged
    paintRegion
End If
End Sub

Public Sub scrollVertical(ByVal amount As Single)
mAutoscale = False
regionTop = regionTop + amount
If Not mSuppressDrawing Then
    determineVisibleObjects False
    paintRegion
End If
End Sub

Public Sub scrollVerticalProportion(ByVal proportion As Single)
mAutoscale = False
regionTop = regionTop - proportion * mSurface.ScaleHeight
If Not mSuppressDrawing Then
    determineVisibleObjects False
    paintRegion
End If
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

With mGraphicObjects(index)
    newPeriodNumber = .graphObj.periodNumber
    If newPeriodNumber <> .periodNumber Then
        If .periodNumber >= 1 Then
            removeEntryFromPeriodList index
        End If
        If newPeriodNumber < 1 Then Exit Sub
        addPeriod newPeriodNumber  ' in case the period hasn't been added yet
        lastindex = mPeriodTable(newPeriodNumber).lastindex
        .periodNumber = newPeriodNumber
        .nextPeriodIndex = lastindex
        .prevPeriodIndex = mGraphicObjects(lastindex).prevPeriodIndex
        mGraphicObjects(.prevPeriodIndex).nextPeriodIndex = index
        mGraphicObjects(lastindex).prevPeriodIndex = index
    End If
End With

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

Private Function addGraphicObject(ByVal graphObj As IGraphicObject) As Long
Dim handle As Long

mGraphicObjectsCount = mGraphicObjectsCount + 1

handle = allocateGraphicObjectTableEntry

graphObj.handle = handle

With mGraphicObjects(handle)
    Set .graphObj = graphObj
    .boundingRect = graphObj.boundingRectangle(True, True, True, mGaugeX, mGaugeY, mRegionRect)
    .extendedObject = graphObj.extendedObject
    .gaugeDependent = graphObj.gaugeDependent
    .includeInAutoscale = graphObj.includeInAutoscale
    .layer = graphObj.layer
    .scaleDependent = graphObj.scaleDependent
    .capabilities = graphObj.capabilities
    .noDraw = graphObj.noDraw
    
    If Not .extendedObject Then
        addEntryToPeriodList handle
    Else
        addEntryToExtendedList handle
    End If

End With

If graphicObjectIsInScope(handle) Then
    addEntryToInScopeList handle
    If graphicObjectIsVisible(handle) Then
        addEntryToVisibleList handle
    End If
End If
addGraphicObject = handle
End Function

Private Function allocateGraphicObjectTableEntry() As Long
Static graphicObjectsIndex As Long
allocateGraphicObjectTableEntry = graphicObjectsIndex
graphicObjectsIndex = graphicObjectsIndex + 1

If graphicObjectsIndex = UBound(mGraphicObjects) Then
    ReDim Preserve mGraphicObjects(UBound(mGraphicObjects) + 10000) As GraphicObjectTableEntry
End If

End Function

Private Function calculateGaugeX() As Double
calculateGaugeX = (mSurface.ScaleWidth / mSurface.Width) * Screen.TwipsPerPixelX
End Function

Private Function calculateGaugeY() As Double
calculateGaugeY = -(mSurface.ScaleHeight / mSurface.Height) * Screen.TwipsPerPixelY
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

Private Sub calculateVerticalBounds()
Dim objMax As Double
Dim objMin As Double
Dim scaleMax As Double
Dim scaleMin As Double
Dim index As Long
Dim newGaugeX As Double
Dim newGaugeY As Double
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
                                                rectGetXInterval(mRegionRect), _
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

With mSurface
    If scaleMin > scaleMax Then
        ' no autoscale relevant in-view graphic objects -
        ' so leave default scale
    ElseIf mMinimumHeight <> 0 Then
        If scaleMax - scaleMin < mMinimumHeight Then
            .ScaleHeight = -mMinimumHeight
            .ScaleTop = (scaleMax + scaleMin) / 2 + mMinimumHeight / 2
        Else
            .ScaleHeight = (scaleMin - scaleMax) * 1.05
            .ScaleTop = scaleMax + (scaleMax - scaleMin) * 0.025
        End If
    Else
        If scaleMin = scaleMax Then
            scaleMax = scaleMax * 1.001
            scaleMin = scaleMax * 0.998
        End If
        .ScaleHeight = (scaleMin - scaleMax) * 1.05
        .ScaleTop = scaleMax + (scaleMax - scaleMin) * 0.025
    End If
    mRegionRect.top = .ScaleTop
    mRegionRect.bottom = .ScaleTop + .ScaleHeight
    mRegionRect.isvalid = True
End With

If Not someGaugeDependentObjects Then Exit Sub

' there are gauge-dependent objects that may affect the max and min scales.
' We need to recalculate their bounds with the new gauge determined from
' the other objects to get a better estimate of their sizes.
newGaugeX = calculateGaugeX
newGaugeY = calculateGaugeY
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
                                        newGaugeX, _
                                        newGaugeY, _
                                        mRegionRect)
            calculateObjectHighAndLowInInterval index, _
                                            rectGetXInterval(mRegionRect), _
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

With mSurface
    .ScaleHeight = (scaleMin - scaleMax) * 1.05
    .ScaleTop = scaleMax + (scaleMax - scaleMin) * 0.025
    mRegionRect.top = .ScaleTop
    mRegionRect.bottom = .ScaleTop + .ScaleHeight
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

timerIn = Timer

clearInScopeObjectList

For i = IIf(mMinInViewPeriod < 1, 1, mMinInViewPeriod) To mMaxInViewPeriod
    index = mGraphicObjects(mPeriodTable(i).firstIndex).nextPeriodIndex
    Do While index <> 0
        numberExamined = numberExamined + 1
        With mGraphicObjects(index)
            If .boundingRect.isvalid Then
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
        .boundingRect = .graphObj.boundingRectangle(True, False, False, mGaugeX, mGaugeY, mRegionRect)
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

Private Sub determineVisibleObjects(ByVal gaugeChanged As Boolean)
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
                                    gaugeChanged, _
                                    mGaugeX, _
                                    mGaugeY, _
                                    mRegionRect)
        If graphicObjectIsVisible(index) Then
            addEntryToVisibleList index
            numberVisible = numberVisible + 1
        End If
        index = .nextInScopeIndex
    End With
Loop

Debug.Print "DetermineVisibleObjects: total=" & numberExamined & " visible=" & numberVisible & " time=" & 1000 * (Timer - timerIn) & "ms"
End Sub

Private Sub drawCrosshairs(ByVal X As Single, ByVal Y As Single)
Dim lineX As Single

If mSuppressDrawing Then Exit Sub

With mSurface
    .DrawWidth = 1
    .DrawStyle = LineStyles.LineSolid
    .DrawMode = vbXorPen
    mSurface.Line (.ScaleLeft, Y)-(YAxisPosition, Y), vbRed Xor .backColor
    lineX = Round(X)
'    If lineX < 1 Then lineX = 1
    mSurface.Line (lineX, .ScaleTop + .ScaleHeight)-(lineX, .ScaleTop), vbRed Xor .backColor
    .CurrentX = YAxisPosition + 0.5
    .CurrentY = Y - mGridTextHeight / 2
    .ForeColor = vbRed 'Xor BackColor 'vbWhite 'Xor BackColor
    .FontTransparent = False
    mSurface.Print Format(Y, mYScaleSubFormatStr)
    mPrevCursorX = X
    mPrevCursorY = Y
    
End With

End Sub

Private Sub drawPointer(ByVal X As Single, ByVal Y As Single)
Dim lineX As Single

If mSuppressDrawing Then Exit Sub

With mSurface
    .MousePointer = vbCrosshair
    .DrawWidth = 1
    .DrawStyle = LineStyles.LineSolid
    .DrawMode = vbXorPen
    lineX = Round(X)
'    if lineX < 1 Then lineX = 1
    .FillColor = vbYellow Xor .backColor
    .FillStyle = vbSolid
    mSurface.Circle (lineX, Y), newPoint(0.8, 0, PositionDistance, PositionDistance).XAbsolute - mSurface.ScaleLeft, vbYellow Xor .backColor
    ''debug.print "DrawPointer x=" & lineX & " y=" & y
    .CurrentX = YAxisPosition + 0.5
    .CurrentY = Y - mGridTextHeight / 2
    .ForeColor = vbRed 'Xor BackColor 'vbWhite 'Xor BackColor
    .FontTransparent = False
    mSurface.Print Format(Y, mYScaleSubFormatStr)
    mPrevCursorX = X
    mPrevCursorY = Y
    
End With

End Sub

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
If Not overlaps(0).isvalid Then
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
                If boundingRects1(j).isvalid Then
                    For k = 0 To UBound(boundingRects2)
                        If boundingRects2(k).isvalid Then
                            If getObjectOverlaps > UBound(overlaps) Then
                                ReDim Preserve overlaps(UBound(overlaps) + 5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
                            End If
                            overlaps(getObjectOverlaps) = _
                                rectIntersection(boundingRects1(j), boundingRects2(k))
                            If overlaps(getObjectOverlaps).isvalid Then
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
                If overlaps(getObjectOverlaps).isvalid Then
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
                If overlaps(getObjectOverlaps).isvalid Then
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
    If Not .boundingRect.isvalid Then Exit Function
    If .extendedObject Then
        graphicObjectIsInScope = intOverlaps(rectGetXInterval(mRegionRect), rectGetXInterval(.boundingRect))
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
    If Not .boundingRect.isvalid Or .noDraw Then Exit Function
    If .extendedObject And (.capabilities And capabilities.ComplexBounding) Then
        overlapXInterval = intIntersection(rectGetXInterval(mRegionRect), rectGetXInterval(.boundingRect))
        For i = Int(overlapXInterval.startValue + 0.99999) To _
                Int(overlapXInterval.endValue)
            boundingRects = .graphObj.boundingRectanglesAt(i)
            numrects = -1
            On Error Resume Next
            numrects = UBound(boundingRects)
            On Error GoTo 0
            If numrects >= 0 Then
                For j = 0 To numrects
                    If rectOverlaps(boundingRects(j), mRegionRect) Then
                        graphicObjectIsVisible = True
                        Exit Function
                    End If
                Next
            End If
        Next
    Else
        graphicObjectIsVisible = rectOverlaps(.boundingRect, mRegionRect)
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

Private Sub undrawCrosshairs()
Dim nearestGridLine As Single
Dim nextNearestGridLine As Single
Dim prevLineX As Single

If mSuppressDrawing Then Exit Sub

With mSurface
    .DrawWidth = 1
    .DrawStyle = LineStyles.LineSolid
    .DrawMode = vbXorPen
    mSurface.Line (.ScaleLeft, mPrevCursorY)-(YAxisPosition, mPrevCursorY), vbRed Xor .backColor
    prevLineX = Round(mPrevCursorX)
'    If prevLineX < 1 Then prevLineX = 1
    mSurface.Line (prevLineX, .ScaleTop + .ScaleHeight)-(prevLineX, .ScaleTop), vbRed Xor .backColor
    .ForeColor = .backColor
    .CurrentX = YAxisPosition + 0.5
    .CurrentY = mPrevCursorY - mGridTextHeight / 2
    .FontTransparent = False
    mSurface.Print Format(mPrevCursorY, mYScaleSubFormatStr)
    If mScaleGridSpacingY <> 0 Then
        nearestGridLine = Fix((mPrevCursorY + mScaleGridSpacingY / 2) / mScaleGridSpacingY) * mScaleGridSpacingY
        If Abs(mPrevCursorY - nearestGridLine) < Abs(mGridTextHeight) Then
            .ForeColor = vbWhite Xor .backColor
            .CurrentX = YAxisPosition + 0.5
            .CurrentY = nearestGridLine - mGridTextHeight / 2
            .FontTransparent = True
            mSurface.Print Format(nearestGridLine, mYScaleFormatStr)
        End If
        If nearestGridLine >= mPrevCursorY Then
            nextNearestGridLine = nearestGridLine - mScaleGridSpacingY
        ElseIf nearestGridLine < mPrevCursorY Then
            nextNearestGridLine = nearestGridLine + mScaleGridSpacingY
        End If
        If Abs(mPrevCursorY - nextNearestGridLine) < Abs(mGridTextHeight) Then
            .ForeColor = vbWhite Xor .backColor
            .CurrentX = YAxisPosition + 0.5
            .CurrentY = nextNearestGridLine - mGridTextHeight / 2
            .FontTransparent = True
            mSurface.Print Format(nextNearestGridLine, mYScaleFormatStr)
        End If
    End If
End With

End Sub

Private Sub undrawPointer()
Dim nearestGridLine As Single
Dim nextNearestGridLine As Single
Dim prevLineX As Single

If mSuppressDrawing Then Exit Sub

With mSurface
    .DrawWidth = 1
    .DrawStyle = LineStyles.LineSolid
    .DrawMode = vbXorPen
    prevLineX = Round(mPrevCursorX)
'    If prevLineX < 1 Then prevLineX = 1
    .FillColor = vbYellow Xor .backColor
    .FillStyle = vbSolid
    mSurface.Circle (prevLineX, mPrevCursorY), newPoint(0.8, 0, PositionDistance, PositionDistance).XAbsolute - mSurface.ScaleLeft, vbYellow Xor .backColor
    ''debug.print "UndrawPointer x=" & prevLineX & " y=" & mPrevCursorY
    .ForeColor = .backColor
    .CurrentX = YAxisPosition + 0.5
    .CurrentY = mPrevCursorY - mGridTextHeight / 2
    .FontTransparent = False
    mSurface.Print Format(mPrevCursorY, mYScaleSubFormatStr)
    If mScaleGridSpacingY <> 0 Then
        nearestGridLine = Fix((mPrevCursorY + mScaleGridSpacingY / 2) / mScaleGridSpacingY) * mScaleGridSpacingY
        If Abs(mPrevCursorY - nearestGridLine) < Abs(mGridTextHeight) Then
            .ForeColor = vbWhite Xor .backColor
            .CurrentX = YAxisPosition + 0.5
            .CurrentY = nearestGridLine - mGridTextHeight / 2
            .FontTransparent = True
            mSurface.Print Format(nearestGridLine, mYScaleFormatStr)
        End If
        If nearestGridLine >= mPrevCursorY Then
            nextNearestGridLine = nearestGridLine - mScaleGridSpacingY
        ElseIf nearestGridLine < mPrevCursorY Then
            nextNearestGridLine = nearestGridLine + mScaleGridSpacingY
        End If
        If Abs(mPrevCursorY - nextNearestGridLine) < Abs(mGridTextHeight) Then
            .ForeColor = vbWhite Xor .backColor
            .CurrentX = YAxisPosition + 0.5
            .CurrentY = nextNearestGridLine - mGridTextHeight / 2
            .FontTransparent = True
            mSurface.Print Format(nextNearestGridLine, mYScaleFormatStr)
        End If
    End If
End With

End Sub


