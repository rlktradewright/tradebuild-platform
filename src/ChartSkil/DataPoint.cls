VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements IGraphicObject

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private mCanvas As canvas
Private mLayer As Long
Private mHandle As Long

Private mLocalStyle As DataPointStyle
Private mBaseStyle As DataPointStyle

Private mInScope As Boolean
Private mVisible As Boolean

Private mEventProxy As ChartRegionEventProxy

Private mBoundingRect As TRectangle

Private mX As Double
Private mKey As String

Private mIsSetLineThickness As Boolean
Private mIsSetColor As Boolean
Private mIsSetUpColor As Boolean
Private mIsSetDownColor As Boolean
Private mIsSetLineStyle As Boolean
Private mIsSetPointStyle As Boolean
Private mIsSetDisplayMode As Boolean
Private mIsSetHistBarWidth As Boolean
Private mIsSetIncludeInAutoscale As Boolean

Private mLineThickness As Long
Private mColor As Long
Private mUpColor As Long
Private mDownColor As Long
Private mLineStyle As LineStyles
Private mPointStyle As PointStyles
Private mDisplayMode As DataPointDisplayModes
Private mHistBarWidth As Single
Private mIncludeInAutoscale As Boolean

Private mDataValue As Double
Private mBlank As Boolean   ' indicates that no value has been assigned

Private mPrevDataPoint As DataPoint

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mBlank = True

mBoundingRect.isValid = False
End Sub

'================================================================================
' IGraphicObject Members
'================================================================================
Private Property Get IGraphicObject_boundingRectangle( _
                        ByVal xBoundsChanged As Boolean, _
                        ByVal yBoundsChanged As Boolean, _
                        ByVal gaugeChanged As Boolean, _
                        ByVal gaugeX As Double, _
                        ByVal gaugeY As Double, _
                        ByRef regionRect As TRectangle) As ChartSkilTypes.TRectangle
Static prevGaugeX As Double
Static prevGaugeY As Double
Static prevRect As TRectangle
Dim canvasChanges As Long

canvasChanges = mCanvas.compareTo(prevGaugeX, prevGaugeY, prevRect)
If (canvasChanges And CanvasComparisonCodes.GaugeChangedX) Or _
    (canvasChanges And CanvasComparisonCodes.GaugeChangedY) _
Then
    mBoundingRect = calcBoundingRect
End If

prevGaugeX = mCanvas.gaugeX
prevGaugeY = mCanvas.gaugeY
prevRect = mCanvas.boundary

IGraphicObject_boundingRectangle = mBoundingRect
End Property

Private Function IGraphicObject_boundingRectanglesAt(ByVal periodNumber As Long) As ChartSkilTypes.TRectangle()
'
End Function

Private Property Get IGraphicObject_capabilities() As Long
IGraphicObject_capabilities = capabilities.PartialRedraw
End Property

Private Sub IGraphicObject_draw( _
                ByRef areas() As TRectangle, _
                ByVal hdc As Long)
draw
End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = False
End Property

Private Sub IGraphicObject_finish()
Set mCanvas = Nothing
End Sub

Private Property Get IGraphicObject_gaugeDependent() As Boolean
Select Case displayMode
Case DataPointDisplayModePoint
    IGraphicObject_gaugeDependent = True
Case DataPointDisplayModeLine
    IGraphicObject_gaugeDependent = True
Case DataPointDisplayModeStep
    IGraphicObject_gaugeDependent = True
Case DataPointDisplayModeHistogram
    IGraphicObject_gaugeDependent = False
End Select
End Property

Private Property Let IGraphicObject_handle(ByVal value As Long)
mHandle = value
End Property

Private Property Get IGraphicObject_handle() As Long
IGraphicObject_handle = mHandle
End Property

Private Property Get IGraphicObject_includeInAutoscale() As Boolean
IGraphicObject_includeInAutoscale = includeInAutoscale
End Property

Private Property Let IGraphicObject_inScope(ByVal value As Boolean)
mInScope = value
End Property

Private Property Get IGraphicObject_layer() As Long
IGraphicObject_layer = layer
End Property

Private Property Get IGraphicObject_noDraw() As Boolean
IGraphicObject_noDraw = mBlank Or (displayMode = 0)
End Property

Private Property Get IGraphicObject_periodNumber() As Long
IGraphicObject_periodNumber = Int(mX)
End Property

Private Property Get IGraphicObject_scaleDependent() As Boolean
IGraphicObject_scaleDependent = False
End Property

Private Sub IGraphicObject_undraw( _
                ByRef area As TRectangle, _
                ByVal hdc As Long)
undraw
End Sub

Private Property Let IGraphicObject_visible(ByVal value As Boolean)
mVisible = value
End Property


'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get blank() As Boolean
blank = mBlank
End Property

Public Property Let Color(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "DataPoint" & ":" & "color", _
            "Value is not a valid color"

If mColor = value And mIsSetColor Then Exit Property
mColor = value
mIsSetColor = True
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get Color() As Long
Color = mBaseStyle.Color
If Not mLocalStyle Is Nothing Then Color = mLocalStyle.Color
If mIsSetColor Then Color = mColor
End Property

Public Property Get dataValue() As Double
dataValue = mDataValue
End Property

Public Property Let dataValue(value As Double)
If value = mDataValue Then Exit Property
undraw  ' remove current point
mBlank = False
mDataValue = value
update
End Property

Public Property Let displayMode(ByVal value As DataPointDisplayModes)
If mDisplayMode = value And mIsSetDisplayMode Then Exit Property
undraw
mDisplayMode = value
mIsSetDisplayMode = True
calcBoundingRect
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get displayMode() As DataPointDisplayModes
displayMode = mBaseStyle.displayMode
If Not mLocalStyle Is Nothing Then displayMode = mBaseStyle.displayMode
If mIsSetDisplayMode Then displayMode = mDisplayMode
End Property

Public Property Let downColor(ByVal value As Long)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "DataPoint" & ":" & "downColor", _
            "Value is not a valid color or -1"

If mDownColor = value And mIsSetDownColor Then Exit Property
mDownColor = value
mIsSetDownColor = True
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get downColor() As Long
downColor = mBaseStyle.downColor
If Not mLocalStyle Is Nothing Then downColor = mLocalStyle.downColor
If mIsSetDownColor Then downColor = mDownColor
If downColor = -1 Then downColor = Color
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Let histBarWidth(ByVal value As Single)
If value <= 0 Or value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "DataPoint" & ":" & "histBarWidth", _
            "histBarWidth must be greater than zero but not greater than 1"
If mHistBarWidth = value And mIsSetHistBarWidth Then Exit Property
undraw
mHistBarWidth = value
mIsSetHistBarWidth = True
calcBoundingRect
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get histBarWidth() As Single
histBarWidth = mBaseStyle.histBarWidth
If Not mLocalStyle Is Nothing Then histBarWidth = mLocalStyle.histBarWidth
If mIsSetHistBarWidth Then histBarWidth = mHistBarWidth
End Property

Public Property Let includeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value And mIsSetIncludeInAutoscale Then Exit Property
If Not value Then undraw
mIncludeInAutoscale = value
mIsSetIncludeInAutoscale = True
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get includeInAutoscale() As Boolean
includeInAutoscale = mBaseStyle.includeInAutoscale
If Not mLocalStyle Is Nothing Then includeInAutoscale = mLocalStyle.includeInAutoscale
If mIsSetIncludeInAutoscale Then includeInAutoscale = mIncludeInAutoscale
End Property

Public Property Get inScope() As Boolean
inScope = mInScope
End Property

Public Property Get key() As String
key = mKey
End Property

Public Property Let layer(ByVal value As Long)
mLayer = value
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Let lineStyle(ByVal value As LineStyles)
If mLineStyle = value And mIsSetLineStyle Then Exit Property
mLineStyle = value
mIsSetLineStyle = True
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get lineStyle() As LineStyles
lineStyle = mBaseStyle.lineStyle
If Not mLocalStyle Is Nothing Then lineStyle = mLocalStyle.lineStyle
If mIsSetLineStyle Then lineStyle = mLineStyle
End Property

Public Property Let lineThickness(ByVal value As Long)
If mLineThickness = value And mIsSetLineThickness Then Exit Property
undraw
mLineThickness = value
mIsSetLineThickness = True
calcBoundingRect
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get lineThickness() As Long
lineThickness = mBaseStyle.lineThickness
If Not mLocalStyle Is Nothing Then lineThickness = mLocalStyle.lineThickness
If mIsSetLineThickness Then lineThickness = mLineThickness
End Property

Public Property Get localStyle() As BarStyle
Set localStyle = mLocalStyle.clone
End Property

Public Property Let localStyle( _
                ByVal value As BarStyle)

undraw

Set mLocalStyle = value.clone
calcBoundingRect
mEventProxy.objectChanged mHandle, mBoundingRect

End Property

Public Property Let pointStyle(ByVal value As PointStyles)
If mPointStyle = value And mIsSetPointStyle Then Exit Property
mPointStyle = value
mIsSetPointStyle = True
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get pointStyle() As PointStyles
pointStyle = mBaseStyle.pointStyle
If Not mLocalStyle Is Nothing Then pointStyle = mLocalStyle.pointStyle
If mIsSetPointStyle Then pointStyle = mPointStyle
End Property

Public Property Get prevDataPoint() As DataPoint
Set prevDataPoint = mPrevDataPoint
End Property

Public Property Let prevDataPoint(ByVal value As DataPoint)
Set mPrevDataPoint = value
End Property

Public Property Let upColor(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "DataPoint" & ":" & "upColor", _
            "Value is not a valid color"

If mUpColor = value And mIsSetUpColor Then Exit Property
mUpColor = value
mIsSetUpColor = True
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get upColor() As Long
upColor = mBaseStyle.upColor
If Not mLocalStyle Is Nothing Then upColor = mLocalStyle.upColor
If mIsSetUpColor Then upColor = mUpColor
If upColor = -1 Then upColor = Color
End Property

Public Property Get visible() As Boolean
visible = mVisible
End Property

Friend Property Let x(ByVal value As Double)
mX = value
End Property

Friend Property Get x() As Double
x = mX
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub initialise( _
                ByVal key As String, _
                ByVal eventProxy As ChartRegionEventProxy, _
                ByVal layer As Long, _
                ByVal canvas As canvas, _
                ByVal baseStyle As DataPointStyle, _
                ByVal localStyle As DataPointStyle)
                
mKey = key
Set mEventProxy = eventProxy
mLayer = layer
Set mCanvas = canvas
Set mBaseStyle = baseStyle
If Not localStyle Is Nothing Then Set mLocalStyle = localStyle.clone ' local style cannot be shared
calcBoundingRect
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function calcBoundingRect() As TRectangle
Dim boundingRect As TRectangle
Dim thicknessScaleX As Double
Dim thicknessScaleY As Double
Dim prevX As Long
Dim prevValue As Double

thicknessScaleX = mCanvas.convertPixelsToLogicalX(lineThickness + 2)
thicknessScaleY = mCanvas.convertPixelsToLogicalY(lineThickness + 2)

If displayMode = DataPointDisplayModes.DataPointDisplayModeHistogram Then
    With boundingRect
        .left = mX - histBarWidth / 2
        .right = mX + histBarWidth / 2
        If mDataValue >= 0 Then
            .bottom = 0
            .top = mDataValue
        Else
            .bottom = mDataValue
            .top = 0
        End If
    End With
ElseIf drawAsPoint Then
    With boundingRect
        .left = mX - thicknessScaleX / 2
        .right = mX + thicknessScaleX / 2
        .bottom = mDataValue - thicknessScaleY / 2
        .top = mDataValue + thicknessScaleY / 2
    End With
ElseIf displayMode = DataPointDisplayModes.DataPointDisplayModeLine Then
    With mPrevDataPoint
        prevX = .x
        prevValue = .dataValue
    End With
    With boundingRect
        If prevX <= mX Then
            .left = prevX - thicknessScaleX / 2
            .right = mX + thicknessScaleX / 2
        Else
            .left = mX - thicknessScaleX / 2
            .right = prevX + thicknessScaleX / 2
        End If
        If mDataValue > prevValue Then
            .bottom = prevValue - thicknessScaleY / 2
            .top = mDataValue + thicknessScaleY / 2
        Else
            .bottom = mDataValue - thicknessScaleY / 2
            .top = prevValue + thicknessScaleY / 2
        End If
    End With
ElseIf displayMode = DataPointDisplayModes.DataPointDisplayModeStep Then
    With mPrevDataPoint
        prevX = .x
        prevValue = .dataValue
    End With
    With boundingRect
        If prevX <= mX Then
            .left = prevX - thicknessScaleX / 2
            .right = mX + thicknessScaleX / 2
        Else
            .left = mX - thicknessScaleX / 2
            .right = prevX + thicknessScaleX / 2
        End If
        If mDataValue > prevValue Then
            .bottom = prevValue - thicknessScaleY / 2
            .top = mDataValue + thicknessScaleY / 2
        Else
            .bottom = mDataValue - thicknessScaleY / 2
            .top = prevValue + thicknessScaleY / 2
        End If
    End With
End If
boundingRect.isValid = True
calcBoundingRect = boundingRect
End Function

Private Sub draw()
Dim lColor As Long

If Not mPrevDataPoint Is Nothing Then
    If mDataValue >= mPrevDataPoint.dataValue Then
        lColor = upColor
    Else
        lColor = downColor
    End If
Else
    lColor = Color
End If
drawPoint lColor
End Sub

Private Function drawAsPoint() As Boolean
If mPrevDataPoint Is Nothing Then
    drawAsPoint = True
ElseIf mPrevDataPoint.blank Then
    drawAsPoint = True
ElseIf displayMode = DataPointDisplayModes.DataPointDisplayModePoint Then
'ElseIf DisplayMode = DataPointDisplayModes.DataPointDisplayModePoint Or _
'    (DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine And Not mPrevDataPoint.inScope) Or _
'    (DisplayMode = DataPointDisplayModes.DataPointDisplayModeStep And Not mPrevDataPoint.inScope) _
'Then
    drawAsPoint = True
Else
    drawAsPoint = False
End If
End Function

Private Sub drawPoint(ByVal Color As Long)
If blank Then Exit Sub
If Not mVisible Then Exit Sub

If displayMode = 0 Then Exit Sub

mCanvas.setPenAttributes Color, _
                        lineThickness, _
                        lineStyle, _
                        DrawModes.DrawModeCopyPen

If displayMode = DataPointDisplayModes.DataPointDisplayModeHistogram Then
    mCanvas.setBrushAttributes Color, FillStyles.FillSolid
    mCanvas.drawRectangle mX - histBarWidth / 2, 0, mX + histBarWidth / 2, mDataValue
ElseIf drawAsPoint Then
    If pointStyle = PointRound Then
        mCanvas.drawPoint mX, mDataValue
    Else
        Dim thicknessScaleX As Double
        Dim thicknessScaleY As Double
        
        thicknessScaleX = mCanvas.convertPixelsToLogicalX(lineThickness)
        thicknessScaleY = mCanvas.convertPixelsToLogicalY(lineThickness)
        mCanvas.setPenAttributes Color, _
                                1, _
                                LineInsideSolid, _
                                DrawModes.DrawModeCopyPen
        mCanvas.setBrushAttributes Color, FillSolid
        mCanvas.drawRectangle mX - thicknessScaleX / 2, _
                                mDataValue - thicknessScaleY / 2, _
                                mX + thicknessScaleX / 2, _
                                mDataValue + thicknessScaleY / 2
    End If
ElseIf displayMode = DataPointDisplayModes.DataPointDisplayModeLine Then
    mCanvas.drawLine mPrevDataPoint.x, mPrevDataPoint.dataValue, mX, mDataValue
ElseIf displayMode = DataPointDisplayModes.DataPointDisplayModeStep Then
    mCanvas.drawLine mPrevDataPoint.x, mPrevDataPoint.dataValue, mX - 0.5, mPrevDataPoint.dataValue
    mCanvas.drawLine mX - 0.5, mPrevDataPoint.dataValue, mX - 0.5, mDataValue
    mCanvas.drawLine mX - 0.5, mDataValue, mX, mDataValue
End If
End Sub

Private Sub undraw()
If Not mVisible Then Exit Sub
drawPoint mCanvas.backColor
mEventProxy.objectUndrawn mHandle, mBoundingRect
End Sub

Private Sub update()
If mHandle = 0 Then
    ' haven't yet been added to chart region
    Exit Sub
End If
mBoundingRect = calcBoundingRect
mEventProxy.objectChanged mHandle, mBoundingRect
End Sub

