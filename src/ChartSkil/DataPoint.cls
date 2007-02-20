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

Private mInScope As Boolean
Private mVisible As Boolean

Private mEventProxy As ChartRegionEventProxy

Private mBoundingRect As TRectangle

Private mHistBarWidth As Single

Private mIncludeInAutoscale  As Boolean

Private mX As Double
Private mKey As String

Private mDataValue As Double
Private mBlank As Boolean   ' indicates that no value has been assigned
Private mLineThickness As Long
Private mLineColor As Long
Private mLineStyle As LineStyles
Private mDisplayMode As DataPointDisplayModes

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
Select Case mDisplayMode
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
IGraphicObject_includeInAutoscale = mIncludeInAutoscale
End Property

Private Property Let IGraphicObject_inScope(ByVal value As Boolean)
mInScope = value
End Property

Private Property Get IGraphicObject_layer() As Long
IGraphicObject_layer = layer
End Property

Private Property Get IGraphicObject_noDraw() As Boolean
IGraphicObject_noDraw = mBlank Or (mDisplayMode = 0)
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

Friend Property Let canvas(ByVal value As canvas)
Set mCanvas = value
End Property

Public Property Get dataValue() As Double
dataValue = mDataValue
End Property

Public Property Let dataValue(newValue As Double)
If newValue = mDataValue Then Exit Property
undraw  ' remove current point
mBlank = False
mDataValue = newValue
update
End Property

Friend Property Let displayMode(ByVal value As DataPointDisplayModes)
mDisplayMode = value
End Property

Friend Property Let eventProxy(ByVal value As ChartRegionEventProxy)
Set mEventProxy = value
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Let histBarWidth(ByVal value As Single)
mHistBarWidth = value
End Property

Public Property Get includeInAutoscale() As Boolean
includeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let includeInAutoscale(ByVal newValue As Boolean)
mIncludeInAutoscale = newValue
End Property

Friend Property Let key(ByVal value As String)
mKey = value
End Property

Public Property Get key() As String
key = mKey
End Property

Friend Property Let layer(ByVal value As Long)
mLayer = value
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Let lineColor(ByVal value As Long)
mLineColor = value
End Property

Friend Property Let lineStyle(ByVal value As LineStyles)
mLineStyle = value
End Property

Friend Property Let lineThickness(ByVal value As Long)
mLineThickness = value
calcBoundingRect
End Property

Public Property Get inScope() As Boolean
inScope = mInScope
End Property

'Public Property Get periodNumber() As Long
'periodNumber = Int(mX)
'End Property

'Public Property Let periodNumber(ByVal newValue As Long)
'mx = newValue
'End Property

Public Property Get prevDataPoint() As DataPoint
Set prevDataPoint = mPrevDataPoint
End Property

Public Property Let prevDataPoint(ByVal newValue As DataPoint)
Set mPrevDataPoint = newValue
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

'================================================================================
' Helper Functions
'================================================================================

Private Function calcBoundingRect() As TRectangle
Dim boundingRect As TRectangle
Dim thicknessScaleX As Double
Dim thicknessScaleY As Double
Dim prevX As Long
Dim prevValue As Double

thicknessScaleX = mCanvas.convertPixelsToLogicalX(mLineThickness + 2)
thicknessScaleY = mCanvas.convertPixelsToLogicalY(mLineThickness + 2)

If mDisplayMode = DataPointDisplayModes.DataPointDisplayModeHistogram Then
    With boundingRect
        .left = mX - mHistBarWidth / 2
        .right = mX + mHistBarWidth / 2
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
ElseIf mDisplayMode = DataPointDisplayModes.DataPointDisplayModeLine Then
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
ElseIf mDisplayMode = DataPointDisplayModes.DataPointDisplayModeStep Then
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
drawPoint mLineColor
End Sub

Private Function drawAsPoint() As Boolean
If mPrevDataPoint Is Nothing Then
    drawAsPoint = True
ElseIf mPrevDataPoint.blank Then
    drawAsPoint = True
ElseIf mDisplayMode = DataPointDisplayModes.DataPointDisplayModePoint Then
'ElseIf mDisplayMode = DataPointDisplayModes.DataPointDisplayModePoint Or _
'    (mDisplayMode = DataPointDisplayModes.DataPointDisplayModeLine And Not mPrevDataPoint.inScope) Or _
'    (mDisplayMode = DataPointDisplayModes.DataPointDisplayModeStep And Not mPrevDataPoint.inScope) _
'Then
    drawAsPoint = True
Else
    drawAsPoint = False
End If
End Function

Private Sub drawPoint(ByVal Color As Long)
If blank Then Exit Sub
If Not mVisible Then Exit Sub

If mDisplayMode = 0 Then Exit Sub

mCanvas.setPenAttributes Color, _
                        mLineThickness, _
                        mLineStyle, _
                        DrawModes.DrawModeCopyPen

If mDisplayMode = DataPointDisplayModes.DataPointDisplayModeHistogram Then
    mCanvas.setBrushAttributes Color, FillStyles.FillSolid
    mCanvas.drawRectangle mX - mHistBarWidth / 2, 0, mX + mHistBarWidth / 2, mDataValue
ElseIf drawAsPoint Then
    mCanvas.drawPoint mX, mDataValue
ElseIf mDisplayMode = DataPointDisplayModes.DataPointDisplayModeLine Then
    mCanvas.drawLine mPrevDataPoint.x, mPrevDataPoint.dataValue, mX, mDataValue
ElseIf mDisplayMode = DataPointDisplayModes.DataPointDisplayModeStep Then
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
'If mIncludeInAutoscale And mBoundingRect.isvalid Then
'    If mBoundingRect.top >= mSurface.ScaleTop Or mBoundingRect.bottom <= (mSurface.ScaleTop + mSurface.ScaleHeight) Then
'        mEventProxy.boundsExceeded
'    End If
'End If
End Sub

