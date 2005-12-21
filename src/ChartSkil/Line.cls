VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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

Private mCanvas  As canvas
Private mLayer As Long
Private mHandle As Long

Private mInScope As Boolean
Private mVisible As Boolean

Private mEventProxy As ChartRegionEventProxy

Private mBoundingRect As TRectangle
Private mStartArrowRect As TRectangle
Private mEndArrowRect As TRectangle
Private mRegionRect As TRectangle

Private mEndArrowX1Offset As Double   ' horiz offset from apex of upper arrow endpoint
Private mEndArrowY1Offset As Double   ' vert offset from apex of upper arrow endpoint
Private mEndArrowX2Offset As Double   ' horiz offset from apex of lower arrow endpoint
Private mEndArrowY2Offset As Double   ' vert offset from apex of lower arrow endpoint

Private mStartArrowX1Offset As Double   ' horiz offset from apex of upper arrow endpoint
Private mStartArrowY1Offset As Double   ' vert offset from apex of upper arrow endpoint
Private mStartArrowX2Offset As Double   ' horiz offset from apex of lower arrow endpoint
Private mStartArrowY2Offset As Double   ' vert offset from apex of lower arrow endpoint

Private mGaugeX As Double
Private mGaugeY As Double

Private mThicknessOffsetX As Double
Private mThicknessOffsetY As Double

Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

Private mBlank As Boolean
Private mHidden As Boolean

Private mColour As Long
Private mThickness As Long
Private mStyle As LineStyles
Private mExtendBefore As Boolean
Private mExtendAfter As Boolean
Private mArrowStartStyle As ArrowStyles
Private mArrowStartLength As Long
Private mArrowStartWidth As Long
Private mArrowStartColor As Long
Private mArrowStartFillColor As Long
Private mArrowStartFillStyle As FillStyles
Private mArrowEndStyle As ArrowStyles
Private mArrowEndLength As Long
Private mArrowEndWidth As Long
Private mArrowEndColor As Long
Private mArrowEndFillColor As Long
Private mArrowEndFillStyle As FillStyles

Private mPoint1 As Point
Private mPoint2 As Point
Private mDirection As Quadrants
Private mVertical As Verticals
Private mLeftToRight As Boolean
Private mBottomToTop As Boolean
Private mA As Double    ' from the equation ...
Private mB As Double    ' ... y = ax + b
Private mLineAngle As Double    ' slope in radians
Private mLineAngleSin As Double
Private mLineAngleCos As Double

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mBlank = True

mBoundingRect.isValid = False
mStartArrowRect.isValid = False
mEndArrowRect.isValid = False
mRegionRect.isValid = False
End Sub

'================================================================================
' IGraphicObject Event Handlers
'================================================================================

Private Property Get IGraphicObject_boundingRectangle( _
                        ByVal xBoundsChanged As Boolean, _
                        ByVal yBoundsChanged As Boolean, _
                        ByVal gaugeChanged As Boolean, _
                        ByVal gaugeX As Double, _
                        ByVal gaugeY As Double, _
                        ByRef regionRect As TRectangle) As ChartSkilTypes.TRectangle

Dim canvasChanges As Long

canvasChanges = mCanvas.compareTo(mGaugeX, mGaugeY, mRegionRect)

mRegionRect = mCanvas.boundary
mGaugeX = mCanvas.gaugeX
mGaugeY = mCanvas.gaugeY

If (canvasChanges And CanvasComparisonCodes.GaugeChangedX) Or _
    (canvasChanges And CanvasComparisonCodes.GaugeChangedY) Or _
    (canvasChanges And CanvasComparisonCodes.BoundsChangedX) Or _
    (canvasChanges And CanvasComparisonCodes.BoundsChangedY) _
Then
    calcConstants
    calcEndArrowConstants
    calcStartArrowConstants
    calcCurrentBoundingRect
End If

IGraphicObject_boundingRectangle = mBoundingRect
End Property

Private Function IGraphicObject_boundingRectanglesAt(ByVal periodNumber As Long) As ChartSkilTypes.TRectangle()
Dim periodrect As TRectangle
Dim rects() As TRectangle

If mBlank Then Exit Function

With periodrect
    .left = periodNumber
    .right = periodNumber + 1
    .bottom = mRegionRect.bottom
    .top = mRegionRect.top
    .isValid = True
End With

If mVertical <> VerticalNot Then
    Dim vertRect As TRectangle
    
    vertRect = rectIntersection(mBoundingRect, periodrect)
    If vertRect.isValid Then
        
        ReDim rects(0) As TRectangle

        rects(0) = vertRect
    End If
Else
    If periodNumber = Int(mRegionRect.right) Or _
        (Not mLeftToRight And periodNumber = mPoint2.x And Not mExtendAfter) Or _
        (Not mLeftToRight And periodNumber = mPoint1.x And Not mExtendBefore) _
    Then
        If mThickness <= 1 Then Exit Function
    End If
    
    ReDim rects(0) As TRectangle
    
    With rects(0)
        .left = periodNumber
        If mLeftToRight Then
            If periodNumber < mPoint1.x And _
                Not mExtendBefore _
            Then
                .left = mPoint1.x - 1
            End If
        Else
            If periodNumber < mPoint2.x And _
                Not mExtendAfter _
            Then
                .left = mPoint2.x - 1
            End If
        End If
        If mLeftToRight Then
            If (Not mExtendAfter And periodNumber = mPoint2.x) Or _
                periodNumber = Int(mRegionRect.right) _
            Then
                .right = .left
            Else
                .right = periodNumber + 1
            End If
        Else
            If (Not mExtendBefore And periodNumber = mPoint1.x) Or _
                periodNumber = Int(mRegionRect.right) _
            Then
                .right = .left
            Else
                .right = periodNumber + 1
            End If
        End If
        
        If mA > 0 Then
            .bottom = mA * .left + mB - mThicknessOffsetY
            .top = mA * .right + mB + mThicknessOffsetY
        Else
            .top = mA * .left + mB + mThicknessOffsetY
            .bottom = mA * .right + mB - mThicknessOffsetY
        End If
        If mLeftToRight Then
            If (Not mExtendBefore And periodNumber = mPoint1.x) Then .left = .left - mThicknessOffsetX
            If (Not mExtendAfter And periodNumber = mPoint2.x) Or _
                periodNumber = Int(mRegionRect.right) _
            Then
                .right = .right + mThicknessOffsetX
            End If
        Else
            If (Not mExtendAfter And periodNumber = mPoint2.x) Then .left = .left - mThicknessOffsetX
            If (Not mExtendBefore And periodNumber = mPoint1.x) Or _
                periodNumber = Int(mRegionRect.right) _
            Then
                .right = .right + mThicknessOffsetX
            End If
        End If
        .isValid = True
    End With
End If

If Not mArrowEndStyle = ArrowNone Then
    If rectOverlaps(mEndArrowRect, periodrect) Then
        rects(0) = rectUnion(rects(0), mEndArrowRect)
    End If
End If
If Not mArrowStartStyle = ArrowNone Then
    If rectOverlaps(mStartArrowRect, periodrect) Then
        rects(0) = rectUnion(rects(0), mStartArrowRect)
    End If
End If
IGraphicObject_boundingRectanglesAt = rects
End Function

Private Property Get IGraphicObject_capabilities() As Long
IGraphicObject_capabilities = capabilities.ComplexBounding Or _
                            capabilities.PartialRedraw
End Property

Private Sub IGraphicObject_draw( _
                ByRef areas() As TRectangle, _
                ByVal hdc As Long)
Dim i As Long
Dim j As Long
Dim areasToRedraw() As TRectangle

ReDim areasToRedraw(UBound(areas)) As TRectangle

' amalagamate any adjacent areas
areasToRedraw(0) = areas(0)
For i = 1 To UBound(areas)
    If rectOverlaps(areasToRedraw(j), areas(i)) Then
        areasToRedraw(j) = rectUnion(areasToRedraw(j), areas(i))
    Else
        j = j + 1
        areasToRedraw(j) = areas(i)
    End If
Next

For i = j To 0 Step -1
    draw areasToRedraw(i), undraw:=False
Next
End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = isExtended
End Property

Private Sub IGraphicObject_finish()
Set mCanvas = Nothing
End Sub

Private Property Get IGraphicObject_gaugeDependent() As Boolean
IGraphicObject_gaugeDependent = True
End Property

Private Property Let IGraphicObject_handle(ByVal value As Long)
mHandle = value
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
IGraphicObject_noDraw = mHidden Or mBlank
End Property

Private Property Get IGraphicObject_periodNumber() As Long
IGraphicObject_periodNumber = -1
If isExtended Then Exit Sub
If mPoint1 Is Nothing Then Exit Sub
If Not mPoint1.isAssigned Then Exit Sub
IGraphicObject_periodNumber = mPoint1.XLogical
End Property

Private Property Get IGraphicObject_scaleDependent() As Boolean
IGraphicObject_scaleDependent = mFixedX Or mFixedY
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
' Properties
'================================================================================

Public Property Get arrowEndColor() As Long
arrowEndColor = mArrowEndColor
End Property

Public Property Let arrowEndColor(ByVal newValue As Long)
If mArrowEndColor = newValue Then Exit Property
undraw
mArrowEndColor = newValue
update
End Property

Public Property Get arrowEndFillColor() As Long
arrowEndFillColor = mArrowEndFillColor
End Property

Public Property Let arrowEndFillColor(ByVal newValue As Long)
If mArrowEndFillColor = newValue Then Exit Property
undraw
mArrowEndFillColor = newValue
update
End Property

Public Property Get arrowEndFillStyle() As FillStyles
arrowEndFillStyle = mArrowEndFillStyle
End Property

Public Property Let arrowEndFillStyle(ByVal newValue As FillStyles)
If mArrowEndFillStyle = newValue Then Exit Property
undraw
mArrowEndFillStyle = newValue
update
End Property

Public Property Get arrowEndLength() As Long
arrowEndLength = mArrowEndLength
End Property

Public Property Let arrowEndLength(ByVal newValue As Long)
If mArrowEndLength = newValue Then Exit Property
undraw
mArrowEndLength = newValue
update
End Property

Public Property Get arrowEndStyle() As ArrowStyles
arrowEndStyle = mArrowEndStyle
End Property

Public Property Let arrowEndStyle(ByVal newValue As ArrowStyles)
If mArrowEndStyle = newValue Then Exit Property
undraw
mArrowEndStyle = newValue
update
End Property

Public Property Get arrowEndWidth() As Long
arrowEndWidth = mArrowEndWidth
End Property

Public Property Let arrowEndWidth(ByVal newValue As Long)
If mArrowEndWidth = newValue Then Exit Property
undraw
mArrowEndWidth = newValue
update
End Property

Public Property Get arrowStartColor() As Long
arrowStartColor = mArrowStartColor
End Property

Public Property Let arrowStartColor(ByVal newValue As Long)
If mArrowStartColor = newValue Then Exit Property
undraw
mArrowStartColor = newValue
update
End Property

Public Property Get arrowStartFillColor() As Long
arrowStartFillColor = mArrowStartFillColor
End Property

Public Property Let arrowStartFillColor(ByVal newValue As Long)
If mArrowStartFillColor = newValue Then Exit Property
undraw
mArrowStartFillColor = newValue
update
End Property

Public Property Get arrowStartFillStyle() As FillStyles
arrowStartFillStyle = arrowStartFillStyle
End Property

Public Property Let arrowStartFillStyle(ByVal newValue As FillStyles)
If mArrowStartFillStyle = newValue Then Exit Property
undraw
mArrowStartFillStyle = newValue
update
End Property

Public Property Get arrowStartLength() As Long
arrowStartLength = mArrowStartLength
End Property

Public Property Let arrowStartLength(ByVal newValue As Long)
If mArrowStartLength = newValue Then Exit Property
undraw
mArrowStartLength = newValue
update
End Property

Public Property Get arrowStartStyle() As ArrowStyles
arrowStartStyle = mArrowStartStyle
End Property

Public Property Let arrowStartStyle(ByVal newValue As ArrowStyles)
If mArrowStartStyle = newValue Then Exit Property
undraw
mArrowStartStyle = newValue
update
End Property

Public Property Get arrowStartWidth() As Long
arrowStartWidth = mArrowStartWidth
End Property

Public Property Let arrowStartWidth(ByVal newValue As Long)
If mArrowStartWidth = newValue Then Exit Property
undraw
mArrowStartWidth = newValue
update
End Property

Friend Property Let canvas(ByVal value As canvas)
Set mCanvas = value
End Property

Public Property Get color() As Long
color = mColour
End Property

Public Property Let color(ByVal newValue As Long)
If mColour = newValue Then Exit Property
undraw
mColour = newValue
update
End Property

Public Property Get direction() As Quadrants
direction = mDirection
End Property

Friend Property Let eventProxy(ByVal value As ChartRegionEventProxy)
Set mEventProxy = value
End Property

Friend Property Get extended() As Boolean
extended = mExtended
End Property

Public Property Let extended(ByVal newValue As Boolean)
If mExtended = newValue Then Exit Property
undraw
mExtended = newValue
update
End Property

Public Property Get extendAfter() As Boolean
extendAfter = mExtendAfter
End Property

Public Property Let extendAfter(ByVal newValue As Boolean)
If mExtendAfter = newValue Then Exit Property
undraw
mExtendAfter = newValue
update
End Property

Public Property Get extendBefore() As Boolean
extendBefore = mExtendBefore
End Property

Public Property Let extendBefore(ByVal newValue As Boolean)
If mExtendBefore = newValue Then Exit Property
undraw
mExtendBefore = newValue
update
End Property

Friend Property Get fixedX() As Boolean
fixedX = mFixedX
End Property

Public Property Let fixedX(ByVal newValue As Boolean)
If mFixedX = newValue Then Exit Property
undraw
mFixedX = newValue
convertPositionX
update
End Property

Friend Property Get fixedY() As Boolean
fixedY = mFixedY
End Property

Public Property Let fixedY(ByVal newValue As Boolean)
If mFixedY = newValue Then Exit Property
undraw
mFixedY = newValue
convertPositionY
update
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Get hidden() As Boolean
hidden = mHidden
End Property

Public Property Let hidden(ByVal newValue As Boolean)
If mHidden = newValue Then Exit Property
undraw
mHidden = newValue
update
End Property

Public Property Get includeInAutoscale() As Boolean
includeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let includeInAutoscale(ByVal newValue As Boolean)
If mIncludeInAutoscale = newValue Then Exit Property
undraw
mIncludeInAutoscale = newValue
update
End Property

Public Property Get inScope() As Boolean
inScope = mInScope
End Property

'Public Property Get keepInView() As Boolean
'keepInView = mKeepInView
'End Property

'Public Property Let keepInView(ByVal newValue As Boolean)
'If mKeepInView = newValue Then Exit Property
'undraw
'mKeepInView = newValue
'update
'End Property

Friend Property Let layer(ByVal value As Long)
mLayer = value
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Get point1() As Point
Set point1 = mPoint1
End Property

Public Property Let point1(ByVal value As Point)
If value Is Nothing Or Not value.isAssigned Then Exit Property
If Not mPoint1 Is Nothing Then
    If value.equals(mPoint1) Then Exit Property
End If
undraw
Set mPoint1 = value
convertPosition1
With mPoint1
    If mPoint2 Is Nothing Then Exit Property
    If Not mPoint2.isAssigned Then Exit Property
    setDirection
    mBlank = False
    update
End With
End Property

Public Property Get point2() As Point
Set point2 = mPoint2
End Property

Public Property Let point2(ByVal value As Point)
If value Is Nothing Or Not value.isAssigned Then Exit Property
If Not mPoint2 Is Nothing Then
    If value.equals(mPoint2) Then Exit Property
End If
undraw
Set mPoint2 = value
convertPosition2
With mPoint2
    If mPoint1 Is Nothing Then Exit Property
    If Not mPoint1.isAssigned Then Exit Property
    setDirection
    mBlank = False
    update
End With
End Property

Public Property Get style() As LineStyles
style = mStyle
End Property

Public Property Let style(ByVal newValue As LineStyles)
If mStyle = newValue Then Exit Property
undraw
mStyle = newValue
update
End Property

Public Property Get thickness() As Long
thickness = mThickness
End Property

Public Property Let thickness(ByVal newValue As Long)
If mThickness = newValue Then Exit Property
undraw
mThickness = newValue
update
End Property

Public Property Get vertical() As Verticals
vertical = mVertical
End Property

Public Property Get visible() As Boolean
visible = mVisible
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcConstants()
Dim x1 As Double
Dim x2 As Double
Dim y1 As Double
Dim y2 As Double

If mBlank Then Exit Sub

mGaugeX = mCanvas.gaugeX
mGaugeY = mCanvas.gaugeY

mThicknessOffsetX = mGaugeX * mThickness / 2
mThicknessOffsetY = mGaugeY * mThickness / 2

If mVertical = VerticalUp Then
    mLineAngle = Pi / 2
    Exit Sub
ElseIf mVertical = VerticalDown Then
    mLineAngle = 3 * Pi / 2
    Exit Sub
End If

With mPoint1
    x1 = .XLogical
    y1 = .YLogical
End With
With mPoint2
    x2 = .XLogical
    y2 = .YLogical
End With

mA = (y2 - y1) / (x2 - x1)
mB = (x2 * y1 - x1 * y2) / (x2 - x1)

Select Case mDirection
Case NE
    mLineAngle = Atn(mA * mGaugeX / mGaugeY)
Case NW
    mLineAngle = Pi + Atn(mA * mGaugeX / mGaugeY)
Case SW
    mLineAngle = Pi + Atn(mA * mGaugeX / mGaugeY)
Case SE
    mLineAngle = 2 * Pi + Atn(mA * mGaugeX / mGaugeY)
End Select

mLineAngleSin = Sin(mLineAngle)
mLineAngleCos = Cos(mLineAngle)

'If mLineAngleSin <> 0 Then
'    mThicknessOffsetX = Abs(mGaugeX * (mThickness / mLineAngleSin) / 2)
'Else
'    mThicknessOffsetX = 0
'End If
'
'If mLineAngleCos <> 0 Then
'    mThicknessOffsetY = Abs(mGaugeY * (mThickness / mLineAngleCos) / 2)
'Else
'    mThicknessOffsetY = 0
'End If

End Sub

Private Sub calcCurrentBoundingRect()
Dim lineInt As TInterval
Dim regionInt As TInterval

Dim x1 As Double
Dim x2 As Double
Dim y1 As Double
Dim y2 As Double

mBoundingRect.isValid = False
If mBlank Then Exit Sub

With mPoint1
    x1 = .XLogical
    y1 = .YLogical
End With
With mPoint2
    x2 = .XLogical
    y2 = .YLogical
End With

With mBoundingRect
    If mVertical <> VerticalNot Then
        If mBottomToTop Then
            If mExtendBefore Then
                lineInt.startValue = MinusInfinityDouble
            Else
                lineInt.startValue = y1
            End If
            If mExtendAfter Then
                lineInt.endValue = PlusInfinityDouble
            Else
                lineInt.endValue = y2
            End If
        Else
            If mExtendAfter Then
                lineInt.startValue = MinusInfinityDouble
            Else
                lineInt.startValue = y2
            End If
            If mExtendBefore Then
                lineInt.endValue = PlusInfinityDouble
            Else
                lineInt.endValue = y1
            End If
        End If
        lineInt.isValid = True
        regionInt = rectGetYInterval(mRegionRect)
        rectSetYInterval mBoundingRect, intIntersection(regionInt, lineInt)
        
        .left = x1 - mThicknessOffsetX
        .right = x1 + mThicknessOffsetX
        .bottom = .bottom - mThicknessOffsetY
        .top = .top + mThicknessOffsetY
        .isValid = True
    Else
        If mLeftToRight Then
            If mExtendBefore Then
                lineInt.startValue = MinusInfinityDouble
            Else
                lineInt.startValue = x1
            End If
            If mExtendAfter Then
                lineInt.endValue = PlusInfinityDouble
            Else
                lineInt.endValue = x2
            End If
        Else
            If mExtendAfter Then
                lineInt.startValue = MinusInfinityDouble
            Else
                lineInt.startValue = x2
            End If
            If mExtendBefore Then
                lineInt.endValue = PlusInfinityDouble
            Else
                lineInt.endValue = x1
            End If
        End If
        lineInt.isValid = True
        regionInt = rectGetXInterval(mRegionRect)
        rectSetXInterval mBoundingRect, intIntersection(regionInt, lineInt)
        
        If Not .isValid Then
            Exit Sub
        End If
            
        If .right > Int(mRegionRect.right) Then .right = Int(mRegionRect.right)
        
        Select Case mDirection
        Case NE
            .top = mA * .right + mB + mThicknessOffsetY
            .bottom = mA * .left + mB - mThicknessOffsetY
        Case NW
            .top = mA * .left + mB + mThicknessOffsetY
            .bottom = mA * .right + mB - mThicknessOffsetY
        Case SW
            .bottom = mA * .left + mB - mThicknessOffsetY
            .top = mA * .right + mB + mThicknessOffsetY
        Case SE
            .bottom = mA * .right + mB - mThicknessOffsetY
            .top = mA * .left + mB + mThicknessOffsetY
        End Select
        
        If mThickness > 1 Then
            .right = .right + mThicknessOffsetX
            .left = .left - mThicknessOffsetX
        End If
        .isValid = True
    End If
'Debug.Print "Calc bounds : (" & .left & "," & (mA * .left + mB) & ")-(" & .right & "," & (mA * .right + mB) & ")"
End With

If mArrowEndStyle <> ArrowNone Then mBoundingRect = rectUnion(mBoundingRect, mEndArrowRect)
If mArrowStartStyle <> ArrowNone Then mBoundingRect = rectUnion(mBoundingRect, mStartArrowRect)

End Sub

Private Sub calcEndArrowConstants()
Dim hyp As Double
Dim halfAngle As Double
Dim x As Double
Dim y As Double

If mArrowEndStyle = ArrowNone Then Exit Sub
If mBlank Then Exit Sub

hyp = Sqr(mArrowEndLength * mArrowEndLength + mArrowEndWidth * mArrowEndWidth / 4)
halfAngle = Atn(mArrowEndWidth / (2 * mArrowEndLength))
mEndArrowX1Offset = -hyp * Cos(mLineAngle - halfAngle) * mGaugeX
mEndArrowY1Offset = -hyp * Sin(mLineAngle - halfAngle) * mGaugeY
mEndArrowX2Offset = -hyp * Sin(Pi / 2 - mLineAngle - halfAngle) * mGaugeX
mEndArrowY2Offset = -hyp * Cos(Pi / 2 - mLineAngle - halfAngle) * mGaugeY

With mPoint2
    x = .XLogical
    y = .YLogical
End With

With mEndArrowRect
    Select Case mDirection
    Case NE
        .bottom = y + mEndArrowY2Offset - mThicknessOffsetY
        .left = x + mEndArrowX1Offset - mThicknessOffsetX
        .top = y + IIf(mEndArrowY1Offset > 0, mEndArrowY1Offset, 0) + mThicknessOffsetY
        .right = x + IIf(mEndArrowX2Offset > 0, mEndArrowX2Offset, 0) + mThicknessOffsetX
        .isValid = True
    Case NW
        .bottom = y + mEndArrowY1Offset - mThicknessOffsetY
        .left = x + IIf(mEndArrowX1Offset < 0, mEndArrowX1Offset, 0) - mThicknessOffsetX
        .top = y + IIf(mEndArrowY2Offset > 0, mEndArrowY2Offset, 0) + mThicknessOffsetY
        .right = x + mEndArrowX2Offset + mThicknessOffsetX
        .isValid = True
    Case SW
        .bottom = y + IIf(mEndArrowY1Offset < 0, mEndArrowY1Offset, 0) - mThicknessOffsetY
        .left = x + IIf(mEndArrowX2Offset < 0, mEndArrowX2Offset, 0) - mThicknessOffsetX
        .top = y + mEndArrowY2Offset + mThicknessOffsetY
        .right = x + mEndArrowX1Offset + mThicknessOffsetX
        .isValid = True
    Case SE
        .bottom = y + IIf(mEndArrowY2Offset < 0, mEndArrowY2Offset, 0) - mThicknessOffsetY
        .left = x + mEndArrowX2Offset - mThicknessOffsetX
        .top = y + mEndArrowY1Offset + mThicknessOffsetY
        .right = x + IIf(mEndArrowX1Offset > 0, mEndArrowX1Offset, 0) + mThicknessOffsetX
        .isValid = True
    End Select
    'Debug.Assert (.right - .left) < 5 And (.top - .bottom) < 5
End With
End Sub

Private Sub calcStartArrowConstants()
Dim hyp As Double
Dim halfAngle As Double
Dim x As Double
Dim y As Double

If mArrowStartStyle = ArrowNone Then Exit Sub
If mBlank Then Exit Sub

hyp = Sqr(mArrowStartLength * mArrowStartLength + mArrowStartWidth * mArrowStartWidth / 4)
halfAngle = Atn(mArrowStartWidth / (2 * mArrowStartLength))
mStartArrowX1Offset = hyp * Cos(mLineAngle - halfAngle) * mGaugeX
mStartArrowY1Offset = hyp * Sin(mLineAngle - halfAngle) * mGaugeY
mStartArrowX2Offset = hyp * Sin(Pi / 2 - mLineAngle - halfAngle) * mGaugeX
mStartArrowY2Offset = hyp * Cos(Pi / 2 - mLineAngle - halfAngle) * mGaugeY

With mPoint1
    x = .XLogical
    y = .YLogical
End With

With mStartArrowRect
    Select Case mDirection
    Case NE
        .bottom = y + IIf(mStartArrowY1Offset < 0, mStartArrowY1Offset, 0) - mThicknessOffsetY
        .left = x + IIf(mStartArrowX2Offset < 0, mStartArrowX2Offset, 0) - mThicknessOffsetX
        .top = y + mStartArrowY2Offset + mThicknessOffsetY
        .right = x + mStartArrowX1Offset + mThicknessOffsetY
        .isValid = True
    Case NW
        .bottom = y + IIf(mEndArrowY2Offset < 0, mEndArrowY2Offset, 0) - mThicknessOffsetY
        .left = x + mEndArrowX2Offset - mThicknessOffsetX
        .top = y + mEndArrowY1Offset + mThicknessOffsetY
        .right = x + IIf(mEndArrowX1Offset > 0, mEndArrowX1Offset, 0) + mThicknessOffsetX
        .isValid = True
    Case SW
        .bottom = y + mStartArrowY2Offset - mThicknessOffsetY
        .left = x + mStartArrowX1Offset - mThicknessOffsetX
        .top = y + IIf(mStartArrowY1Offset > 0, mStartArrowY1Offset, 0) + mThicknessOffsetY
        .right = x + IIf(mStartArrowX2Offset > 0, mStartArrowX2Offset, 0) + mThicknessOffsetX
        .isValid = True
    Case SE
        .bottom = y + mStartArrowY1Offset - mThicknessOffsetY
        .left = x + IIf(mStartArrowX1Offset < 0, mStartArrowX1Offset, 0) - mThicknessOffsetX
        .top = y + IIf(mStartArrowY2Offset > 0, mStartArrowY2Offset, 0) + mThicknessOffsetY
        .right = x + mStartArrowX2Offset + mThicknessOffsetX
        .isValid = True
        
'    Case SW
'        .bottom = y + mStartArrowY2Offset - mThicknessOffsetY
'        .left = x + mStartArrowX1Offset - mThicknessOffsetX
'        .top = y + IIf(mStartArrowY1Offset > 0, mStartArrowY1Offset, 0) + mThicknessOffsetY
'        .right = x + IIf(mStartArrowX2Offset > 0, mStartArrowX2Offset, 0) + mThicknessOffsetX
'        .isvalid = True
'    Case SE
'        .bottom = y + mStartArrowY1Offset - mThicknessOffsetY
'        .left = x + IIf(mStartArrowX1Offset < 0, mStartArrowX1Offset, 0) - mThicknessOffsetX
'        .top = y + IIf(mStartArrowY2Offset > 0, mStartArrowY2Offset, 0) + mThicknessOffsetY
'        .right = x + mStartArrowX2Offset + mThicknessOffsetX
'        .isvalid = True
'    Case NE
'        .bottom = y + IIf(mStartArrowY1Offset < 0, mStartArrowY1Offset, 0) - mThicknessOffsetY
'        .left = x + IIf(mStartArrowX2Offset < 0, mStartArrowX2Offset, 0) - mThicknessOffsetX
'        .top = y + mStartArrowY2Offset + mThicknessOffsetY
'        .right = x + mStartArrowX1Offset + mThicknessOffsetX
'        .isvalid = True
'    Case NW
'        .bottom = y + IIf(mStartArrowY2Offset < 0, mStartArrowY2Offset, 0) - mThicknessOffsetY
'        .left = x + mStartArrowX2Offset - mThicknessOffsetX
'        .top = y + mStartArrowY1Offset + mThicknessOffsetY
'        .right = x + IIf(mStartArrowX1Offset > 0, mStartArrowX1Offset, 0) + mThicknessOffsetX
'        .isvalid = True
    End Select
    'Debug.Assert (.right - .left) < 5 And (.top - .bottom) < 5
End With
End Sub

Private Sub convertPositionX()
convertPosition1X
convertPosition2X
End Sub

Private Sub convertPositionY()
convertPosition1Y
convertPosition2Y
End Sub

Private Sub convertPosition1()
convertPosition1X
convertPosition1Y
End Sub

Private Sub convertPosition2()
convertPosition2X
convertPosition2Y
End Sub

Private Sub convertPosition1X()
If mPoint1 Is Nothing Then Exit Sub
If mFixedX Then
    If mPoint1.CoordinateSystemX = CoordsLogical Then
        mPoint1.CoordinateSystemX = CoordsRelative
    End If
Else
    If mPoint1.CoordinateSystemX <> CoordsLogical Then
        mPoint1.CoordinateSystemX = CoordsLogical
    End If
End If
End Sub

Private Sub convertPosition1Y()
If mPoint1 Is Nothing Then Exit Sub
If mFixedY Then
    If mPoint1.CoordinateSystemY = CoordsLogical Then
        mPoint1.CoordinateSystemY = CoordsRelative
    End If
Else
    If mPoint1.CoordinateSystemY <> CoordsLogical Then
        mPoint1.CoordinateSystemY = CoordsLogical
    End If
End If
End Sub

Private Sub convertPosition2X()
If mPoint2 Is Nothing Then Exit Sub
If mFixedX Then
    If mPoint2.CoordinateSystemX = CoordsLogical Then
        mPoint2.CoordinateSystemX = CoordsRelative
    End If
Else
    If mPoint2.CoordinateSystemX <> CoordsLogical Then
        mPoint2.CoordinateSystemX = CoordsLogical
    End If
End If
End Sub

Private Sub convertPosition2Y()
If mPoint2 Is Nothing Then Exit Sub
If mFixedY Then
    If mPoint2.CoordinateSystemY = CoordsLogical Then
        mPoint2.CoordinateSystemY = CoordsRelative
    End If
Else
    If mPoint2.CoordinateSystemY <> CoordsLogical Then
        mPoint2.CoordinateSystemY = CoordsLogical
    End If
End If
End Sub

Private Sub draw(ByRef area As TRectangle, ByVal undraw As Boolean)
Dim right As Double
Dim left As Double
Dim bottom As Double
Dim top As Double
Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double
Dim vertices(3) As Point
Dim slopeThicknessY As Double
Dim xBottom As Double
Dim xTop As Double

If mBlank Or mHidden Then Exit Sub
If Not mVisible Then
    Exit Sub
End If

mCanvas.setPenAttributes IIf(undraw, mCanvas.backColor, mColour), _
                        mThickness, _
                        mStyle, _
                        DrawModes.DrawModeCopyPen
    
With mPoint1
    x1 = .x
    y1 = .y
End With

With mPoint2
    x2 = .XLogical
    y2 = .YLogical
End With

If mVertical <> VerticalNot Then
    With area
        If mBottomToTop Then
            If mExtendAfter Then
                top = .top
            Else
                If .top > y2 Then
                    top = y2
                Else
                    top = .top
                End If
            End If
            If mExtendBefore Then
                bottom = .bottom
            Else
                If .bottom < y1 Then
                    bottom = y1
                Else
                    bottom = .bottom
                End If
            End If
        Else
            If mExtendAfter Then
                bottom = .bottom
            Else
                If .bottom < y2 Then
                    bottom = y2
                Else
                    bottom = .bottom
                End If
            End If
            If mExtendBefore Then
                top = .top
            Else
                If .top > y1 Then
                    top = y1
                Else
                    top = .top
                End If
            End If
        End If
    End With
    mCanvas.drawLine x1, bottom, x1, top
    mCanvas.drawPoint x1, top
Else
    With area
        
        left = .left
        right = .right
        
        If mLeftToRight And Not mExtendBefore Then
            If left < x1 Then left = x1
            If right < x1 Then right = x1
        ElseIf Not mLeftToRight And Not mExtendAfter Then
            If left < x2 Then left = x2
            If right < x2 Then right = x2
        End If
            
        If mLeftToRight And Not mExtendAfter Then
            If left > x2 Then left = x2
            If right > x2 Then right = x2
        ElseIf Not mLeftToRight And Not mExtendBefore Then
            If left > x1 Then left = x1
            If right > x1 Then right = x1
        End If
    
        If left > Int(mRegionRect.right) Then
            left = Int(mRegionRect.right)
        End If
        
        If right > Int(mRegionRect.right) Then
            right = Int(mRegionRect.right)
        End If
    
'        If mLeftToRight And .right > x2 And Not mExtendAfter Then
'            right = x2
'        ElseIf Not mLeftToRight And .right > x1 And Not mExtendBefore Then
'            right = x1
'        Else
'            right = .right
'        End If
'
'        If right > Int(mRegionRect.right) Then
'            right = Int(mRegionRect.right)
'        End If
'
'        If mLeftToRight And .left < x1 And Not mExtendBefore Then
'            left = x1
'        ElseIf Not mLeftToRight And .left < x2 And Not mExtendAfter Then
'            left = x2
'        Else
'            left = .left
'        End If
'
'        If mLeftToRight And left > x2 And Not mExtendAfter Then
'            left = right
'        ElseIf Not mLeftToRight And left > x1 And Not mExtendBefore Then
'            left = right
'        End If
    End With
    
    With mCanvas
        Debug.Assert left <= right
        If left <> right Then
            If mThickness = 1 Then
                ' now calculate the x coordinates that ensure the line
                ' does not go outside the top or bottom of the area
                ' to be drawn
                If mA = 0 Then
                    xBottom = MinusInfinityDouble
                Else
                    xBottom = (area.bottom - mB) / mA
                End If
                If mA = 0 Then
                    xTop = PlusInfinityDouble
                Else
                    xTop = (area.top - mB) / mA
                End If
                If xBottom > left And xBottom < right Then
                    If mA >= 0 Then
                        left = xBottom
                    Else
                        right = xBottom
                    End If
                End If
                If xTop > left And xTop < right Then
                    If mA >= 0 Then
                        right = xTop
                    Else
                        left = xTop
                    End If
                End If
                .drawLine left, mA * left + mB, right, mA * right + mB
            Else
                .setPenAttributes mColour, _
                                mThickness, _
                                LineStyles.LineSolid, _
                                DrawModes.DrawModeCopyPen
                If left = x1 Then
                    .drawPoint left, y1
                ElseIf left = x2 Then
                    .drawPoint left, y2
                End If
                
                slopeThicknessY = Abs(mThickness / mLineAngleCos)
                                
                If left = x1 Or left = x2 Then
                    Set vertices(0) = .newPoint(left - .convertPixelsToLogicalX(mThickness) * mLineAngleSin / 2, _
                                            mA * left + mB - .convertPixelsToLogicalY(mThickness) * mLineAngleCos / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                Else
                    Set vertices(0) = .newPoint(left, _
                                            mA * left + mB - .convertPixelsToLogicalY(slopeThicknessY - 2) / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                End If
                
                If left = x1 Or left = x2 Then
                    Set vertices(1) = .newPoint(left + .convertPixelsToLogicalX(mThickness) * mLineAngleSin / 2, _
                                            mA * left + mB + .convertPixelsToLogicalY(mThickness) * mLineAngleCos / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                Else
                    Set vertices(1) = .newPoint(left, _
                                            mA * left + mB + .convertPixelsToLogicalY(slopeThicknessY - 2) / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                End If
                
                If right = x1 Or right = x2 Then
                    Set vertices(2) = .newPoint(right + .convertPixelsToLogicalX(mThickness) * mLineAngleSin / 2, _
                                            mA * right + mB + .convertPixelsToLogicalY(mThickness) * mLineAngleCos / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                Else
                    Set vertices(2) = .newPoint(right, _
                                            mA * right + mB + .convertPixelsToLogicalY(slopeThicknessY - 2) / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                End If
                
                If right = x1 Or right = x2 Then
                    Set vertices(3) = .newPoint(right - .convertPixelsToLogicalX(mThickness) * mLineAngleSin / 2, _
                                            mA * right + mB - .convertPixelsToLogicalY(mThickness) * mLineAngleCos / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                Else
                    Set vertices(3) = .newPoint(right, _
                                            mA * right + mB - .convertPixelsToLogicalY(slopeThicknessY - 2) / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                End If
                .setPenAttributes mColour, _
                                1, _
                                LineStyles.LineSolid, _
                                DrawModes.DrawModeCopyPen
                .setBrushAttributes mColour, _
                                FillStyles.FillSolid
                
                .drawPolygon vertices
                
                .setPenAttributes mColour, _
                                mThickness, _
                                LineStyles.LineSolid, _
                                DrawModes.DrawModeCopyPen
                If right = x1 Then
                    .drawPoint right, y1
                ElseIf right = x2 Then
                    .drawPoint right, y2
                End If
                
                Debug.Assert IIf(mLeftToRight, right <= x2, True)
                Debug.Assert IIf(mLeftToRight, True, right <= x1)
            End If
        End If
    End With
End If
If mArrowEndStyle <> ArrowNone Then
    drawArrowEnd area, undraw
End If
If mArrowStartStyle <> ArrowNone Then
    drawArrowStart area, undraw
End If
    
End Sub

Private Sub drawArrowEnd(ByRef area As TRectangle, ByVal undraw As Boolean)
Dim x As Double
Dim y As Double
Dim vertices(2) As Point

With mPoint2
    x = .XLogical
    y = .YLogical
End With

If Not rectOverlaps(mEndArrowRect, area) Then Exit Sub

mCanvas.setPenAttributes IIf(undraw, mCanvas.backColor, mColour), _
                        mThickness, _
                        mStyle, _
                        DrawModes.DrawModeCopyPen

Select Case mArrowEndStyle
Case ArrowNone

Case ArrowSingleOpen
    mCanvas.drawLine x, y, x + mEndArrowX1Offset, y + mEndArrowY1Offset
    mCanvas.drawLine x, y, x + mEndArrowX2Offset, y + mEndArrowY2Offset
Case ArrowDoubleOpen

Case ArrowClosed
    With mCanvas
        .setBrushAttributes IIf(undraw, .backColor, mArrowEndFillColor), mArrowEndFillStyle
        
        Set vertices(0) = .newPoint(x, y)
        Set vertices(1) = .newPoint(x + mEndArrowX1Offset, _
                                    y + mEndArrowY1Offset)
        Set vertices(2) = .newPoint(x + mEndArrowX2Offset, _
                                    y + mEndArrowY2Offset)
        .drawPolygon vertices
    End With
Case ArrowSingleBar

Case ArrowDoubleBar

Case ArrowLollipop

Case ArrowDiamond

Case ArrowBarb

End Select

End Sub

Private Sub drawArrowStart(ByRef area As TRectangle, ByVal undraw As Boolean)
Dim x As Double
Dim y As Double
Dim vertices(2) As Point

With mPoint1
    x = .XLogical
    y = .YLogical
End With

If Not rectOverlaps(mStartArrowRect, area) Then Exit Sub

mCanvas.setPenAttributes IIf(undraw, mCanvas.backColor, mColour), _
                        mThickness, _
                        mStyle, _
                        DrawModes.DrawModeCopyPen

Select Case mArrowStartStyle
Case ArrowNone

Case ArrowSingleOpen
    mCanvas.drawLine x, y, x + mStartArrowX1Offset, y + mStartArrowY1Offset
    mCanvas.drawLine x, y, x + mStartArrowX2Offset, y + mStartArrowY2Offset
Case ArrowDoubleOpen

Case ArrowClosed
    With mCanvas
        .setBrushAttributes IIf(undraw, .backColor, mArrowEndFillColor), mArrowEndFillStyle
        
        Set vertices(0) = .newPoint(x, y)
        Set vertices(1) = .newPoint(x + mStartArrowX1Offset, _
                                    y + mStartArrowY1Offset)
        Set vertices(2) = .newPoint(x + mStartArrowX2Offset, _
                                    y + mStartArrowY2Offset)
        .drawPolygon vertices
    End With
Case ArrowSingleBar

Case ArrowDoubleBar

Case ArrowLollipop

Case ArrowDiamond

Case ArrowBarb

End Select

End Sub

Private Function isExtended() As Boolean
isExtended = mExtended Or mFixedX Or mFixedY
End Function

Private Sub setDirection()
Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double

With mPoint1
    x1 = .XLogical
    y1 = .YLogical
End With

With mPoint2
    x2 = .XLogical
    y2 = .YLogical
End With

If x1 = x2 Then
    If y1 <= y2 Then
        mVertical = VerticalUp
    Else
        mVertical = VerticalDown
    End If
Else
    mVertical = VerticalNot
End If
If x1 <= x2 Then
    If y1 <= y2 Then
        mDirection = NE
        mLeftToRight = True
        mBottomToTop = True
    Else
        mDirection = SE
        mLeftToRight = True
        mBottomToTop = False
    End If
Else
    If y1 <= y2 Then
        mDirection = NW
        mLeftToRight = False
        mBottomToTop = True
    Else
        mDirection = SW
        mLeftToRight = False
        mBottomToTop = False
    End If
End If
End Sub

Private Sub undraw()

If Not mVisible Then Exit Sub
draw mBoundingRect, True
mEventProxy.objectUndrawn mHandle, mBoundingRect
Exit Sub

End Sub

Private Sub update()
If mHandle = 0 Then
    ' haven't yet been added to chart region
    Exit Sub
End If
calcConstants
calcEndArrowConstants
calcStartArrowConstants
calcCurrentBoundingRect
mEventProxy.objectChanged mHandle, mBoundingRect
'If mIncludeInAutoscale And mBoundingRect.isvalid Then
'    If mBoundingRect.top >= mSurface.ScaleTop Or mBoundingRect.bottom <= (mSurface.ScaleTop + mSurface.ScaleHeight) Then
'        mEventProxy.boundsExceeded
'    End If
'End If
End Sub


