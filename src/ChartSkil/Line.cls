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

'@================================================================================
' Interfaces
'@================================================================================

Implements IGraphicObject

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "ChartSkil26"
Private Const ModuleName                    As String = "Line"

'================================================================================
' Member variables
'================================================================================

Private mCanvas  As canvas
Private mLayer As Long
Private mHandle As Long

Private mLocalStyle As linestyle
Private mBaseStyle As linestyle

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

Private mBlank As Boolean
Private mHidden As Boolean

Private mIsSetColor As Boolean
Private mIsSetThickness As Boolean
Private mIsSetLineStyle As Boolean
Private mIsSetExtendBefore As Boolean
Private mIsSetExtendAfter As Boolean
Private mIsSetArrowStartStyle As Boolean
Private mIsSetArrowStartLength As Boolean
Private mIsSetArrowStartWidth As Boolean
Private mIsSetArrowStartColor As Boolean
Private mIsSetArrowStartFillColor As Boolean
Private mIsSetArrowStartFillStyle As Boolean
Private mIsSetArrowEndStyle As Boolean
Private mIsSetArrowEndLength As Boolean
Private mIsSetArrowEndWidth As Boolean
Private mIsSetArrowEndColor As Boolean
Private mIsSetArrowEndFillColor As Boolean
Private mIsSetArrowEndFillStyle As Boolean
Private mIsSetFixedX As Boolean
Private mIsSetFixedY As Boolean
Private mIsSetIncludeInAutoscale As Boolean
Private mIsSetExtended As Boolean

Private mColor As Long
Private mThickness As Long
Private mLineStyle As LineStyles
Private mExtendBefore As Boolean
Private mExtendAfter As Boolean
Private mArrowStartStyle As ArrowStyles
Private mArrowStartLength As Long
Private mArrowStartWidth As Long
Private mArrowStartColor As Long
Private mArrowStartFillColor As Long
Private marrowStartfillstyle As FillStyles
Private mArrowEndStyle As ArrowStyles
Private mArrowEndLength As Long
Private mArrowEndWidth As Long
Private mArrowEndColor As Long
Private mArrowEndFillColor As Long
Private mArrowEndFillStyle As FillStyles
Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

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

Private mKey As String

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

Dim failpoint As Long
On Error GoTo Err

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

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_boundingRectangle" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Property

Private Function IGraphicObject_boundingRectanglesAt(ByVal periodNumber As Long) As ChartSkilTypes.TRectangle()
Dim periodrect As TRectangle
Dim rects() As TRectangle

Dim failpoint As Long
On Error GoTo Err

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
        (Not mLeftToRight And periodNumber = mPoint2.x And Not extendAfter) Or _
        (Not mLeftToRight And periodNumber = mPoint1.x And Not extendBefore) _
    Then
        If thickness <= 1 Then Exit Function
    End If
    
    ReDim rects(0) As TRectangle
    
    With rects(0)
        .left = periodNumber
        If mLeftToRight Then
            If periodNumber < mPoint1.x And _
                Not extendBefore _
            Then
                .left = mPoint1.x - 1
            End If
        Else
            If periodNumber < mPoint2.x And _
                Not extendAfter _
            Then
                .left = mPoint2.x - 1
            End If
        End If
        If mLeftToRight Then
            If (Not extendAfter And periodNumber = mPoint2.x) Or _
                periodNumber = Int(mRegionRect.right) _
            Then
                .right = .left
            Else
                .right = periodNumber + 1
            End If
        Else
            If (Not extendBefore And periodNumber = mPoint1.x) Or _
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
            If (Not extendBefore And periodNumber = mPoint1.x) Then .left = .left - mThicknessOffsetX
            If (Not extendAfter And periodNumber = mPoint2.x) Or _
                periodNumber = Int(mRegionRect.right) _
            Then
                .right = .right + mThicknessOffsetX
            End If
        Else
            If (Not extendAfter And periodNumber = mPoint2.x) Then .left = .left - mThicknessOffsetX
            If (Not extendBefore And periodNumber = mPoint1.x) Or _
                periodNumber = Int(mRegionRect.right) _
            Then
                .right = .right + mThicknessOffsetX
            End If
        End If
        .isValid = True
    End With
End If

If Not arrowEndStyle = ArrowNone Then
    If rectOverlaps(mEndArrowRect, periodrect) Then
        rects(0) = rectUnion(rects(0), mEndArrowRect)
    End If
End If
If Not arrowStartStyle = ArrowNone Then
    If rectOverlaps(mStartArrowRect, periodrect) Then
        rects(0) = rectUnion(rects(0), mStartArrowRect)
    End If
End If
IGraphicObject_boundingRectanglesAt = rects

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_boundingRectanglesAt" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Private Property Get IGraphicObject_capabilities() As Long
'IGraphicObject_capabilities = capabilities.ComplexBounding Or _
'                            capabilities.PartialRedraw
IGraphicObject_capabilities = capabilities.PartialRedraw
End Property

Private Sub IGraphicObject_draw( _
                ByRef areas() As TRectangle, _
                ByVal hdc As Long)
Dim i As Long
Dim j As Long
Dim areasToRedraw() As TRectangle

Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_draw" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

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
IGraphicObject_noDraw = mHidden Or mBlank
End Property

Private Property Get IGraphicObject_periodNumber() As Long
IGraphicObject_periodNumber = -1
If isExtended Then Exit Sub
If mPoint1 Is Nothing Then Exit Sub
If Not mPoint1.isAssigned Then Exit Sub
IGraphicObject_periodNumber = Int(mPoint1.XLogical + 0.5)
End Property

Private Property Get IGraphicObject_scaleDependent() As Boolean
IGraphicObject_scaleDependent = fixedX Or fixedY
End Property

Private Property Get IGraphicObject_timestamp() As Date
IGraphicObject_timestamp = 0
End Property

Private Sub IGraphicObject_undraw( _
                ByRef area As TRectangle, _
                ByVal hdc As Long)
Dim failpoint As Long
On Error GoTo Err

undraw

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_undraw" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Private Property Let IGraphicObject_visible(ByVal value As Boolean)
mVisible = value
End Property


'================================================================================
' Properties
'================================================================================

Public Property Get arrowEndColor() As Long
If mIsSetArrowEndColor Then
    arrowEndColor = mArrowEndColor
ElseIf Not mLocalStyle Is Nothing Then
    arrowEndColor = mLocalStyle.arrowEndColor
Else
    arrowEndColor = mBaseStyle.arrowEndColor
End If
End Property

Public Property Let arrowEndColor(ByVal value As Long)
If mArrowEndColor = value And mIsSetArrowEndColor Then Exit Property
mArrowEndColor = value
mIsSetArrowEndColor = True
update
End Property

Public Property Get arrowEndFillColor() As Long
arrowEndFillColor = mBaseStyle.arrowEndFillColor
If Not mLocalStyle Is Nothing Then arrowEndFillColor = mLocalStyle.arrowEndFillColor
If mIsSetArrowEndFillColor Then arrowEndFillColor = mArrowEndFillColor
End Property

Public Property Let arrowEndFillColor(ByVal value As Long)
If mArrowEndFillColor = value And mIsSetArrowEndFillColor Then Exit Property
mArrowEndFillColor = value
mIsSetArrowEndFillColor = True
update
End Property

Public Property Get arrowEndFillStyle() As FillStyles
arrowEndFillStyle = mBaseStyle.arrowEndFillStyle
If Not mLocalStyle Is Nothing Then arrowEndFillStyle = mLocalStyle.arrowEndFillStyle
If mIsSetArrowEndFillStyle Then arrowEndFillStyle = mArrowEndFillStyle
End Property

Public Property Let arrowEndFillStyle(ByVal value As FillStyles)
If mArrowEndFillStyle = value And mIsSetArrowEndFillStyle Then Exit Property
mArrowEndFillStyle = value
mIsSetArrowEndFillStyle = True
update
End Property

Public Property Get arrowEndLength() As Long
arrowEndLength = mBaseStyle.arrowEndLength
If Not mLocalStyle Is Nothing Then arrowEndLength = mLocalStyle.arrowEndLength
If mIsSetArrowEndLength Then arrowEndLength = mArrowEndLength
End Property

Public Property Let arrowEndLength(ByVal value As Long)
If mArrowEndLength = value And mIsSetArrowEndLength Then Exit Property
undraw
mArrowEndLength = value
mIsSetArrowEndLength = True
update
End Property

Public Property Get arrowEndStyle() As ArrowStyles
arrowEndStyle = mBaseStyle.arrowEndStyle
If Not mLocalStyle Is Nothing Then arrowEndStyle = mLocalStyle.arrowEndStyle
If mIsSetArrowEndStyle Then arrowEndStyle = mArrowEndStyle
End Property

Public Property Let arrowEndStyle(ByVal value As ArrowStyles)
If mArrowEndStyle = value And mIsSetArrowEndStyle Then Exit Property
undraw
mArrowEndStyle = value
mIsSetArrowEndStyle = True
update
End Property

Public Property Get arrowEndWidth() As Long
arrowEndWidth = mBaseStyle.arrowEndWidth
If Not mLocalStyle Is Nothing Then arrowEndWidth = mLocalStyle.arrowEndWidth
If mIsSetArrowEndWidth Then arrowEndWidth = mArrowEndWidth
End Property

Public Property Let arrowEndWidth(ByVal value As Long)
If mArrowEndWidth = value And mIsSetArrowEndWidth Then Exit Property
undraw
mArrowEndWidth = value
mIsSetArrowEndWidth = True
update
End Property

Public Property Get arrowStartColor() As Long
arrowStartColor = mBaseStyle.arrowStartColor
If Not mLocalStyle Is Nothing Then arrowStartColor = mLocalStyle.arrowStartColor
If mIsSetArrowStartColor Then arrowStartColor = mArrowStartColor
End Property

Public Property Let arrowStartColor(ByVal value As Long)
If mArrowStartColor = value And mIsSetArrowStartColor Then Exit Property
mArrowStartColor = value
mIsSetArrowStartColor = True
update
End Property

Public Property Get arrowStartFillColor() As Long
arrowStartFillColor = mBaseStyle.arrowStartFillColor
If Not mLocalStyle Is Nothing Then arrowStartFillColor = mLocalStyle.arrowStartFillColor
If mIsSetArrowStartFillColor Then arrowStartFillColor = mArrowStartFillColor
End Property

Public Property Let arrowStartFillColor(ByVal value As Long)
If mArrowStartFillColor = value And mIsSetArrowStartFillColor Then Exit Property
mArrowStartFillColor = value
update
mIsSetArrowStartFillColor = True
End Property

Public Property Get arrowStartfillstyle() As FillStyles
arrowStartfillstyle = mBaseStyle.arrowStartfillstyle
If Not mLocalStyle Is Nothing Then arrowStartfillstyle = mLocalStyle.arrowStartfillstyle
If mIsSetArrowStartFillStyle Then arrowStartfillstyle = marrowStartfillstyle
End Property

Public Property Let arrowStartfillstyle(ByVal value As FillStyles)
If marrowStartfillstyle = value And mIsSetArrowStartFillStyle Then Exit Property
marrowStartfillstyle = value
mIsSetArrowStartFillStyle = True
update
End Property

Public Property Get arrowStartLength() As Long
arrowStartLength = mBaseStyle.arrowStartLength
If Not mLocalStyle Is Nothing Then arrowStartLength = mLocalStyle.arrowStartLength
If mIsSetArrowStartLength Then arrowStartLength = mArrowStartLength
End Property

Public Property Let arrowStartLength(ByVal value As Long)
If mArrowStartLength = value And mIsSetArrowStartLength Then Exit Property
undraw
mArrowStartLength = value
mIsSetArrowStartLength = True
update
End Property

Public Property Get arrowStartStyle() As ArrowStyles
arrowStartStyle = mBaseStyle.arrowStartStyle
If Not mLocalStyle Is Nothing Then arrowStartStyle = mLocalStyle.arrowStartStyle
If mIsSetArrowStartStyle Then arrowStartStyle = mArrowStartStyle
End Property

Public Property Let arrowStartStyle(ByVal value As ArrowStyles)
If mArrowStartStyle = value And mIsSetArrowStartStyle Then Exit Property
undraw
mArrowStartStyle = value
mIsSetArrowStartStyle = True
update
End Property

Public Property Get arrowStartWidth() As Long
arrowStartWidth = mBaseStyle.arrowStartWidth
If Not mLocalStyle Is Nothing Then arrowStartWidth = mLocalStyle.arrowStartWidth
If mIsSetArrowStartWidth Then arrowStartWidth = mArrowStartWidth
End Property

Public Property Let arrowStartWidth(ByVal value As Long)
If mArrowStartWidth = value And mIsSetArrowStartWidth Then Exit Property
undraw
mArrowStartWidth = value
mIsSetArrowStartWidth = True
update
End Property

Public Property Get Color() As Long
Color = mBaseStyle.Color
If Not mLocalStyle Is Nothing Then Color = mLocalStyle.Color
If mIsSetColor Then Color = mColor
End Property

Public Property Let Color(ByVal value As Long)
If mColor = value And mIsSetColor Then Exit Property
mColor = value
mIsSetColor = True
update
End Property

Public Property Get direction() As Quadrants
direction = mDirection
End Property

Friend Property Get extended() As Boolean
extended = mBaseStyle.extended
If Not mLocalStyle Is Nothing Then extended = mLocalStyle.extended
If mIsSetExtended Then extended = mExtended
End Property

Public Property Let extended(ByVal value As Boolean)
If mExtended = value And mIsSetExtended Then Exit Property
undraw
mExtended = value
mIsSetExtended = True
update
End Property

Public Property Get extendAfter() As Boolean
extendAfter = mBaseStyle.extendAfter
If Not mLocalStyle Is Nothing Then extendAfter = mLocalStyle.extendAfter
If mIsSetExtendAfter Then extendAfter = mExtendAfter
End Property

Public Property Let extendAfter(ByVal value As Boolean)
If mExtendAfter = value And mIsSetExtendAfter Then Exit Property
undraw
mExtendAfter = value
mIsSetExtendAfter = True
update
End Property

Public Property Get extendBefore() As Boolean
extendBefore = mBaseStyle.extendBefore
If Not mLocalStyle Is Nothing Then extendBefore = mLocalStyle.extendBefore
If mIsSetExtendBefore Then extendBefore = mExtendBefore
End Property

Public Property Let extendBefore(ByVal value As Boolean)
If mExtendBefore = value And mIsSetExtendBefore Then Exit Property
undraw
mExtendBefore = value
mIsSetExtendBefore = True
update
End Property

Friend Property Get fixedX() As Boolean
fixedX = mBaseStyle.fixedX
If Not mLocalStyle Is Nothing Then fixedX = mLocalStyle.fixedX
If mIsSetFixedX Then fixedX = mFixedX
End Property

Public Property Let fixedX(ByVal value As Boolean)
If mFixedX = value And mIsSetFixedX Then Exit Property
undraw
mFixedX = value
mIsSetFixedX = True
If mFixedX Then convertPositionX
update
End Property

Friend Property Get fixedY() As Boolean
fixedY = mBaseStyle.fixedY
If Not mLocalStyle Is Nothing Then fixedY = mLocalStyle.fixedY
If mIsSetFixedY Then fixedY = mFixedY
End Property

Public Property Let fixedY(ByVal value As Boolean)
If mFixedY = value And mIsSetFixedY Then Exit Property
undraw
mIsSetFixedY = True
mFixedY = value
If mFixedY Then convertPositionY
update
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Get hidden() As Boolean
hidden = mHidden
End Property

Public Property Let hidden(ByVal value As Boolean)
If mHidden = value Then Exit Property
If value Then undraw
mHidden = value
update
End Property

Public Property Get includeInAutoscale() As Boolean
includeInAutoscale = mBaseStyle.includeInAutoscale
If Not mLocalStyle Is Nothing Then includeInAutoscale = mLocalStyle.includeInAutoscale
If mIsSetIncludeInAutoscale Then includeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let includeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value And mIsSetIncludeInAutoscale Then Exit Property
If Not value Then undraw
mIncludeInAutoscale = value
mIsSetIncludeInAutoscale = True
update
End Property

Public Property Get inScope() As Boolean
inScope = mInScope
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

Public Property Get linestyle() As LineStyles
linestyle = mBaseStyle.linestyle
If Not mLocalStyle Is Nothing Then linestyle = mLocalStyle.linestyle
If mIsSetLineStyle Then linestyle = mLineStyle
End Property

Public Property Let linestyle(ByVal value As LineStyles)
If mLineStyle = value And mIsSetLineStyle Then Exit Property
undraw
mLineStyle = value
mIsSetLineStyle = True
update
End Property

Public Property Get localStyle() As BarStyle
Set localStyle = mLocalStyle.clone
End Property

Public Property Let localStyle( _
                ByVal value As BarStyle)

undraw
Set mLocalStyle = value.clone
calcEndArrowConstants
calcStartArrowConstants
calcCurrentBoundingRect
mEventProxy.objectChanged mHandle, mBoundingRect

End Property

Public Property Get point1() As Point
Set point1 = mPoint1
End Property

Public Property Let point1(ByVal value As Point)
If value Is Nothing Then
    undraw
    Set mPoint1 = Nothing
    mBlank = True
    Exit Property
ElseIf Not value.isAssigned Then
    undraw
    Set mPoint1 = Nothing
    mBlank = True
    Exit Property
ElseIf Not mPoint1 Is Nothing Then
    If value.equals(mPoint1) Then Exit Property
End If
undraw
Set mPoint1 = value
If mFixedX Then convertPositionX
If mFixedY Then convertPositionY

If mPoint2 Is Nothing Then Exit Property
If Not mPoint2.isAssigned Then Exit Property
setDirection
mBlank = False
update

End Property

Public Property Get point2() As Point
Set point2 = mPoint2
End Property

Public Property Let point2(ByVal value As Point)
If value Is Nothing Then
    undraw
    Set mPoint2 = Nothing
    mBlank = True
    Exit Property
ElseIf Not value.isAssigned Then
    undraw
    Set mPoint2 = Nothing
    mBlank = True
    Exit Property
ElseIf Not mPoint2 Is Nothing Then
    If value.equals(mPoint2) Then Exit Property
End If
undraw
Set mPoint2 = value
If mFixedX Then convertPositionX
If mFixedY Then convertPositionY

If mPoint1 Is Nothing Then Exit Property
If Not mPoint1.isAssigned Then Exit Property
setDirection
mBlank = False
update

End Property

Public Property Get thickness() As Long
thickness = mBaseStyle.thickness
If Not mLocalStyle Is Nothing Then thickness = mLocalStyle.thickness
If mIsSetThickness Then thickness = mThickness
End Property

Public Property Let thickness(ByVal value As Long)
If mThickness = value And mIsSetThickness Then Exit Property
undraw
mThickness = value
mIsSetThickness = True
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

Friend Sub initialise( _
                ByVal key As String, _
                ByVal eventProxy As ChartRegionEventProxy, _
                ByVal layer As Long, _
                ByVal canvas As canvas, _
                ByVal baseStyle As linestyle, _
                ByVal localStyle As linestyle)
                
mKey = key
Set mEventProxy = eventProxy
mLayer = layer
Set mCanvas = canvas
Set mBaseStyle = baseStyle
If Not localStyle Is Nothing Then Set mLocalStyle = localStyle.clone ' local style cannot be shared
calcCurrentBoundingRect
End Sub

Public Sub setPosition( _
                ByVal point1 As Point, _
                ByVal point2 As Point)
Dim failpoint As Long
On Error GoTo Err

If point1 Is Nothing Or point2 Is Nothing Then

    failpoint = 100

    undraw
    Set mPoint1 = point1
    Set mPoint2 = point2
    mBlank = True
    Exit Sub
ElseIf (Not point1.isAssigned) Then

    failpoint = 200

    undraw
    Set mPoint1 = Nothing
    Set mPoint2 = point2
    mBlank = True
    Exit Sub
ElseIf (Not point2.isAssigned) Then

    failpoint = 300

    undraw
    Set mPoint1 = point1
    Set mPoint2 = Nothing
    mBlank = True
    Exit Sub
ElseIf (Not mPoint1 Is Nothing) And (Not mPoint2 Is Nothing) Then

    failpoint = 400
    
    If point1.equals(mPoint1) And point2.equals(mPoint2) Then Exit Sub
End If

failpoint = 500

undraw
Set mPoint1 = point1
Set mPoint2 = point2

failpoint = 600

If mFixedX Then convertPositionX
If mFixedY Then convertPositionY

failpoint = 800

setDirection
mBlank = False

failpoint = 900

update

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "setPosition" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

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

mThicknessOffsetX = mGaugeX * thickness / 2
mThicknessOffsetY = mGaugeY * thickness / 2

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
'    mThicknessOffsetX = Abs(mGaugeX * (Thickness / mLineAngleSin) / 2)
'Else
'    mThicknessOffsetX = 0
'End If
'
'If mLineAngleCos <> 0 Then
'    mThicknessOffsetY = Abs(mGaugeY * (Thickness / mLineAngleCos) / 2)
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
            If extendBefore Then
                lineInt.startValue = MinusInfinityDouble
            Else
                lineInt.startValue = y1
            End If
            If extendAfter Then
                lineInt.endValue = PlusInfinityDouble
            Else
                lineInt.endValue = y2
            End If
        Else
            If extendAfter Then
                lineInt.startValue = MinusInfinityDouble
            Else
                lineInt.startValue = y2
            End If
            If extendBefore Then
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
            If extendBefore Then
                lineInt.startValue = MinusInfinityDouble
            Else
                lineInt.startValue = x1
            End If
            If extendAfter Then
                lineInt.endValue = PlusInfinityDouble
            Else
                lineInt.endValue = x2
            End If
        Else
            If extendAfter Then
                lineInt.startValue = MinusInfinityDouble
            Else
                lineInt.startValue = x2
            End If
            If extendBefore Then
                lineInt.endValue = PlusInfinityDouble
            Else
                lineInt.endValue = x1
            End If
        End If
        lineInt.isValid = True
        regionInt = rectGetXInterval(mRegionRect)
        rectSetXInterval mBoundingRect, intIntersection(regionInt, lineInt)
        'rectSetXInterval mBoundingRect, lineInt
        
        If Not .isValid Then
            .bottom = MinusInfinityDouble
            .top = PlusInfinityDouble
            .left = MinusInfinityDouble
            .right = PlusInfinityDouble
'            .isValid = True
            Exit Sub
        End If
            
        'If .right > Int(mRegionRect.right) Then .right = Int(mRegionRect.right)
        
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
        
        If thickness > 1 Then
            .right = .right + mThicknessOffsetX
            .left = .left - mThicknessOffsetX
        End If
        .isValid = True
    End If
'Debug.Print "Calc bounds : (" & .left & "," & (mA * .left + mB) & ")-(" & .right & "," & (mA * .right + mB) & ")"
End With

If arrowEndStyle <> ArrowNone Then mBoundingRect = rectUnion(mBoundingRect, mEndArrowRect)
If arrowStartStyle <> ArrowNone Then mBoundingRect = rectUnion(mBoundingRect, mStartArrowRect)

End Sub

Private Sub calcEndArrowConstants()
Dim hyp As Double
Dim halfAngle As Double
Dim x As Double
Dim y As Double

If arrowEndStyle = ArrowNone Then Exit Sub
If mBlank Then Exit Sub

hyp = Sqr(arrowEndLength * arrowEndLength + (arrowEndWidth * arrowEndWidth) / 4#)
halfAngle = Atn(arrowEndWidth / (2 * arrowEndLength))
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

If arrowStartStyle = ArrowNone Then Exit Sub
If mBlank Then Exit Sub

hyp = Sqr(arrowStartLength * arrowStartLength + arrowStartWidth * arrowStartWidth / 4)
halfAngle = Atn(arrowStartWidth / (2 * arrowStartLength))
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
If fixedX Then
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
If fixedY Then
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
Dim includeStartArrow As Boolean
Dim includeEndArrow As Boolean

Dim failpoint As Long
On Error GoTo Err

If mBlank Or mHidden Then Exit Sub
If Not mVisible Then
    Exit Sub
End If

failpoint = 100

If thickness = 1 Then
    mCanvas.setPenAttributes IIf(undraw, mCanvas.backColor, Color), _
                            thickness, _
                            linestyle, _
                            DrawModes.DrawModeCopyPen
Else
    mCanvas.setPenAttributes IIf(undraw, mCanvas.backColor, Color), _
                            thickness, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
End If

failpoint = 200

With mPoint1
    x1 = .XLogical
    y1 = .YLogical
End With

With mPoint2
    x2 = .XLogical
    y2 = .YLogical
End With

If mVertical <> VerticalNot Then

    failpoint = 300

    With area
        If mBottomToTop Then
            If extendAfter Then
                top = .top
            Else
                If .top >= y2 Then
                    top = y2
                    If .bottom <= y2 Then includeEndArrow = True
                Else
                    top = .top
                End If
            End If
            If extendBefore Then
                bottom = .bottom
            Else
                If .bottom <= y1 Then
                    bottom = y1
                    If .top >= y1 Then includeStartArrow = True
                Else
                    bottom = .bottom
                End If
            End If
        Else
            If extendAfter Then
                bottom = .bottom
            Else
                If .bottom <= y2 Then
                    bottom = y2
                    If .top >= y2 Then includeEndArrow = True
                Else
                    bottom = .bottom
                End If
            End If
            If extendBefore Then
                top = .top
            Else
                If .top >= y1 Then
                    top = y1
                    If .bottom <= y1 Then includeStartArrow = True
                Else
                    top = .top
                End If
            End If
        End If
    End With

    failpoint = 400

    mCanvas.drawLine mCanvas.newPoint(x1, bottom), mCanvas.newPoint(x1, top)

    failpoint = 500

    mCanvas.drawPoint mCanvas.newPoint(x1, top)
Else
    With area
        
        failpoint = 600

        left = .left
        right = .right
        
        If mLeftToRight And Not extendBefore Then
            If left < x1 Then left = x1
            If right < x1 Then right = x1
            If left <= x1 And right >= x1 Then includeStartArrow = True
        ElseIf Not mLeftToRight And Not extendAfter Then
            If left < x2 Then left = x2
            If right < x2 Then right = x2
            If left <= x2 And right >= x2 Then includeEndArrow = True
        End If
            
        failpoint = 700

        If mLeftToRight And Not extendAfter Then
            If left > x2 Then left = x2
            If right > x2 Then right = x2
            If left <= x2 And right >= x2 Then includeEndArrow = True
        ElseIf Not mLeftToRight And Not extendBefore Then
            If left > x1 Then left = x1
            If right > x1 Then right = x1
            If left <= x1 And right >= x1 Then includeStartArrow = True
        End If
    
        failpoint = 800

        If left > Int(mRegionRect.right) Then
            left = Int(mRegionRect.right)
        End If
        
        failpoint = 900

        If right > Int(mRegionRect.right) Then
            right = Int(mRegionRect.right)
        End If
    
    End With
    
    With mCanvas

        failpoint = 1000

        Debug.Assert left <= right
        If left <> right Then
            If thickness = 1 Then

                failpoint = 1100

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

                failpoint = 1200

                .drawLine mCanvas.newPoint(left, mA * left + mB), mCanvas.newPoint(right, mA * right + mB)
            Else

                failpoint = 1300

                If left = x1 Then
                    .drawPoint mCanvas.newPoint(left, y1)
                ElseIf left = x2 Then
                    .drawPoint mCanvas.newPoint(left, y2)
                End If
                
                slopeThicknessY = Abs(thickness / mLineAngleCos)
                                
                failpoint = 1400

                If left = x1 Or left = x2 Then
                    Set vertices(0) = .newPoint(left + .convertPixelsToLogicalX(thickness) * mLineAngleSin / 2, _
                                            mA * left + mB - .convertPixelsToLogicalY(thickness) * mLineAngleCos / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                Else
                    Set vertices(0) = .newPoint(left, _
                                            mA * left + mB - .convertPixelsToLogicalY(slopeThicknessY) / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                End If
                
                If left = x1 Or left = x2 Then
                    Set vertices(1) = .newPoint(left - .convertPixelsToLogicalX(thickness) * mLineAngleSin / 2, _
                                            mA * left + mB + .convertPixelsToLogicalY(thickness) * mLineAngleCos / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                Else
                    Set vertices(1) = .newPoint(left, _
                                            mA * left + mB + .convertPixelsToLogicalY(slopeThicknessY) / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                End If
                
                If right = x1 Or right = x2 Then
                    Set vertices(2) = .newPoint(right - .convertPixelsToLogicalX(thickness) * mLineAngleSin / 2, _
                                            mA * right + mB + .convertPixelsToLogicalY(thickness) * mLineAngleCos / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                Else
                    Set vertices(2) = .newPoint(right, _
                                            mA * right + mB + .convertPixelsToLogicalY(slopeThicknessY) / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                End If
                
                If right = x1 Or right = x2 Then
                    Set vertices(3) = .newPoint(right + .convertPixelsToLogicalX(thickness) * mLineAngleSin / 2, _
                                            mA * right + mB - .convertPixelsToLogicalY(thickness) * mLineAngleCos / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                Else
                    Set vertices(3) = .newPoint(right, _
                                            mA * right + mB - .convertPixelsToLogicalY(slopeThicknessY) / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                End If

                failpoint = 1500

                .setPenAttributes Color, _
                                1, _
                                LineStyles.LineSolid, _
                                DrawModes.DrawModeCopyPen

                failpoint = 1600

                .setBrushAttributes Color, _
                                FillStyles.FillSolid
                

                failpoint = 1700

                .drawPolygon vertices
                
                failpoint = 1800

                .setPenAttributes Color, _
                                thickness, _
                                LineStyles.LineSolid, _
                                DrawModes.DrawModeCopyPen

                failpoint = 1900

                If right = x1 Then
                    .drawPoint mCanvas.newPoint(right, y1)
                ElseIf right = x2 Then
                    .drawPoint mCanvas.newPoint(right, y2)
                End If
                
                Debug.Assert IIf(mLeftToRight, right <= x2, True)
                Debug.Assert IIf(mLeftToRight, True, right <= x1)
            End If
        End If
    End With
End If

failpoint = 2000

If arrowEndStyle <> ArrowNone And includeEndArrow Then
    drawArrowEnd area, undraw
End If

failpoint = 2100

If arrowStartStyle <> ArrowNone And includeStartArrow Then
    drawArrowStart area, undraw
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "draw" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


    
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

Select Case arrowEndStyle
Case ArrowNone

Case ArrowSingleOpen
    mCanvas.setPenAttributes IIf(undraw, mCanvas.backColor, arrowEndColor), _
                            thickness, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
    
    mCanvas.drawLine mCanvas.newPoint(x, y), mCanvas.newPoint(x + mEndArrowX1Offset, y + mEndArrowY1Offset)
    mCanvas.drawLine mCanvas.newPoint(x, y), mCanvas.newPoint(x + mEndArrowX2Offset, y + mEndArrowY2Offset)
Case ArrowDoubleOpen

Case ArrowClosed
    mCanvas.setPenAttributes IIf(undraw, mCanvas.backColor, arrowEndColor), _
                            1, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
    
    With mCanvas
        .setBrushAttributes IIf(undraw, .backColor, arrowEndFillColor), arrowEndFillStyle
        
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

Select Case arrowStartStyle
Case ArrowNone

Case ArrowSingleOpen
    mCanvas.setPenAttributes IIf(undraw, mCanvas.backColor, arrowStartColor), _
                            thickness, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
    
    mCanvas.drawLine mCanvas.newPoint(x, y), mCanvas.newPoint(x + mStartArrowX1Offset, y + mStartArrowY1Offset)
    mCanvas.drawLine mCanvas.newPoint(x, y), mCanvas.newPoint(x + mStartArrowX2Offset, y + mStartArrowY2Offset)
Case ArrowDoubleOpen

Case ArrowClosed
    mCanvas.setPenAttributes IIf(undraw, mCanvas.backColor, arrowStartColor), _
                            1, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
    
    With mCanvas
        .setBrushAttributes IIf(undraw, .backColor, arrowStartFillColor), arrowStartfillstyle
        
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
isExtended = extended Or fixedX
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
'draw rectIntersection(mBoundingRect, mRegionRect), True
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


