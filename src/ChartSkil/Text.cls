VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Text"
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

Private mCanvas As canvas
Private mLayer As Long
Private mHandle As Long

Private mInScope As Boolean
Private mVisible As Boolean

Private mEventProxy As ChartRegionEventProxy

Private mRawBoundingRect As TRectangle
Private mBoundingRect As TRectangle ' the rectangle that surrounds the box (if
                                ' any - including the thickness of the lines) or
                                ' the text if no box
Private mTextRect As TRectangle ' the rectangle containing the text
Private mBoxRect As TRectangle  ' the rectangle used to draw the box

Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

Private mBlank As Boolean
Private mDrawn As Boolean

Private mText As String
Private mFont As StdFont
Private mColor As Long
Private mBox As Boolean
Private mBoxColor As Long
Private mBoxStyle As LineStyles
Private mBoxThickness As Long
Private mBoxFillColor As Long
Private mBoxFillStyle As FillStyles
Private mPaddingX  As Double
Private mPaddingY  As Double

Private mPosition As Point
Private mOffset As Dimension
Private mAlign As TextAlignModes
Private mKeepInView As Boolean

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mBlank = True

Set mPosition = New Point

mRawBoundingRect.isValid = False
mBoundingRect.isValid = False
mTextRect.isValid = False
mBoxRect.isValid = False
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
Static prevGaugeX As Double
Static prevGaugeY As Double
Static prevRect As TRectangle
Dim canvasChanges As Long

canvasChanges = mCanvas.compareTo(prevGaugeX, prevGaugeY, prevRect)
If (canvasChanges And CanvasComparisonCodes.GaugeChangedX) Or _
    (canvasChanges And CanvasComparisonCodes.GaugeChangedY) _
Then
    calcRawBoundingRect
End If

If (canvasChanges And CanvasComparisonCodes.GaugeChangedX) Or _
    (canvasChanges And CanvasComparisonCodes.GaugeChangedY) Or _
    (mFixedX And (canvasChanges And CanvasComparisonCodes.BoundsChangedX)) Or _
    (mFixedY And (canvasChanges And CanvasComparisonCodes.BoundsChangedY)) _
Then
    calcCurrentBoundingRect
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
IGraphicObject_capabilities = capabilities.blockUndraw
End Property

Private Sub IGraphicObject_draw( _
                ByRef areas() As TRectangle, _
                ByVal hdc As Long)
'If mText = "Started here" Then Stop
draw
End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = mExtended Or mFixedX
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
IGraphicObject_noDraw = mBlank
End Property

Private Property Get IGraphicObject_periodNumber() As Long
If mExtended Or Not mBoundingRect.isValid Then
    IGraphicObject_periodNumber = -1
Else
    Select Case mAlign
    Case AlignTopLeft, AlignCentreLeft, AlignBottomLeft, _
            AlignBoxTopLeft, AlignBoxCentreLeft, AlignBoxBottomLeft
        IGraphicObject_periodNumber = Int(mBoundingRect.left + 0.5)
    Case AlignTopCentre, AlignCentreCentre, AlignBottomCentre, _
            AlignBoxTopCentre, AlignBoxCentreCentre, AlignBoxBottomCentre
        IGraphicObject_periodNumber = Int((mBoundingRect.left + mBoundingRect.right) / 2 + 0.5)
    Case AlignTopRight, AlignCentreRight, AlignBottomRight, _
            AlignBoxTopRight, AlignBoxCentreRight, AlignBoxBottomRight
        IGraphicObject_periodNumber = Int(mBoundingRect.right + 0.5)
    End Select
End If
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

Public Property Get align() As TextAlignModes
align = mAlign
End Property

Public Property Let align(ByVal newValue As TextAlignModes)
If mAlign = newValue Then Exit Property
undraw
mAlign = newValue
update
End Property

Friend Property Get box() As Boolean
box = mBox
End Property

Public Property Let box(ByVal newValue As Boolean)
If mBox = newValue Then Exit Property
undraw
mBox = newValue
update
End Property

Public Property Get boxColor() As Long
boxColor = mBoxColor
End Property

Public Property Let boxColor(ByVal newValue As Long)
If mBoxColor = newValue Then Exit Property
undraw
mBoxColor = newValue
update
End Property

Friend Property Get boxFillColor() As Long
boxFillColor = mBoxFillColor
End Property

Public Property Let boxFillColor(ByVal newValue As Long)
If mBoxFillColor = newValue Then Exit Property
undraw
mBoxFillColor = newValue
update
End Property

Public Property Get boxFillStyle() As FillStyles
boxFillStyle = mBoxFillStyle
End Property

Public Property Let boxFillStyle(ByVal newValue As FillStyles)
If mBoxFillStyle = newValue Then Exit Property
undraw
mBoxFillStyle = newValue
update
End Property

Public Property Get boxStyle() As LineStyles
boxStyle = mBoxStyle
End Property

Public Property Let boxStyle(ByVal newValue As LineStyles)
If mBoxStyle = newValue Then Exit Property
undraw
mBoxStyle = newValue
update
End Property

Public Property Get boxThickness() As Long
boxThickness = mBoxThickness
End Property

Public Property Let boxThickness(ByVal newValue As Long)
If mBoxThickness = newValue Then Exit Property
undraw
mBoxThickness = newValue
update
End Property

Friend Property Let canvas(ByVal value As canvas)
Set mCanvas = value
End Property

Public Property Get Color() As Long
Color = mColor
End Property

Public Property Let Color(ByVal newValue As Long)
If mColor = newValue Then Exit Property
undraw
mColor = newValue
update
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

Public Property Get font() As StdFont
Set font = mFont
End Property

Public Property Let font(ByVal value As StdFont)
If value Is Nothing Then Exit Property
undraw
Set mFont = value
update
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Friend Property Get includeInAutoscale() As Boolean
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

Public Property Get keepInView() As Boolean
keepInView = mKeepInView
End Property

Public Property Let keepInView(ByVal newValue As Boolean)
If mKeepInView = newValue Then Exit Property
undraw
mKeepInView = newValue
update
End Property

Friend Property Let layer(ByVal value As Long)
If mLayer = value Then Exit Property
undraw
mLayer = value
update
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Let paddingX(ByVal newValue As Double)
If mPaddingX = newValue Then Exit Property
undraw
mPaddingX = newValue
update
End Property

Public Property Get paddingX() As Double
paddingX = mPaddingX
End Property

Public Property Let paddingY(ByVal newValue As Double)
If mPaddingY = newValue Then Exit Property
undraw
mPaddingY = newValue
update
End Property

Public Property Get paddingY() As Double
paddingY = mPaddingY
End Property

Public Property Get position() As Point
Set position = mPosition
End Property

Public Property Let position(ByVal value As Point)
If value Is Nothing Or Not value.isAssigned Then Exit Property
undraw
Set mPosition = value
convertPosition
update
End Property

Public Property Get offset() As Dimension
Set offset = mOffset
End Property

Public Property Let offset(ByVal value As Dimension)
If Not value.isAssigned Then Exit Property
undraw
Set mOffset = value
update
End Property

Public Property Let text(value As String)
If value = mText Then Exit Property
undraw  ' remove current text
mText = value
mBlank = (value = "")

update
End Property

Public Property Get visible() As Boolean
visible = mVisible
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub draw()
If mBlank Then Exit Sub
If Not mVisible Then Exit Sub
If Not mPosition.isAssigned Then Exit Sub
mDrawn = True

With mCanvas
    If mBox Then
        .setPenAttributes mBoxColor, _
                        mBoxThickness, _
                        mBoxStyle, _
                        DrawModes.DrawModeCopyPen
        .setBrushAttributes mBoxFillColor, mBoxFillStyle
        .drawRectangle mBoxRect.left, mBoxRect.bottom, mBoxRect.right, mBoxRect.top
    End If
    
    .setTextAttributes mColor, mFont, True
'    If mText = "Volume" Then Stop
    .drawText mText, mTextRect.left, mTextRect.top
End With
    
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcRawBoundingRect()

' NB: calculates the raw dimensions of the bounding box in twips, because
' the scale may not have been finalised when this is called, and is anyway
' subject to change
If mText = "" Then Exit Sub

mCanvas.setTextAttributes mColor, mFont, True
With mRawBoundingRect
    .left = 0
    .right = mCanvas.TextWidth(mText)
    .bottom = 0
    .top = mCanvas.TextHeight(mText)
    .isValid = True
End With

End Sub

Private Sub calcCurrentBoundingRect()
Dim xshift As Double
Dim yshift As Double
Dim xOffset As Double
Dim yOffset As Double
Dim XIncrement As Double
Dim YIncrement As Double
Dim XLineThickness As Double
Dim YLineThickness As Double

'If mText = "Volume" Then Stop
'If mHandle = 520 Then Stop

If Not mPosition.isAssigned Then
    mBoundingRect.isValid = False
    Exit Sub
End If

With mTextRect
    .left = mPosition.XLogical
    .right = .left + mRawBoundingRect.right
    .top = mPosition.YLogical
    .bottom = .top - mRawBoundingRect.top
    
    ' now add the offset if any
    
    If Not mOffset Is Nothing Then
        xOffset = mOffset.XLogical
        yOffset = mOffset.YLogical
        .left = .left + xOffset
        .right = .right + xOffset
        .bottom = .bottom + yOffset
        .top = .top + yOffset
    End If
End With

If mBox Then
    ' make the inside of the box 1mm wider and taller than the text, plus the thickness of the lines
    
    XIncrement = mCanvas.convertDistanceToLogicalX(mPaddingX / 10)
    YIncrement = mCanvas.convertDistanceToLogicalY((mPaddingY - 0.5) / 10)
    
    If mBoxThickness > 1 Then
        XLineThickness = mCanvas.convertPixelsToLogicalX(mBoxThickness - 1)
        YLineThickness = mCanvas.convertPixelsToLogicalY(mBoxThickness - 1)
        If mBoxStyle = LineInsideSolid Then
            XIncrement = XIncrement + XLineThickness
            YIncrement = YIncrement + YLineThickness
        Else
            XIncrement = XIncrement + XLineThickness / 2
            YIncrement = YIncrement + YLineThickness / 2
        End If
    End If
    
    With mBoxRect
        .left = mTextRect.left - XIncrement
        .right = mTextRect.right + XIncrement
        .bottom = mTextRect.bottom - YIncrement
        .top = mTextRect.top + YIncrement
    End With
    
    If mBoxThickness > 1 And mBoxStyle <> LineInsideSolid Then
        With mBoundingRect
            .left = mBoxRect.left - XLineThickness / 2
            .right = mBoxRect.right + XLineThickness / 2
            .bottom = mBoxRect.bottom - YLineThickness / 2
            .top = mBoxRect.top + YLineThickness / 2
        End With
    Else
        mBoundingRect = mBoxRect
    End If
    ' now adjust the rectangle to take account of alignment

    Select Case mAlign
    Case AlignTopLeft
        xshift = 0
        yshift = 0
    Case AlignCentreLeft
        xshift = 0
        yshift = (mTextRect.top - mTextRect.bottom) / 2
    Case AlignBottomLeft
        xshift = 0
        yshift = mTextRect.top - mTextRect.bottom
    Case AlignTopCentre
        xshift = (mTextRect.left - mTextRect.right) / 2
        yshift = 0
    Case AlignCentreCentre
        xshift = (mTextRect.left - mTextRect.right) / 2
        yshift = (mTextRect.top - mTextRect.bottom) / 2
    Case AlignBottomCentre
        xshift = (mTextRect.left - mTextRect.right) / 2
        yshift = mTextRect.top - mTextRect.bottom
    Case AlignTopRight
        xshift = mTextRect.left - mTextRect.right
        yshift = 0
    Case AlignCentreRight
        xshift = mTextRect.left - mTextRect.right
        yshift = (mTextRect.top - mTextRect.bottom) / 2
    Case AlignBottomRight
        xshift = mTextRect.left - mTextRect.right
        yshift = mTextRect.top - mTextRect.bottom
    Case AlignBoxTopLeft
        xshift = mTextRect.left - mBoundingRect.left
        yshift = mTextRect.top - mBoundingRect.top
    Case AlignBoxCentreLeft
        xshift = mTextRect.left - mBoundingRect.left
        yshift = (mTextRect.top - mTextRect.bottom) / 2
    Case AlignBoxBottomLeft
        xshift = mTextRect.left - mBoundingRect.left
        yshift = mTextRect.top - mTextRect.bottom + mTextRect.top - mBoundingRect.top
    Case AlignBoxTopCentre
        xshift = (mTextRect.left - mTextRect.right) / 2
        yshift = mTextRect.top - mBoundingRect.top
    Case AlignBoxCentreCentre
        xshift = (mTextRect.left - mTextRect.right) / 2
        yshift = (mTextRect.top - mTextRect.bottom) / 2
    Case AlignBoxBottomCentre
        xshift = (mTextRect.left - mTextRect.right) / 2
        yshift = mTextRect.top - mTextRect.bottom + mTextRect.top - mBoundingRect.top
    Case AlignBoxTopRight
        xshift = mTextRect.left - mTextRect.right - (mBoundingRect.right - mTextRect.right)
        yshift = mTextRect.top - mBoundingRect.top
    Case AlignBoxCentreRight
        xshift = mTextRect.left - mTextRect.right - (mBoundingRect.right - mTextRect.right)
        yshift = (mTextRect.top - mTextRect.bottom) / 2
    Case AlignBoxBottomRight
        xshift = mTextRect.left - mTextRect.right - (mBoundingRect.right - mTextRect.right)
        yshift = mTextRect.top - mTextRect.bottom + mTextRect.top - mBoundingRect.top
    End Select

    With mTextRect
        .left = .left + xshift
        .right = .right + xshift
        .bottom = .bottom + yshift
        .top = .top + yshift
    End With

    With mBoxRect
        .left = .left + xshift
        .right = .right + xshift
        .bottom = .bottom + yshift
        .top = .top + yshift
    End With

    With mBoundingRect
        .left = .left + xshift
        .right = .right + xshift
        .bottom = .bottom + yshift
        .top = .top + yshift
    End With

Else
    ' now adjust the rectangle to take account of alignment
    
    With mTextRect
        Select Case mAlign
        Case AlignTopLeft, AlignBoxTopLeft
            xshift = 0
            yshift = 0
        Case AlignCentreLeft, AlignBoxCentreLeft
            xshift = 0
            yshift = (.top - .bottom) / 2
        Case AlignBottomLeft, AlignBoxBottomLeft
            xshift = 0
            yshift = .top - .bottom
        Case AlignTopCentre, AlignBoxTopCentre
            xshift = (.left - .right) / 2
            yshift = 0
        Case AlignCentreCentre, AlignBoxCentreCentre
            xshift = (.left - .right) / 2
            yshift = (.top - .bottom) / 2
        Case AlignBottomCentre, AlignBoxBottomCentre
            xshift = (.left - .right) / 2
            yshift = .top - .bottom
        Case AlignTopRight, AlignBoxTopRight
            xshift = .left - .right
            yshift = 0
        Case AlignCentreRight, AlignBoxCentreRight
            xshift = .left - .right
            yshift = (.top - .bottom) / 2
        Case AlignBottomRight, AlignBoxBottomRight
            xshift = .left - .right
            yshift = .top - .bottom
        End Select
    
        .left = .left + xshift
        .right = .right + xshift
        .bottom = .bottom + yshift
        .top = .top + yshift
    End With
    
    mBoxRect = mTextRect
    mBoundingRect = mTextRect
End If

If mKeepInView Then
    ' now adjust the rectangles to take account of keepInView
End If

mTextRect.isValid = True
mBoxRect.isValid = True
mBoundingRect.isValid = True

End Sub

Private Sub convertPosition()
convertPositionX
convertPositionY
End Sub

Private Sub convertPositionX()
If mPosition Is Nothing Then Exit Sub
If mFixedX Then
    If mPosition.CoordinateSystemX = CoordsLogical Then
        mPosition.CoordinateSystemX = CoordsRelative
    End If
Else
    If mPosition.CoordinateSystemX <> CoordsLogical Then
        mPosition.CoordinateSystemX = CoordsLogical
    End If
End If
End Sub

Private Sub convertPositionY()
If mPosition Is Nothing Then Exit Sub
If mFixedY Then
    If mPosition.CoordinateSystemY = CoordsLogical Then
        mPosition.CoordinateSystemY = CoordsRelative
    End If
Else
    If mPosition.CoordinateSystemY <> CoordsLogical Then
        mPosition.CoordinateSystemY = CoordsLogical
    End If
End If
End Sub

Private Sub undraw()

If mBlank Then Exit Sub
If Not mVisible Then Exit Sub
If Not mPosition.isAssigned Then Exit Sub
mDrawn = True

With mCanvas
    .setPenAttributes .backColor, _
                    1, _
                    LineStyles.LineInsideSolid, _
                    DrawModes.DrawModeCopyPen
    .setBrushAttributes .backColor, FillStyles.FillSolid
    .drawRectangle mBoundingRect.left, mBoundingRect.bottom, mBoundingRect.right, mBoundingRect.top
End With
mEventProxy.objectUndrawn mHandle, mBoundingRect

End Sub

Private Sub update()
If mHandle = 0 Then
    ' haven't yet been added to chart region
    Exit Sub
End If
calcRawBoundingRect
calcCurrentBoundingRect
mEventProxy.objectChanged mHandle, mBoundingRect
'If mIncludeInAutoscale And mBoundingRect.isvalid Then
'    If mBoundingRect.top >= mSurface.ScaleTop Or mBoundingRect.bottom <= (mSurface.ScaleTop + mSurface.ScaleHeight) Then
'        mEventProxy.boundsExceeded
'    End If
'End If
End Sub
