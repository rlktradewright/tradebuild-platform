VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bar"
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
' Constants
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mCanvas As canvas
Private mLayer As Long
Private mHandle As Long

Private mLocalStyle As BarStyle
Private mBaseStyle As BarStyle

Private mIsSetTailThickness As Boolean
Private mIsSetOutlineThickness As Boolean
Private mIsSetBarColor As Boolean
Private mIsSetUpColor As Boolean
Private mIsSetDownColor As Boolean
Private mIsSetDisplayMode As Boolean
Private mIsSetSolidUpBody As Boolean
Private mIsSetBarThickness As Boolean
Private mIsSetBarWidth As Boolean

Private mTailThickness As Long
Private mOutlineThickness As Long
Private mBarColor As Long
Private mUpColor As Long
Private mDownColor As Long
Private mDisplayMode As BarDisplayModes
Private mSolidUpBody As Boolean
Private mBarThickness As Long
Private mBarWidth As Single

Private mNoDraw As Boolean

Private mInScope As Boolean
Private mVisible As Boolean

Private mEventProxy As ChartRegionEventProxy

Private mBoundingRect As TRectangle

'Private mPeriodNumber As Long
Private mX As Double

Private mKey As String

Private mOpen As Double
Private mHigh As Double
Private mLow As Double
Private mClose As Double
Private mBlank As Boolean   ' indicates that there were no ticks during the bar

Private mBarSemiWidth As Single ' half the width of the bar - this actually depends
                                ' on the bar display mode, rather than being simply
                                ' half of BarWidth

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mOpen = MinusInfinityDouble
mLow = PlusInfinityDouble
mHigh = MinusInfinityDouble
mClose = MinusInfinityDouble
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
Dim i As Long
For i = 0 To UBound(areas)
    draw areas(i)
Next
End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = False
End Property

Private Sub IGraphicObject_finish()
Set mCanvas = Nothing
End Sub

Private Property Get IGraphicObject_gaugeDependent() As Boolean
IGraphicObject_gaugeDependent = False
End Property

Private Property Get IGraphicObject_handle() As Long
IGraphicObject_handle = mHandle
End Property

Private Property Let IGraphicObject_handle(ByVal value As Long)
mHandle = value
End Property

Private Property Get IGraphicObject_includeInAutoscale() As Boolean
IGraphicObject_includeInAutoscale = True
End Property

Private Property Let IGraphicObject_inScope(ByVal value As Boolean)
mInScope = value
End Property

Private Property Get IGraphicObject_layer() As Long
IGraphicObject_layer = layer
End Property

Private Property Get IGraphicObject_noDraw() As Boolean
IGraphicObject_noDraw = mBlank Or mNoDraw
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
undraw area
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

Public Property Let barColor(ByVal value As Long)
mBarColor = value
mIsSetBarColor = True
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get barColor() As Long
barColor = mBaseStyle.barColor
If Not mLocalStyle Is Nothing Then barColor = mBaseStyle.barColor
If mIsSetBarColor Then barColor = mBarColor
End Property

Public Property Let barThickness(ByVal value As Long)
mBarThickness = value
mIsSetBarThickness = True
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get barThickness() As Long
barThickness = mBaseStyle.barThickness
If Not mLocalStyle Is Nothing Then barThickness = mBaseStyle.barThickness
If mIsSetBarThickness Then barThickness = mBarThickness
End Property

Public Property Let barWidth(ByVal value As Single)
mBarWidth = value
mIsSetBarWidth = True
setBarSemiWidth
calcBoundingRect
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get barWidth() As Single
barWidth = mBaseStyle.barWidth
If Not mLocalStyle Is Nothing Then barWidth = mBaseStyle.barWidth
If mIsSetBarWidth Then barWidth = mBarWidth
End Property

Public Property Get blank() As Boolean
blank = mBlank
End Property

Public Property Get closePrice() As Double
closePrice = mClose
End Property

Public Property Let displayMode(ByVal value As BarDisplayModes)
mDisplayMode = value
mIsSetDisplayMode = True
setBarSemiWidth
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get displayMode() As BarDisplayModes
displayMode = mBaseStyle.displayMode
If Not mLocalStyle Is Nothing Then displayMode = mBaseStyle.displayMode
If mIsSetDisplayMode Then displayMode = mDisplayMode
End Property

Public Property Get downBar() As Boolean
downBar = (mClose <= mOpen)
End Property

Public Property Let downColor(ByVal value As Long)
mDownColor = value
mIsSetDownColor = True
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get downColor() As Long
downColor = mBaseStyle.downColor
If Not mLocalStyle Is Nothing Then downColor = mBaseStyle.downColor
If mIsSetDownColor Then downColor = mDownColor
End Property

Friend Property Let eventProxy(ByVal value As ChartRegionEventProxy)
Set mEventProxy = value
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Get highPrice() As Double
highPrice = mHigh
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

Public Property Get localStyle() As BarStyle
Set localStyle = mLocalStyle.clone
End Property

Public Property Let localStyle( _
                ByVal value As BarStyle)
Dim somethingChanged As Boolean

mNoDraw = True

Set mLocalStyle = value.clone
setBarSemiWidth
calcBoundingRect
mEventProxy.objectChanged mHandle, mBoundingRect

mNoDraw = False
End Property

Public Property Get lowPrice() As Double
lowPrice = mLow
End Property

Public Property Get openPrice() As Double
openPrice = mOpen
End Property

Public Property Let outlineThickness(ByVal value As Long)
mOutlineThickness = value
mIsSetOutlineThickness = True
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get outlineThickness() As Long
outlineThickness = mBaseStyle.outlineThickness
If Not mLocalStyle Is Nothing Then outlineThickness = mBaseStyle.outlineThickness
If mIsSetOutlineThickness Then outlineThickness = mOutlineThickness
End Property

Public Property Let solidUpBody(ByVal value As Boolean)
mSolidUpBody = value
mIsSetSolidUpBody = True
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get solidUpBody() As Boolean
solidUpBody = mBaseStyle.solidUpBody
If Not mLocalStyle Is Nothing Then solidUpBody = mBaseStyle.solidUpBody
If mIsSetSolidUpBody Then solidUpBody = mSolidUpBody
End Property

Public Property Let tailThickness(ByVal value As Long)
mTailThickness = value
mIsSetTailThickness = True
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get tailThickness() As Long
tailThickness = mBaseStyle.tailThickness
If Not mLocalStyle Is Nothing Then tailThickness = mBaseStyle.tailThickness
If mIsSetTailThickness Then tailThickness = mTailThickness
End Property

Public Property Get upBar() As Boolean
upBar = (mClose >= mOpen)
End Property

Public Property Let upColor(ByVal value As Long)
mUpColor = value
mIsSetUpColor = True
mEventProxy.objectChanged mHandle, mBoundingRect
End Property

Public Property Get upColor() As Long
upColor = mBaseStyle.upColor
If Not mLocalStyle Is Nothing Then upColor = mBaseStyle.upColor
If mIsSetUpColor Then upColor = mUpColor
End Property

Public Property Get visible() As Boolean
visible = mVisible
End Property

Friend Property Let x(ByVal value As Double)
mX = value
calcBoundingRect
End Property

Public Property Get x() As Double
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
                ByVal baseStyle As BarStyle, _
                ByVal localStyle As BarStyle)
mKey = key
Set mEventProxy = eventProxy
mLayer = layer
Set mCanvas = canvas
Set mBaseStyle = baseStyle
If Not localStyle Is Nothing Then Set mLocalStyle = localStyle.clone ' local style cannot be shared
setBarSemiWidth
calcBoundingRect
End Sub

Public Sub resetStyle()
' reset the style to the initial style (ie the style for the bar series) with no
' local modifications
Set mLocalStyle = Nothing

mIsSetTailThickness = False
mIsSetOutlineThickness = False
mIsSetBarColor = False
mIsSetUpColor = False
mIsSetDownColor = False
mIsSetDisplayMode = False
mIsSetSolidUpBody = False
mIsSetBarThickness = False
mIsSetBarWidth = False

setBarSemiWidth
calcBoundingRect
mEventProxy.objectChanged mHandle, mBoundingRect
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcBoundingRect()
mBoundingRect.left = mX - mBarSemiWidth
mBoundingRect.right = mX + mBarSemiWidth
mBoundingRect.bottom = mLow
mBoundingRect.top = mHigh
End Sub

Private Sub draw(ByRef area As TRectangle)
If Not area.isValid Then
    drawBar IIf(upBar, upColor, downColor), _
            IIf(barColor <> -1, barColor, IIf(upBar, upColor, downColor))
Else
    drawPartialBar area, _
                    IIf(upBar, upColor, downColor), _
                    IIf(barColor <> -1, barColor, IIf(upBar, upColor, downColor))
End If
End Sub


Private Sub drawBar( _
                ByVal Color As Long, _
                ByVal outlineColor As Long)
Dim barThicknessScaleX As Double
Dim barThicknessScaleY As Double

If Not mVisible Then Exit Sub
If blank Then Exit Sub

If displayMode = BarDisplayModeCandlestick Then
    If upBar Then
        drawBody mOpen, mClose, Color, outlineColor
    Else
        drawBody mClose, mOpen, Color, outlineColor
    End If
    
    If upBar Then
        If mLow < mOpen Then
            drawTail mLow, mOpen, Color
        End If
        If mHigh > mClose Then
            drawTail mClose, mHigh, Color
        End If
    Else
        If mHigh > mOpen Then
            drawTail mOpen, mHigh, Color
        End If
        If mLow < mClose Then
            drawTail mLow, mClose, Color
        End If
    End If
Else
    If Color <> -1 Then
        mCanvas.setBrushAttributes Color, FillStyles.FillSolid
    Else
        mCanvas.setBrushAttributes outlineColor, FillStyles.FillSolid
    End If
    If barThickness = 1 Then
        Dim onePixel As Double
        onePixel = mCanvas.convertPixelsToLogicalX(1)
        mCanvas.setPenAttributes Color, _
                                1, _
                                LineStyles.LineSolid, _
                                DrawModes.DrawModeCopyPen
        mCanvas.drawLine mX - mBarSemiWidth, _
                        mOpen, _
                        mX, _
                        mOpen
        mCanvas.drawLine mX, _
                        mLow, _
                        mX, _
                        mHigh
        mCanvas.drawLine mX, _
                        mClose, _
                        mX + mBarSemiWidth + onePixel, _
                        mClose
    Else
        mCanvas.setPenAttributes Color, _
                                1, _
                                LineStyles.LineInsideSolid, _
                                DrawModes.DrawModeCopyPen
        barThicknessScaleX = mCanvas.convertPixelsToLogicalX(barThickness - 1)
        barThicknessScaleY = mCanvas.convertPixelsToLogicalY(barThickness - 1)
        
        If (mOpen - barThicknessScaleY / 2) < mLow Then
            mCanvas.drawRectangle mX - mBarSemiWidth, _
                                    mLow, _
                                    mX + barThicknessScaleX / 2, _
                                    mLow + barThicknessScaleY
        ElseIf (mOpen + barThicknessScaleY / 2) > mHigh Then
            mCanvas.drawRectangle mX - mBarSemiWidth, _
                                    mHigh - barThicknessScaleY, _
                                    mX + barThicknessScaleX / 2, _
                                    mHigh
        Else
            mCanvas.drawRectangle mX - mBarSemiWidth, _
                                    mOpen - barThicknessScaleY / 2, _
                                    mX + barThicknessScaleX / 2, _
                                    mOpen + barThicknessScaleY / 2
        End If
        
        mCanvas.drawRectangle mX - barThicknessScaleX / 2, _
                                mLow, _
                                mX + barThicknessScaleX / 2, _
                                mHigh
        
        If (mClose - barThicknessScaleY / 2) < mLow Then
            mCanvas.drawRectangle mX - barThicknessScaleX / 2, _
                                    mLow, _
                                    mX + mBarSemiWidth, _
                                    mLow + barThicknessScaleY
        ElseIf (mClose + barThicknessScaleY / 2) > mHigh Then
            mCanvas.drawRectangle mX - barThicknessScaleX / 2, _
                                    mHigh - barThicknessScaleY, _
                                    mX + mBarSemiWidth, _
                                    mHigh
        Else
            mCanvas.drawRectangle mX - barThicknessScaleX / 2, _
                                    mClose - barThicknessScaleY / 2, _
                                    mX + mBarSemiWidth, _
                                    mClose + barThicknessScaleY / 2
        End If
    End If
End If

End Sub

Private Sub drawBody( _
                ByVal bottom As Double, _
                ByVal top As Double, _
                ByVal Color As Long, _
                ByVal outlineColor As Long)
Dim barOutlineScaleX As Double
Dim barOutlineScaleY As Double

If upBar And Not solidUpBody Then
    mCanvas.setBrushAttributes mCanvas.backColor, FillStyles.FillSolid
    barOutlineScaleX = mCanvas.convertPixelsToLogicalX(outlineThickness - 1)
    barOutlineScaleY = mCanvas.convertPixelsToLogicalY(outlineThickness - 1)
    
    If bottom = mOpen And top = mClose Then
        ' draw the whole of the body
        mCanvas.setPenAttributes outlineColor, _
                                outlineThickness, _
                                LineStyles.LineInsideSolid, _
                                DrawModes.DrawModeCopyPen
        mCanvas.drawRectangle mX - mBarSemiWidth, _
                        bottom, _
                        mX + mBarSemiWidth, _
                        top
    Else
        
        ' first blank out the area
        mCanvas.setPenAttributes mCanvas.backColor, _
                                1, _
                                LineStyles.LineInsideSolid, _
                                DrawModes.DrawModeCopyPen
        
        mCanvas.drawRectangle mX - mBarSemiWidth, _
                        bottom, _
                        mX + mBarSemiWidth, _
                        top
        
        mCanvas.setPenAttributes Color, _
                                1, _
                                LineStyles.LineInsideSolid, _
                                DrawModes.DrawModeCopyPen
        ' we'll draw the outlines of the bar as filled boxes
        ' so set the fill color to the outline color
        mCanvas.setBrushAttributes IIf(Color <> -1, Color, outlineColor), FillStyles.FillSolid
        
        If bottom = mOpen Then
            ' draw the base of the body
            If outlineThickness = 1 Then
                mCanvas.drawLine mX - mBarSemiWidth, _
                                bottom, _
                                mX + mBarSemiWidth, _
                                bottom
            Else
                mCanvas.drawRectangle mX - mBarSemiWidth, _
                                bottom, _
                                mX + mBarSemiWidth, _
                                bottom + barOutlineScaleY
            End If
        End If
        If top = mClose Then
            ' draw the top of the body
            If outlineThickness = 1 Then
                mCanvas.drawLine mX - mBarSemiWidth, _
                                top, _
                                mX + mBarSemiWidth, _
                                top
            Else
                mCanvas.drawRectangle mX - mBarSemiWidth, _
                                top - barOutlineScaleY, _
                                mX + mBarSemiWidth, _
                                top
            End If
        End If
        If outlineThickness = 1 Then
            mCanvas.drawLine mX - mBarSemiWidth, _
                            bottom, _
                            mX - mBarSemiWidth, _
                            top
            mCanvas.drawLine mX + mBarSemiWidth, _
                            bottom, _
                            mX + mBarSemiWidth, _
                            top
        Else
            mCanvas.drawRectangle mX - mBarSemiWidth, _
                            bottom, _
                            mX - mBarSemiWidth + barOutlineScaleX, _
                            top
            mCanvas.drawRectangle mX + mBarSemiWidth - barOutlineScaleX, _
                            bottom, _
                            mX + mBarSemiWidth, _
                            top
        End If
    End If
Else
    mCanvas.setBrushAttributes Color, FillStyles.FillSolid
    mCanvas.setPenAttributes outlineColor, _
                            outlineThickness, _
                            LineStyles.LineInsideSolid, _
                            DrawModes.DrawModeCopyPen
    mCanvas.drawRectangle mX - mBarSemiWidth, _
                        bottom, _
                        mX + mBarSemiWidth, _
                        top
End If
End Sub

Friend Sub drawPartialBar( _
                ByRef area As TRectangle, _
                ByVal Color As Long, _
                ByVal outlineColor As Long)
Dim barThicknessScaleX As Double
Dim barThicknessScaleY As Double
If Not mVisible Then Exit Sub
If blank Then Exit Sub

If displayMode = BarDisplayModeCandlestick Then
    If downBar Then
        If area.top < mClose Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), area.top, outlineColor
        ElseIf area.top >= mClose And area.top <= mOpen And area.bottom < mClose Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), mClose, outlineColor
            drawBody mClose, area.top, Color, outlineColor
        ElseIf area.top > mOpen And area.bottom < mClose Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), mClose, outlineColor
            drawBody mClose, mOpen, Color, outlineColor
            drawTail mOpen, IIf(mHigh > area.top, area.top, mHigh), outlineColor
        ElseIf area.top <= mOpen And area.bottom >= mClose Then
            drawBody area.bottom, area.top, Color, outlineColor
        ElseIf area.top > mOpen And area.bottom <= mOpen And area.bottom >= mClose Then
            drawTail mOpen, IIf(mHigh > area.top, area.top, mHigh), outlineColor
            drawBody area.bottom, mOpen, Color, outlineColor
        Else
            drawTail area.bottom, IIf(mHigh > area.top, area.top, mHigh), outlineColor
        End If
    Else
        If area.top < mOpen Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), area.top, outlineColor
        ElseIf area.top >= mOpen And area.top <= mClose And area.bottom < mOpen Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), mOpen, outlineColor
            drawBody mOpen, area.top, Color, outlineColor
        ElseIf area.top > mClose And area.bottom < mOpen Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), mOpen, outlineColor
            drawBody mOpen, mClose, Color, outlineColor
            drawTail mClose, IIf(mHigh > area.top, area.top, mHigh), outlineColor
        ElseIf area.top <= mClose And area.bottom >= mOpen Then
            drawBody area.bottom, area.top, Color, outlineColor
        ElseIf area.top > mClose And area.bottom <= mClose And area.bottom >= mOpen Then
            drawTail mClose, IIf(mHigh > area.top, area.top, mHigh), outlineColor
            drawBody area.bottom, mClose, Color, outlineColor
        Else
            drawTail area.bottom, IIf(mHigh > area.top, area.top, mHigh), outlineColor
        End If
    End If
Else
    mCanvas.setPenAttributes outlineColor, _
                            1, _
                            LineStyles.LineInsideSolid, _
                            DrawModes.DrawModeCopyPen
    If barThickness = 1 Then
        Dim onePixel As Double
        onePixel = mCanvas.convertPixelsToLogicalX(1)
        If mOpen >= area.bottom And mOpen <= area.top Then
            mCanvas.drawLine mX - mBarSemiWidth, _
                            mOpen, _
                            mX, _
                            mOpen
        End If
        If mClose >= area.bottom And mClose <= area.top Then
            mCanvas.drawLine mX, _
                            mClose, _
                            mX + mBarSemiWidth + onePixel, _
                            mClose
        End If
        mCanvas.drawLine mX, _
                        IIf(area.bottom > mLow, area.bottom, mLow), _
                        mX, _
                        IIf(area.top < mHigh, area.top, mHigh)
    Else
        mCanvas.setBrushAttributes outlineColor, FillStyles.FillSolid
        barThicknessScaleX = mCanvas.convertPixelsToLogicalX(barThickness - 1)
        barThicknessScaleY = mCanvas.convertPixelsToLogicalY(barThickness - 1)
        
        If mOpen >= area.bottom And mOpen <= area.top Then
            If (mOpen - barThicknessScaleY / 2) < mLow Then
                mCanvas.drawRectangle mX - mBarSemiWidth, _
                                mLow, _
                                mX + barThicknessScaleX / 2, _
                                mLow + barThicknessScaleY
            ElseIf (mOpen + barThicknessScaleY / 2) > mHigh Then
                mCanvas.drawRectangle mX - mBarSemiWidth, _
                                mHigh - barThicknessScaleY, _
                                mX + barThicknessScaleX / 2, _
                                mHigh
            Else
                mCanvas.drawRectangle mX - mBarSemiWidth, _
                                mOpen - barThicknessScaleY / 2, _
                                mX + barThicknessScaleX / 2, _
                                mOpen + barThicknessScaleY / 2
            End If
        End If
        
        If mClose >= area.bottom And mClose <= area.top Then
            If (mClose - barThicknessScaleY / 2) < mLow Then
                mCanvas.drawRectangle mX - barThicknessScaleX / 2, _
                                mLow, _
                                mX + mBarSemiWidth, _
                                mLow + barThicknessScaleY
            ElseIf (mClose + barThicknessScaleY / 2) > mHigh Then
                mCanvas.drawRectangle mX - barThicknessScaleX / 2, _
                                mHigh - barThicknessScaleY, _
                                mX + mBarSemiWidth, _
                                mHigh
            Else
                mCanvas.drawRectangle mX - barThicknessScaleX / 2, _
                                mClose - barThicknessScaleY / 2, _
                                mX + mBarSemiWidth, _
                                mClose + barThicknessScaleY / 2
            End If
        End If
        
        mCanvas.drawRectangle mX - barThicknessScaleX / 2, _
                        IIf(area.bottom > mLow, area.bottom, mLow), _
                        mX + barThicknessScaleX / 2, _
                        IIf(area.top < mHigh, area.top, mHigh)
    End If
End If

End Sub

Private Sub drawTail(ByVal bottom As Double, ByVal top As Double, ByVal Color As Long)
Dim tailThicknessScaleX As Double
mCanvas.setPenAttributes Color, _
                        1, _
                        LineStyles.LineInsideSolid, _
                        DrawModes.DrawModeCopyPen
If tailThickness > 1 Then
    mCanvas.setBrushAttributes Color, FillStyles.FillSolid
    tailThicknessScaleX = mCanvas.convertPixelsToLogicalX(tailThickness - 1)
    mCanvas.drawRectangle mX - tailThicknessScaleX / 2, _
                    bottom, _
                    mX + tailThicknessScaleX / 2, _
                    top
Else
    mCanvas.drawRectangle mX, bottom, mX, top
End If
End Sub

Private Sub setBarSemiWidth()
Select Case displayMode
Case BarDisplayModeBar
    mBarSemiWidth = (2 * barWidth - 0.2) / 2
    If mBarSemiWidth > 0.5 Then mBarSemiWidth = 0.5
Case BarDisplayModeCandlestick
    mBarSemiWidth = barWidth / 2
Case BarDisplayModeLine
    mBarSemiWidth = barWidth / 2
End Select
End Sub

Public Sub tick(ByVal Price As Double)
Dim clearRect As TRectangle

If Price = mClose Then Exit Sub

If Not mBlank Then
    With clearRect
        .isValid = False
        If upBar Then
            If Price < mOpen Then
                ' is now a downbar so redraw the whole thing
                .left = mX - mBarSemiWidth
                .right = mX + mBarSemiWidth
                .bottom = IIf(Price < mLow, Price, mLow)
                .top = mHigh
                .isValid = True
            ElseIf Price < mClose Then
                .left = mX - mBarSemiWidth
                .right = mX + mBarSemiWidth
                .bottom = Price
                .top = mClose
                .isValid = True
            ElseIf Price > mClose Then
                .left = mX - mBarSemiWidth
                .right = mX + mBarSemiWidth
                .bottom = mClose
                .top = Price
                .isValid = True
            End If
        Else
            If Price >= mOpen Then
                ' now an upbar so redraw the whole thing
                .left = mX - mBarSemiWidth
                .right = mX + mBarSemiWidth
                .bottom = mLow
                .top = IIf(Price > mHigh, Price, mHigh)
                .isValid = True
            ElseIf Price > mClose Then
                .left = mX - mBarSemiWidth
                .right = mX + mBarSemiWidth
                .bottom = mClose
                .top = Price
                .isValid = True
            ElseIf Price < mClose Then
                .left = mX - mBarSemiWidth
                .right = mX + mBarSemiWidth
                .bottom = Price
                .top = mClose
                .isValid = True
            End If
        End If
        If .isValid Then undraw clearRect
    End With
End If

If mBlank Then
    mBlank = False
    mOpen = Price
End If
If Price > mHigh Then
    mHigh = Price
    mBoundingRect.top = Price
End If
If Price < mLow Then
    mLow = Price
    mBoundingRect.bottom = Price
End If
mBoundingRect.isValid = True
mClose = Price

mEventProxy.objectChanged mHandle, mBoundingRect

'If mHigh >= mSurface.ScaleTop Or mLow <= (mSurface.ScaleTop + mSurface.ScaleHeight) Then
'    mEventProxy.boundsExceeded
'End If
End Sub

Private Sub undraw(ByRef area As TRectangle)
If Not mVisible Then Exit Sub
If Not area.isValid Then
    drawBar mCanvas.backColor, mCanvas.backColor
    mEventProxy.objectUndrawn mHandle, mBoundingRect
Else
    drawPartialBar area, mCanvas.backColor, mCanvas.backColor
    mEventProxy.objectUndrawn mHandle, area
End If
End Sub



