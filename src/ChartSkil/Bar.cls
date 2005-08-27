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

Private Const BarSemiWidth As Single = 0.4

'================================================================================
' Member variables
'================================================================================

Private mCanvas As canvas
Private mLayer As Long
Private mHandle As Long

Private mInScope As Boolean
Private mVisible As Boolean

Private mEventProxy As ChartRegionEventProxy

Private mBoundingRect As TRectangle

Private mPeriodNumber As Long
Private mOpen As Double
Private mHigh As Double
Private mLow As Double
Private mClose As Double
Private mBlank As Boolean   ' indicates that there were no ticks during the bar

Private mTailThickness As Long
Private mOutlineThickness As Long
Private mUpColour As Long
Private mDownColour As Long
Private mDisplayAsCandlestick As Boolean
Private mSolidUpBody As Boolean
Private mBarThickness As Long
Private mCandleWidth As Single

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mOpen = 0
mLow = PlusInfinityDouble
mHigh = MinusInfinityDouble
mClose = 0
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
IGraphicObject_noDraw = mBlank
End Property

Private Property Get IGraphicObject_periodNumber() As Long
IGraphicObject_periodNumber = mPeriodNumber
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

Friend Property Let barThickness(ByVal value As Long)
mBarThickness = value
End Property

Public Property Get blank() As Boolean
blank = mBlank
End Property

Friend Property Let candleWidth(ByVal value As Single)
mCandleWidth = value
End Property

Friend Property Let canvas(ByVal value As canvas)
Set mCanvas = value
End Property

Public Property Get closePrice() As Double
closePrice = mClose
End Property

Friend Property Let displayAsCandlestick(ByVal value As Boolean)
mDisplayAsCandlestick = value
End Property

Public Property Get downBar() As Boolean
downBar = (mClose <= mOpen)
End Property

Friend Property Let downColour(ByVal value As Long)
mDownColour = value
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

Friend Property Let layer(ByVal value As Long)
mLayer = value
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Get lowPrice() As Double
lowPrice = mLow
End Property

Public Property Get openPrice() As Double
openPrice = mOpen
End Property

Friend Property Let outlineThickness(ByVal value As Long)
mOutlineThickness = value
End Property

Public Property Get periodNumber() As Long
periodNumber = mPeriodNumber
End Property

Public Property Let periodNumber(ByVal newValue As Long)
mPeriodNumber = newValue
mBoundingRect.left = mPeriodNumber - 0.5
mBoundingRect.right = mPeriodNumber + 0.5
End Property

Friend Property Let solidUpBody(ByVal value As Boolean)
mSolidUpBody = value
End Property

Friend Property Let tailThickness(ByVal value As Long)
mTailThickness = value
End Property

Public Property Get upBar() As Boolean
upBar = (mClose >= mOpen)
End Property

Friend Property Let upColour(ByVal value As Long)
mUpColour = value
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

Private Sub draw(ByRef area As TRectangle)
If Not area.isValid Then
    drawBar IIf(upBar, mUpColour, mDownColour)
Else
    drawPartialBar area, IIf(upBar, mUpColour, mDownColour)
End If
End Sub


Private Sub drawBar(ByVal colour As Long)
Dim barThicknessScaleX As Double
Dim barThicknessScaleY As Double

If Not mVisible Then Exit Sub
If blank Then Exit Sub

If mDisplayAsCandlestick Then
    If upBar Then
        drawBody mOpen, mClose, colour
    Else
        drawBody mClose, mOpen, colour
    End If
    
    If upBar Then
        If mLow < mOpen Then
            drawTail mLow, mOpen, colour
        End If
        If mHigh > mClose Then
            drawTail mClose, mHigh, colour
        End If
    Else
        If mHigh > mOpen Then
            drawTail mOpen, mHigh, colour
        End If
        If mLow < mClose Then
            drawTail mLow, mClose, colour
        End If
    End If
Else
    mCanvas.setBrushAttributes colour, FillStyles.FillSolid
    If mBarThickness = 1 Then
        Dim onePixel As Double
        onePixel = mCanvas.convertPixelsToLogicalX(1)
        mCanvas.setPenAttributes colour, _
                                1, _
                                LineStyles.LineSolid, _
                                DrawModes.DrawModeCopyPen
        mCanvas.drawLine mPeriodNumber - BarSemiWidth, _
                        mOpen, _
                        mPeriodNumber, _
                        mOpen
        mCanvas.drawLine mPeriodNumber, _
                        mLow, _
                        mPeriodNumber, _
                        mHigh
        mCanvas.drawLine mPeriodNumber, _
                        mClose, _
                        mPeriodNumber + BarSemiWidth + onePixel, _
                        mClose
    Else
        mCanvas.setPenAttributes colour, _
                                1, _
                                LineStyles.LineInsideSolid, _
                                DrawModes.DrawModeCopyPen
        barThicknessScaleX = mCanvas.convertPixelsToLogicalX(mBarThickness - 1)
        barThicknessScaleY = mCanvas.convertPixelsToLogicalY(mBarThickness - 1)
        
        If (mOpen - barThicknessScaleY / 2) < mLow Then
            mCanvas.drawRectangle mPeriodNumber - BarSemiWidth, _
                                    mLow, _
                                    mPeriodNumber + barThicknessScaleX / 2, _
                                    mLow + barThicknessScaleY
        ElseIf (mOpen + barThicknessScaleY / 2) > mHigh Then
            mCanvas.drawRectangle mPeriodNumber - BarSemiWidth, _
                                    mHigh - barThicknessScaleY, _
                                    mPeriodNumber + barThicknessScaleX / 2, _
                                    mHigh
        Else
            mCanvas.drawRectangle mPeriodNumber - BarSemiWidth, _
                                    mOpen - barThicknessScaleY / 2, _
                                    mPeriodNumber + barThicknessScaleX / 2, _
                                    mOpen + barThicknessScaleY / 2
        End If
        
        mCanvas.drawRectangle mPeriodNumber - barThicknessScaleX / 2, _
                                mLow, _
                                mPeriodNumber + barThicknessScaleX / 2, _
                                mHigh
        
        If (mClose - barThicknessScaleY / 2) < mLow Then
            mCanvas.drawRectangle mPeriodNumber - barThicknessScaleX / 2, _
                                    mLow, _
                                    mPeriodNumber + BarSemiWidth, _
                                    mLow + barThicknessScaleY
        ElseIf (mClose + barThicknessScaleY / 2) > mHigh Then
            mCanvas.drawRectangle mPeriodNumber - barThicknessScaleX / 2, _
                                    mHigh - barThicknessScaleY, _
                                    mPeriodNumber + BarSemiWidth, _
                                    mHigh
        Else
            mCanvas.drawRectangle mPeriodNumber - barThicknessScaleX / 2, _
                                    mClose - barThicknessScaleY / 2, _
                                    mPeriodNumber + BarSemiWidth, _
                                    mClose + barThicknessScaleY / 2
        End If
    End If
End If

End Sub

Private Sub drawBody(ByVal bottom As Double, ByVal top As Double, ByVal colour As Long)
Dim barOutlineScaleX As Double
Dim barOutlineScaleY As Double
mCanvas.setPenAttributes colour, _
                        mOutlineThickness, _
                        LineStyles.LineInsideSolid, _
                        DrawModes.DrawModeCopyPen
If upBar And Not mSolidUpBody Then
    mCanvas.setBrushAttributes mCanvas.backColor, FillStyles.FillSolid
    barOutlineScaleX = mCanvas.convertPixelsToLogicalX(mOutlineThickness - 1)
    barOutlineScaleY = mCanvas.convertPixelsToLogicalY(mOutlineThickness - 1)
    If bottom = mOpen And top = mClose Then
        mCanvas.drawLine mPeriodNumber - (mCandleWidth / 2), _
                        bottom, _
                        mPeriodNumber + (mCandleWidth / 2), _
                        top
    Else
        mCanvas.drawLine mPeriodNumber - (mCandleWidth / 2), _
                        bottom, _
                        mPeriodNumber + (mCandleWidth / 2), _
                        top
        If bottom = mOpen Then
            If mOutlineThickness = 1 Then
                mCanvas.drawLine mPeriodNumber - (mCandleWidth / 2), _
                                bottom, _
                                mPeriodNumber + (mCandleWidth / 2), _
                                bottom
            Else
                mCanvas.drawRectangle mPeriodNumber - (mCandleWidth / 2), _
                                bottom, _
                                mPeriodNumber + (mCandleWidth / 2), _
                                bottom + barOutlineScaleY
            End If
        End If
        If top = mClose Then
            If mOutlineThickness = 1 Then
                mCanvas.drawLine mPeriodNumber - (mCandleWidth / 2), _
                                top, _
                                mPeriodNumber + (mCandleWidth / 2), _
                                top
            Else
                mCanvas.drawRectangle mPeriodNumber - (mCandleWidth / 2), _
                                top - barOutlineScaleY, _
                                mPeriodNumber + (mCandleWidth / 2), _
                                top
            End If
        End If
        If mOutlineThickness = 1 Then
            mCanvas.drawLine mPeriodNumber - (mCandleWidth / 2), _
                            bottom, _
                            mPeriodNumber - (mCandleWidth / 2), _
                            top
            mCanvas.drawLine mPeriodNumber + (mCandleWidth / 2), _
                            bottom, _
                            mPeriodNumber + (mCandleWidth / 2), _
                            top
        Else
            mCanvas.drawRectangle mPeriodNumber - (mCandleWidth / 2), _
                            bottom, _
                            mPeriodNumber - (mCandleWidth / 2) + barOutlineScaleX, _
                            top
            mCanvas.drawRectangle mPeriodNumber + (mCandleWidth / 2) - barOutlineScaleX, _
                            bottom, _
                            mPeriodNumber + (mCandleWidth / 2), _
                            top
        End If
    End If
Else
    mCanvas.setBrushAttributes colour, FillStyles.FillSolid
    mCanvas.drawRectangle mPeriodNumber - (mCandleWidth / 2), _
                        bottom, _
                        mPeriodNumber + (mCandleWidth / 2), _
                        top
End If
End Sub

Friend Sub drawPartialBar(ByRef area As TRectangle, ByVal colour As Long)
Dim barThicknessScaleX As Double
Dim barThicknessScaleY As Double
If Not mVisible Then Exit Sub
If blank Then Exit Sub

If mDisplayAsCandlestick Then
    If downBar Then
        If area.top < mClose Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), area.top, colour
        ElseIf area.top >= mClose And area.top <= mOpen And area.bottom < mClose Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), mClose, colour
            drawBody mClose, area.top, colour
        ElseIf area.top > mOpen And area.bottom < mClose Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), mClose, colour
            drawBody mClose, mOpen, colour
            drawTail mOpen, IIf(mHigh > area.top, area.top, mHigh), colour
        ElseIf area.top <= mOpen And area.bottom >= mClose Then
            drawBody area.bottom, area.top, colour
        ElseIf area.top > mOpen And area.bottom <= mOpen And area.bottom >= mClose Then
            drawTail mOpen, IIf(mHigh > area.top, area.top, mHigh), colour
            drawBody area.bottom, mOpen, colour
        Else
            drawTail area.bottom, IIf(mHigh > area.top, area.top, mHigh), colour
        End If
    Else
        If area.top < mOpen Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), area.top, colour
        ElseIf area.top >= mOpen And area.top <= mClose And area.bottom < mOpen Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), mOpen, colour
            drawBody mOpen, area.top, colour
        ElseIf area.top > mClose And area.bottom < mOpen Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), mOpen, colour
            drawBody mOpen, mClose, colour
            drawTail mClose, IIf(mHigh > area.top, area.top, mHigh), colour
        ElseIf area.top <= mClose And area.bottom >= mOpen Then
            drawBody area.bottom, area.top, colour
        ElseIf area.top > mClose And area.bottom <= mClose And area.bottom >= mOpen Then
            drawTail mClose, IIf(mHigh > area.top, area.top, mHigh), colour
            drawBody area.bottom, mClose, colour
        Else
            drawTail area.bottom, IIf(mHigh > area.top, area.top, mHigh), colour
        End If
    End If
Else
    mCanvas.setPenAttributes colour, _
                            1, _
                            LineStyles.LineInsideSolid, _
                            DrawModes.DrawModeCopyPen
    If mBarThickness = 1 Then
        Dim onePixel As Double
        onePixel = mCanvas.convertPixelsToLogicalX(1)
        If mOpen >= area.bottom And mOpen <= area.top Then
            mCanvas.drawLine mPeriodNumber - BarSemiWidth, _
                            mOpen, _
                            mPeriodNumber, _
                            mOpen
        End If
        If mClose >= area.bottom And mClose <= area.top Then
            mCanvas.drawLine mPeriodNumber, _
                            mClose, _
                            mPeriodNumber + BarSemiWidth + onePixel, _
                            mClose
        End If
        mCanvas.drawLine mPeriodNumber, _
                        IIf(area.bottom > mLow, area.bottom, mLow), _
                        mPeriodNumber, _
                        IIf(area.top < mHigh, area.top, mHigh)
    Else
        mCanvas.setBrushAttributes colour, FillStyles.FillSolid
        barThicknessScaleX = mCanvas.convertPixelsToLogicalX(mBarThickness - 1)
        barThicknessScaleY = mCanvas.convertPixelsToLogicalY(mBarThickness - 1)
        
        If mOpen >= area.bottom And mOpen <= area.top Then
            If (mOpen - barThicknessScaleY / 2) < mLow Then
                mCanvas.drawRectangle mPeriodNumber - BarSemiWidth, _
                                mLow, _
                                mPeriodNumber + barThicknessScaleX / 2, _
                                mLow + barThicknessScaleY
            ElseIf (mOpen + barThicknessScaleY / 2) > mHigh Then
                mCanvas.drawRectangle mPeriodNumber - BarSemiWidth, _
                                mHigh - barThicknessScaleY, _
                                mPeriodNumber + barThicknessScaleX / 2, _
                                mHigh
            Else
                mCanvas.drawRectangle mPeriodNumber - BarSemiWidth, _
                                mOpen - barThicknessScaleY / 2, _
                                mPeriodNumber + barThicknessScaleX / 2, _
                                mOpen + barThicknessScaleY / 2
            End If
        End If
        
        If mClose >= area.bottom And mClose <= area.top Then
            If (mClose - barThicknessScaleY / 2) < mLow Then
                mCanvas.drawRectangle mPeriodNumber - barThicknessScaleX / 2, _
                                mLow, _
                                mPeriodNumber + BarSemiWidth, _
                                mLow + barThicknessScaleY
            ElseIf (mClose + barThicknessScaleY / 2) > mHigh Then
                mCanvas.drawRectangle mPeriodNumber - barThicknessScaleX / 2, _
                                mHigh - barThicknessScaleY, _
                                mPeriodNumber + BarSemiWidth, _
                                mHigh
            Else
                mCanvas.drawRectangle mPeriodNumber - barThicknessScaleX / 2, _
                                mClose - barThicknessScaleY / 2, _
                                mPeriodNumber + BarSemiWidth, _
                                mClose + barThicknessScaleY / 2
            End If
        End If
        
        mCanvas.drawRectangle mPeriodNumber - barThicknessScaleX / 2, _
                        IIf(area.bottom > mLow, area.bottom, mLow), _
                        mPeriodNumber + barThicknessScaleX / 2, _
                        IIf(area.top < mHigh, area.top, mHigh)
    End If
End If

End Sub

Private Sub drawTail(ByVal bottom As Double, ByVal top As Double, ByVal colour As Long)
Dim tailThicknessScaleX As Double
mCanvas.setPenAttributes colour, _
                        1, _
                        LineStyles.LineInsideSolid, _
                        DrawModes.DrawModeCopyPen
If mTailThickness > 1 Then
    mCanvas.setBrushAttributes colour, FillStyles.FillSolid
    tailThicknessScaleX = mCanvas.convertPixelsToLogicalX(mTailThickness - 1)
    mCanvas.drawRectangle mPeriodNumber - tailThicknessScaleX / 2, _
                    bottom, _
                    mPeriodNumber + tailThicknessScaleX / 2, _
                    top
Else
    mCanvas.drawRectangle mPeriodNumber, bottom, mPeriodNumber, top
End If
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
                .left = mPeriodNumber - mCandleWidth / 2
                .right = mPeriodNumber + mCandleWidth / 2
                .bottom = IIf(Price < mLow, Price, mLow)
                .top = mHigh
                .isValid = True
            ElseIf Price < mClose Then
                .left = mPeriodNumber - mCandleWidth / 2
                .right = mPeriodNumber + mCandleWidth / 2
                .bottom = Price
                .top = mClose
                .isValid = True
            ElseIf Price > mClose Then
                .left = mPeriodNumber - mCandleWidth / 2
                .right = mPeriodNumber + mCandleWidth / 2
                .bottom = mClose
                .top = Price
                .isValid = True
            End If
        Else
            If Price >= mOpen Then
                ' now an upbar so redraw the whole thing
                .left = mPeriodNumber - mCandleWidth / 2
                .right = mPeriodNumber + mCandleWidth / 2
                .bottom = mLow
                .top = IIf(Price > mHigh, Price, mHigh)
                .isValid = True
            ElseIf Price > mClose Then
                .left = mPeriodNumber - mCandleWidth / 2
                .right = mPeriodNumber + mCandleWidth / 2
                .bottom = mClose
                .top = Price
                .isValid = True
            ElseIf Price < mClose Then
                .left = mPeriodNumber - mCandleWidth / 2
                .right = mPeriodNumber + mCandleWidth / 2
                .bottom = Price
                .top = mClose
                .isValid = True
            End If
        End If
        If .isValid Then undraw clearRect
    End With
End If

If mOpen = 0 Then mOpen = Price
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
mBlank = False

mEventProxy.objectChanged mHandle, mBoundingRect

'If mHigh >= mSurface.ScaleTop Or mLow <= (mSurface.ScaleTop + mSurface.ScaleHeight) Then
'    mEventProxy.boundsExceeded
'End If
End Sub

Private Sub undraw(ByRef area As TRectangle)
If Not mVisible Then Exit Sub
If Not area.isValid Then
    drawBar mCanvas.backColor
    mEventProxy.objectUndrawn mHandle, mBoundingRect
Else
    drawPartialBar area, mCanvas.backColor
    mEventProxy.objectUndrawn mHandle, area
End If
End Sub



