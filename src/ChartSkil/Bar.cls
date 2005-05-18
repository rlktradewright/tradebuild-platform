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
' Member variables and constants
'================================================================================

Private mSurface As PictureBox
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

mBoundingRect.isvalid = False
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

Private Sub IGraphicObject_draw(ByRef areas() As TRectangle)
Dim i As Long
For i = 0 To UBound(areas)
    draw areas(i)
Next
End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = False
End Property

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

Private Sub IGraphicObject_undraw(ByRef area As TRectangle)
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

Friend Property Let candleWidth(ByVal value As Single)
mCandleWidth = value
End Property

Public Property Get blank() As Boolean
blank = mBlank
End Property

Public Property Get closePrice() As Double
closePrice = mClose
End Property

'Public Property Let closePrice(newValue As Double)
'mClose = newValue
'mBlank = False
'End Property

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

'Public Property Let highPrice(newValue As Double)
'mHigh = newValue
'mBlank = False
'mBoundingRect.top = newValue
'End Property

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

'Public Property Let lowPrice(newValue As Double)
'mLow = newValue
'mBlank = False
'mBoundingRect.bottom = newValue
'End Property

Public Property Get openPrice() As Double
openPrice = mOpen
End Property

'Public Property Let openPrice(newValue As Double)
'mOpen = newValue
'mBlank = False
'End Property

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

Friend Property Let surface(ByVal value As PictureBox)
Set mSurface = value
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
If Not area.isvalid Then
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

mSurface.DrawMode = vbCopyPen
If mDisplayAsCandlestick Then
'    mSurface.DrawWidth = mOutlineThickness
'    mSurface.DrawStyle = LineStyles.LineInsideSolid
'    mSurface.FillStyle = FillStyles.FillTransparent
    If upBar Then
        drawBody mOpen, mClose, colour
    Else
        drawBody mClose, mOpen, colour
    End If
'    If downBar Or mSolidUpBody Then
'        mSurface.Line (mPeriodNumber - (mCandleWidth / 2), mOpen)-(mPeriodNumber + (mCandleWidth / 2), mClose), colour, BF
'    Else
'        mSurface.Line (mPeriodNumber - (mCandleWidth / 2), mOpen)-(mPeriodNumber + (mCandleWidth / 2), mClose), colour, B
'    End If
    
    
'    mSurface.DrawStyle = LineStyles.LineSolid
'    mSurface.DrawWidth = mTailThickness
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
    mSurface.DrawWidth = 1
    mSurface.DrawStyle = LineStyles.LineSolid
    If mBarThickness = 1 Then
        mSurface.Line (mPeriodNumber - mCandleWidth / 2, mOpen)-(mPeriodNumber, mOpen), IIf(upBar, mUpColour, mDownColour)
        mSurface.Line (mPeriodNumber, mLow)-(mPeriodNumber, mHigh), IIf(upBar, mUpColour, mDownColour)
        mSurface.Line (mPeriodNumber, mClose)-(mPeriodNumber + mCandleWidth / 2, mClose), IIf(upBar, mUpColour, mDownColour)
    Else
        mSurface.DrawStyle = LineStyles.LineInsideSolid
        barThicknessScaleX = Abs(((mBarThickness - 1) * Screen.TwipsPerPixelX) / (mSurface.Width / mSurface.scaleWidth))
        barThicknessScaleY = Abs(((mBarThickness - 1) * Screen.TwipsPerPixelY) / (mSurface.Height / mSurface.scaleHeight))
        
        If (mOpen - barThicknessScaleY / 2) < mLow Then
            mSurface.Line (mPeriodNumber - 1.5 * barThicknessScaleX, mLow)-(mPeriodNumber + barThicknessScaleX / 2, mLow + barThicknessScaleY), IIf(upBar, mUpColour, mDownColour), BF
        ElseIf (mOpen + barThicknessScaleY / 2) > mHigh Then
            mSurface.Line (mPeriodNumber - 1.5 * barThicknessScaleX, mHigh - barThicknessScaleY)-(mPeriodNumber + barThicknessScaleX / 2, mHigh), IIf(upBar, mUpColour, mDownColour), BF
        Else
            mSurface.Line (mPeriodNumber - 1.5 * barThicknessScaleX, mOpen - barThicknessScaleY / 2)-(mPeriodNumber + barThicknessScaleX / 2, mOpen + barThicknessScaleY / 2), IIf(upBar, mUpColour, mDownColour), BF
        End If
        
        mSurface.Line (mPeriodNumber - barThicknessScaleX / 2, mLow)-(mPeriodNumber + barThicknessScaleX / 2, mHigh), IIf(upBar, mUpColour, mDownColour), BF
        
        If (mClose - barThicknessScaleY / 2) < mLow Then
            mSurface.Line (mPeriodNumber - barThicknessScaleX / 2, mLow)-(mPeriodNumber + 1.5 * barThicknessScaleX, mLow + barThicknessScaleY), IIf(upBar, mUpColour, mDownColour), BF
        ElseIf (mClose + barThicknessScaleY / 2) > mHigh Then
            mSurface.Line (mPeriodNumber - barThicknessScaleX / 2, mHigh - barThicknessScaleY)-(mPeriodNumber + 1.5 * barThicknessScaleX, mHigh), IIf(upBar, mUpColour, mDownColour), BF
        Else
            mSurface.Line (mPeriodNumber - barThicknessScaleX / 2, mClose - barThicknessScaleY / 2)-(mPeriodNumber + 1.5 * barThicknessScaleX, mClose + barThicknessScaleY / 2), IIf(upBar, mUpColour, mDownColour), BF
        End If
    End If
End If

End Sub

Private Sub drawBody(ByVal bottom As Double, ByVal top As Double, ByVal colour As Long)
Dim barOutlineScaleX As Double
Dim barOutlineScaleY As Double
mSurface.DrawWidth = 1
mSurface.FillStyle = FillStyles.FillSolid
If upBar And Not mSolidUpBody Then
    mSurface.FillColor = mSurface.backColor
    barOutlineScaleX = Abs(((mOutlineThickness - 1) * Screen.TwipsPerPixelX) / (mSurface.Width / mSurface.scaleWidth))
    barOutlineScaleY = Abs(((mOutlineThickness - 1) * Screen.TwipsPerPixelY) / (mSurface.Height / mSurface.scaleHeight))
    If bottom = mOpen And top = mClose Then
        mSurface.DrawStyle = LineStyles.LineInsideSolid
        mSurface.DrawWidth = mOutlineThickness
        mSurface.Line (mPeriodNumber - (mCandleWidth / 2), bottom)-(mPeriodNumber + (mCandleWidth / 2), top), colour, B
    Else
        mSurface.DrawStyle = LineStyles.LineInsideSolid
        mSurface.Line (mPeriodNumber - (mCandleWidth / 2), bottom)-(mPeriodNumber + (mCandleWidth / 2), top), mSurface.backColor, B
        If bottom = mOpen Then
            If mOutlineThickness = 1 Then
'                mSurface.DrawStyle = LineStyles.LineSolid
'                mSurface.Line (mPeriodNumber - (mCandleWidth / 2), bottom)-(mPeriodNumber + (mCandleWidth / 2), bottom), colour
                mSurface.Line (mPeriodNumber - (mCandleWidth / 2), bottom)-(mPeriodNumber + (mCandleWidth / 2), bottom), colour, B
            Else
                mSurface.Line (mPeriodNumber - (mCandleWidth / 2), bottom)-(mPeriodNumber + (mCandleWidth / 2), bottom + barOutlineScaleY), colour, BF
            End If
        End If
        If top = mClose Then
            If mOutlineThickness = 1 Then
'                mSurface.DrawStyle = LineStyles.LineSolid
'                mSurface.Line (mPeriodNumber - (mCandleWidth / 2), top)-(mPeriodNumber + (mCandleWidth / 2), top), colour
                mSurface.Line (mPeriodNumber - (mCandleWidth / 2), top)-(mPeriodNumber + (mCandleWidth / 2), top), colour, B
            Else
                mSurface.Line (mPeriodNumber - (mCandleWidth / 2), top - barOutlineScaleY)-(mPeriodNumber + (mCandleWidth / 2), top), colour, BF
            End If
        End If
        If mOutlineThickness = 1 Then
'            mSurface.DrawStyle = LineStyles.LineSolid
'            mSurface.Line (mPeriodNumber - (mCandleWidth / 2), bottom)-(mPeriodNumber - (mCandleWidth / 2), top), colour
'            mSurface.Line (mPeriodNumber + (mCandleWidth / 2), bottom)-(mPeriodNumber + (mCandleWidth / 2), top), colour
            mSurface.Line (mPeriodNumber - (mCandleWidth / 2), bottom)-(mPeriodNumber - (mCandleWidth / 2), top), colour, B
            mSurface.Line (mPeriodNumber + (mCandleWidth / 2), bottom)-(mPeriodNumber + (mCandleWidth / 2), top), colour, B
        Else
            mSurface.Line (mPeriodNumber - (mCandleWidth / 2), bottom)-(mPeriodNumber - (mCandleWidth / 2) + barOutlineScaleX, top), colour, BF
            mSurface.Line (mPeriodNumber + (mCandleWidth / 2) - barOutlineScaleX, bottom)-(mPeriodNumber + (mCandleWidth / 2), top), colour, BF
        End If
    End If
Else
    mSurface.DrawStyle = LineStyles.LineInsideSolid
    mSurface.Line (mPeriodNumber - (mCandleWidth / 2), bottom)-(mPeriodNumber + (mCandleWidth / 2), top), colour, BF
End If
End Sub

Friend Sub drawPartialBar(ByRef area As TRectangle, ByVal colour As Long)
Dim barThicknessScaleX As Double
Dim barThicknessScaleY As Double
If Not mVisible Then Exit Sub
If blank Then Exit Sub

mSurface.DrawMode = vbCopyPen
If mDisplayAsCandlestick Then
    mSurface.DrawWidth = mOutlineThickness
    mSurface.DrawStyle = LineStyles.LineInsideSolid
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
    mSurface.DrawWidth = 1
    mSurface.DrawStyle = LineStyles.LineSolid
    If mBarThickness = 1 Then
        If mOpen >= area.bottom And mOpen <= area.top Then
            mSurface.Line (mPeriodNumber - mCandleWidth / 2, mOpen)-(mPeriodNumber, mOpen), colour
        End If
        If mClose >= area.bottom And mClose <= area.top Then
            mSurface.Line (mPeriodNumber, mClose)-(mPeriodNumber + mCandleWidth / 2, mClose), colour
        End If
        mSurface.Line (mPeriodNumber, IIf(area.bottom > mLow, area.bottom, mLow))-(mPeriodNumber, IIf(area.top < mHigh, area.top, mHigh)), colour
    Else
        mSurface.DrawStyle = LineStyles.LineInsideSolid
        barThicknessScaleX = Abs(((mBarThickness - 1) * Screen.TwipsPerPixelX) / (mSurface.Width / mSurface.scaleWidth))
        barThicknessScaleY = Abs(((mBarThickness - 1) * Screen.TwipsPerPixelY) / (mSurface.Height / mSurface.scaleHeight))
        
        If mOpen >= area.bottom And mOpen <= area.top Then
            If (mOpen - barThicknessScaleY / 2) < mLow Then
                mSurface.Line (mPeriodNumber - 1.5 * barThicknessScaleX, mLow)-(mPeriodNumber + barThicknessScaleX / 2, mLow + barThicknessScaleY), colour, BF
            ElseIf (mOpen + barThicknessScaleY / 2) > mHigh Then
                mSurface.Line (mPeriodNumber - 1.5 * barThicknessScaleX, mHigh - barThicknessScaleY)-(mPeriodNumber + barThicknessScaleX / 2, mHigh), colour, BF
            Else
                mSurface.Line (mPeriodNumber - 1.5 * barThicknessScaleX, mOpen - barThicknessScaleY / 2)-(mPeriodNumber + barThicknessScaleX / 2, mOpen + barThicknessScaleY / 2), colour, BF
            End If
        End If
        
        If mClose >= area.bottom And mClose <= area.top Then
            If (mClose - barThicknessScaleY / 2) < mLow Then
                mSurface.Line (mPeriodNumber - barThicknessScaleX / 2, mLow)-(mPeriodNumber + 1.5 * barThicknessScaleX, mLow + barThicknessScaleY), colour, BF
            ElseIf (mClose + barThicknessScaleY / 2) > mHigh Then
                mSurface.Line (mPeriodNumber - barThicknessScaleX / 2, mHigh - barThicknessScaleY)-(mPeriodNumber + 1.5 * barThicknessScaleX, mHigh), colour, BF
            Else
                mSurface.Line (mPeriodNumber - barThicknessScaleX / 2, mClose - barThicknessScaleY / 2)-(mPeriodNumber + 1.5 * barThicknessScaleX, mClose + barThicknessScaleY / 2), colour, BF
            End If
        End If
        
        mSurface.Line (mPeriodNumber - barThicknessScaleX / 2, IIf(area.bottom > mLow, area.bottom, mLow))-(mPeriodNumber + barThicknessScaleX / 2, IIf(area.top < mHigh, area.top, mHigh)), colour, BF
    End If
End If

If mHigh >= mSurface.scaleTop Or mLow <= (mSurface.scaleTop + mSurface.scaleHeight) Then
    mEventProxy.boundsExceeded
End If

End Sub

Private Sub drawTail(ByVal bottom As Double, ByVal top As Double, ByVal colour As Long)
Dim tailThicknessScaleX As Double
mSurface.DrawWidth = 1
If mTailThickness > 1 Then
    mSurface.DrawStyle = LineStyles.LineInsideSolid
    mSurface.FillStyle = FillStyles.FillTransparent
    tailThicknessScaleX = Abs(((mTailThickness - 1) * Screen.TwipsPerPixelX) / (mSurface.Width / mSurface.scaleWidth))
    mSurface.Line (mPeriodNumber - tailThicknessScaleX / 2, bottom)-(mPeriodNumber + tailThicknessScaleX / 2, top), colour, BF
Else
    mSurface.DrawStyle = LineStyles.LineInsideSolid
    mSurface.Line (mPeriodNumber, bottom)-(mPeriodNumber, top), colour, B
End If
End Sub

Public Sub tick(ByVal Price As Double)
Dim clearRect As TRectangle

If Price = mClose Then Exit Sub

If Not mBlank Then
    With clearRect
        .isvalid = False
        If upBar Then
            If Price < mOpen Then
                ' is now a downbar so redraw the whole thing
                .left = mPeriodNumber - mCandleWidth / 2
                .right = mPeriodNumber + mCandleWidth / 2
                .bottom = IIf(Price < mLow, Price, mLow)
                .top = mHigh
                .isvalid = True
            ElseIf Price < mClose Then
                .left = mPeriodNumber - mCandleWidth / 2
                .right = mPeriodNumber + mCandleWidth / 2
                .bottom = Price
                .top = mClose
                .isvalid = True
            ElseIf Price > mClose Then
                .left = mPeriodNumber - mCandleWidth / 2
                .right = mPeriodNumber + mCandleWidth / 2
                .bottom = mClose
                .top = Price
                .isvalid = True
            End If
        Else
            If Price >= mOpen Then
                ' now an upbar so redraw the whole thing
                .left = mPeriodNumber - mCandleWidth / 2
                .right = mPeriodNumber + mCandleWidth / 2
                .bottom = mLow
                .top = IIf(Price > mHigh, Price, mHigh)
                .isvalid = True
            ElseIf Price > mClose Then
                .left = mPeriodNumber - mCandleWidth / 2
                .right = mPeriodNumber + mCandleWidth / 2
                .bottom = mClose
                .top = Price
                .isvalid = True
            ElseIf Price < mClose Then
                .left = mPeriodNumber - mCandleWidth / 2
                .right = mPeriodNumber + mCandleWidth / 2
                .bottom = Price
                .top = mClose
                .isvalid = True
            End If
        End If
        If .isvalid Then undraw clearRect
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
mBoundingRect.isvalid = True
mClose = Price
mBlank = False

mEventProxy.objectChanged mHandle, mBoundingRect

If mHigh >= mSurface.scaleTop Or mLow <= (mSurface.scaleTop + mSurface.scaleHeight) Then
    mEventProxy.boundsExceeded
End If
End Sub

Private Sub undraw(ByRef area As TRectangle)
If Not area.isvalid Then
    drawBar mSurface.backColor
    mEventProxy.objectUndrawn mHandle, mBoundingRect
Else
    drawPartialBar area, mSurface.backColor
    mEventProxy.objectUndrawn mHandle, area
End If
End Sub



