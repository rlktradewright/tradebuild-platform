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

'@================================================================================
' Interfaces
'@================================================================================

Implements IGraphicObject

'@================================================================================
' Events
'@================================================================================

Event Click()

Event DblCLick()

Event SelectionStateChanged()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                As String = "ChartSkil26"
Private Const ModuleName                As String = "Bar"

'================================================================================
' Member variables
'================================================================================

Private mCanvas As Canvas
Private mLayer As Long
Private mHandle As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean
Private mSeriesID As Long

Private mTimestamp As Date

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
Private mIsSetIncludeInAutoscale As Boolean

Private mTailThickness As Long
Private mOutlineThickness As Long
Private mBarColor As Long
Private mUpColor As Long
Private mDownColor As Long
Private mDisplayMode As BarDisplayModes
Private mIncludeInAutoscale As Boolean
Private mSolidUpBody As Boolean
Private mBarThickness As Long
Private mBarWidth As Single

Private mInScope As Boolean
Private mVisible As Boolean

Private mGOSP As GraphObjServiceProvider

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

Private Property Get IGraphicObject_boundingRectangle() As TRectangle
Dim failpoint As Long
On Error GoTo Err

setBarSemiWidth
calcBoundingRect
IGraphicObject_boundingRectangle = mBoundingRect

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_boundingRectangle" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Property

Private Property Get IGraphicObject_BoundingRectanglesAt(ByVal PeriodNumber As Long) As TRectangle()
'
End Property

Private Property Get IGraphicObject_capabilities() As GraphicObjectCapabilities
IGraphicObject_capabilities = GraphicObjectCapabilities.PartialRedraw
End Property

Private Sub IGraphicObject_Click()
RaiseEvent Click
End Sub

Private Sub IGraphicObject_DblCLick()
RaiseEvent DblCLick
End Sub

Private Sub IGraphicObject_draw( _
                ByRef areas() As TRectangle)
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

For i = 0 To UBound(areas)
    Draw areas(i)
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

Private Function IGraphicObject_HitTest(ByVal x As Double, ByVal y As Double) As Boolean
IGraphicObject_HitTest = RectContains(mBoundingRect, x, y)
End Function

Private Property Let IGraphicObject_handle(ByVal value As Long)
mHandle = value
End Property

Private Property Get IGraphicObject_includeInAutoscale() As Boolean
IGraphicObject_includeInAutoscale = True
End Property

Private Property Let IGraphicObject_inScope(ByVal value As Boolean)
mInScope = value
End Property

Private Property Get IGraphicObject_IsSelectable() As Boolean
IGraphicObject_IsSelectable = mIsSelectable
End Property

Private Property Get IGraphicObject_layer() As Long
IGraphicObject_layer = Layer
End Property

Private Property Get IGraphicObject_noDraw() As Boolean
IGraphicObject_noDraw = mBlank
End Property

Private Property Get IGraphicObject_periodNumber() As Long
IGraphicObject_periodNumber = Int(mX)
End Property

Private Property Get IGraphicObject_scaleDependent() As Boolean
IGraphicObject_scaleDependent = False
End Property

Private Property Let IGraphicObject_Selected(ByVal RHS As Boolean)
If RHS = mIsSelected Then
Else
    mIsSelected = RHS
    RaiseEvent SelectionStateChanged
End If
End Property

Private Property Get IGraphicObject_Selected() As Boolean
IGraphicObject_Selected = mIsSelected
End Property

Private Property Get IGraphicObject_SeriesID() As Long
IGraphicObject_SeriesID = mSeriesID
End Property

Private Property Get IGraphicObject_timestamp() As Date
IGraphicObject_timestamp = mTimestamp
End Property

'Private Sub IGraphicObject_undraw( _
'                ByRef area As TRectangle, _
'                ByVal hdc As Long)
'Dim failpoint As Long
'On Error GoTo Err
'
'undraw area
'
'Exit Sub
'
'Err:
'Dim errNumber As Long: errNumber = Err.Number
'Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_undraw" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
'Dim errDescription As String: errDescription = Err.Description
'gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
'Err.Raise errNumber, errSource, errDescription
'
'
'End Sub

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
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "Bar" & ":" & "barColor", _
            "Value is not a valid color or -1"

If mBarColor = value And mIsSetBarColor Then Exit Property
mBarColor = value
mIsSetBarColor = True
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get barColor() As Long
barColor = mBaseStyle.barColor
If Not mLocalStyle Is Nothing Then barColor = mLocalStyle.barColor
If mIsSetBarColor Then barColor = mBarColor
End Property

Public Property Let barThickness(ByVal value As Long)
If mBarThickness = value And mIsSetBarThickness Then Exit Property
undraw mBoundingRect
mBarThickness = value
mIsSetBarThickness = True
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get barThickness() As Long
barThickness = mBaseStyle.barThickness
If Not mLocalStyle Is Nothing Then barThickness = mLocalStyle.barThickness
If mIsSetBarThickness Then barThickness = mBarThickness
End Property

Public Property Let barWidth(ByVal value As Single)
If mBarWidth = value And mIsSetBarWidth Then Exit Property
undraw mBoundingRect
mBarWidth = value
mIsSetBarWidth = True
setBarSemiWidth
calcBoundingRect
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get barWidth() As Single
barWidth = mBaseStyle.barWidth
If Not mLocalStyle Is Nothing Then barWidth = mLocalStyle.barWidth
If mIsSetBarWidth Then barWidth = mBarWidth
End Property

Public Property Get blank() As Boolean
blank = mBlank
End Property

Public Property Get closePrice() As Double
closePrice = mClose
End Property

Public Property Let displayMode(ByVal value As BarDisplayModes)
If mDisplayMode = value And mIsSetDisplayMode Then Exit Property
undraw mBoundingRect
mDisplayMode = value
mIsSetDisplayMode = True
setBarSemiWidth
calcBoundingRect
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get displayMode() As BarDisplayModes
displayMode = mBaseStyle.displayMode
If Not mLocalStyle Is Nothing Then displayMode = mLocalStyle.displayMode
If mIsSetDisplayMode Then displayMode = mDisplayMode
End Property

Public Property Get downBar() As Boolean
downBar = (mClose <= mOpen)
End Property

Public Property Let downColor(ByVal value As Long)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "Bar" & ":" & "downColor", _
            "Value is not a valid color or -1"

If mDownColor = value And mIsSetDownColor Then Exit Property
mDownColor = value
mIsSetDownColor = True
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get downColor() As Long
downColor = mBaseStyle.downColor
If Not mLocalStyle Is Nothing Then downColor = mLocalStyle.downColor
If mIsSetDownColor Then downColor = mDownColor
If downColor = -1 Then downColor = upColor
End Property

Friend Property Get Handle() As Long
Handle = mHandle
End Property

Public Property Get highPrice() As Double
highPrice = mHigh
End Property

Public Property Let includeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value And mIsSetIncludeInAutoscale Then Exit Property
If Not value Then undraw mBoundingRect
mIncludeInAutoscale = value
mIsSetIncludeInAutoscale = True
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get includeInAutoscale() As Boolean
includeInAutoscale = mBaseStyle.includeInAutoscale
If Not mLocalStyle Is Nothing Then includeInAutoscale = mLocalStyle.includeInAutoscale
If mIsSetIncludeInAutoscale Then includeInAutoscale = mIncludeInAutoscale
End Property

Public Property Get InScope() As Boolean
InScope = mInScope
End Property

Public Property Get IsSelectable() As Boolean
IsSelectable = mIsSelectable
End Property

Public Property Get IsSelected() As Boolean
IsSelected = mIsSelected
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Let Layer(ByVal value As Long)
If mLayer = value Then Exit Property
mLayer = value
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get Layer() As Long
Layer = mLayer
End Property

Public Property Get LocalStyle() As BarStyle
Set LocalStyle = mLocalStyle.clone
End Property

Public Property Let LocalStyle( _
                ByVal value As BarStyle)

undraw mBoundingRect

Set mLocalStyle = value.clone
setBarSemiWidth
calcBoundingRect
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get lowPrice() As Double
lowPrice = mLow
End Property

Public Property Get openPrice() As Double
openPrice = mOpen
End Property

Public Property Let outlineThickness(ByVal value As Long)
If mOutlineThickness = value And mIsSetOutlineThickness Then Exit Property
mOutlineThickness = value
mIsSetOutlineThickness = True
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get outlineThickness() As Long
outlineThickness = mBaseStyle.outlineThickness
If Not mLocalStyle Is Nothing Then outlineThickness = mLocalStyle.outlineThickness
If mIsSetOutlineThickness Then outlineThickness = mOutlineThickness
End Property

Public Property Let solidUpBody(ByVal value As Boolean)
If mSolidUpBody = value And mIsSetSolidUpBody Then Exit Property
mSolidUpBody = value
mIsSetSolidUpBody = True
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get solidUpBody() As Boolean
solidUpBody = mBaseStyle.solidUpBody
If Not mLocalStyle Is Nothing Then solidUpBody = mLocalStyle.solidUpBody
If mIsSetSolidUpBody Then solidUpBody = mSolidUpBody
End Property

Public Property Let tailThickness(ByVal value As Long)
If mTailThickness = value And mIsSetTailThickness Then Exit Property
undraw mBoundingRect
mTailThickness = value
mIsSetTailThickness = True
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get tailThickness() As Long
tailThickness = mBaseStyle.tailThickness
If Not mLocalStyle Is Nothing Then tailThickness = mLocalStyle.tailThickness
If mIsSetTailThickness Then tailThickness = mTailThickness
End Property

Public Property Get upBar() As Boolean
upBar = (mClose >= mOpen)
End Property

Public Property Let upColor(ByVal value As Long)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "Bar" & ":" & "upColor", _
            "Value is not a valid color"

If mUpColor = value And mIsSetUpColor Then Exit Property
mUpColor = value
mIsSetUpColor = True
mGOSP.ObjectChanged mHandle, mBoundingRect
End Property

Public Property Get upColor() As Long
upColor = mBaseStyle.upColor
If Not mLocalStyle Is Nothing Then upColor = mLocalStyle.upColor
If mIsSetUpColor Then upColor = mUpColor
End Property

Public Property Get Visible() As Boolean
Visible = mVisible
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
                ByVal seriesId As Long, _
                ByVal timestamp As Date, _
                ByVal Key As String, _
                ByVal GOSP As GraphObjServiceProvider, _
                ByVal Layer As Long, _
                ByVal Canvas As Canvas, _
                ByVal baseStyle As BarStyle, _
                ByVal LocalStyle As BarStyle, _
                Optional ByVal IsSelectable As Boolean)
mSeriesID = seriesId
mTimestamp = timestamp
mKey = Key
Set mGOSP = GOSP
mLayer = Layer
Set mCanvas = Canvas
Set mBaseStyle = baseStyle
If Not LocalStyle Is Nothing Then Set mLocalStyle = LocalStyle.clone ' local style cannot be shared
mIsSelectable = IsSelectable

setBarSemiWidth
calcBoundingRect
End Sub

Public Sub resetStyle()
' reset the style to the initial style (ie the style for the bar series) with no
' local modifications
Dim failpoint As Long
On Error GoTo Err

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
mGOSP.ObjectChanged mHandle, mBoundingRect

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "resetStyle" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


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

Private Sub Draw(ByRef area As TRectangle)
If Not area.isValid Then
    drawBar IIf(upBar, upColor, downColor), _
            IIf(barColor <> -1, barColor, IIf(upBar, upColor, downColor))
Else
    drawPartialBar area, _
                    IIf(upBar, upColor, downColor), _
                    IIf(barColor <> -1 And outlineThickness <> 0, _
                        barColor, _
                        IIf(upBar, upColor, downColor)), _
                    IIf(barColor <> -1, _
                        barColor, _
                        IIf(upBar, upColor, downColor))
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
        mCanvas.drawLine mCanvas.newPoint(mX - mBarSemiWidth, mOpen), _
                        mCanvas.newPoint(mX, mOpen)
        mCanvas.drawLine mCanvas.newPoint(mX, mLow), _
                        mCanvas.newPoint(mX, mHigh)
        mCanvas.drawLine mCanvas.newPoint(mX, mClose), _
                        mCanvas.newPoint(mX + mBarSemiWidth + onePixel, mClose)
    Else
        mCanvas.setPenAttributes Color, _
                                1, _
                                LineStyles.LineInsideSolid, _
                                DrawModes.DrawModeCopyPen
        barThicknessScaleX = mCanvas.convertPixelsToLogicalX(barThickness - 1)
        barThicknessScaleY = mCanvas.convertPixelsToLogicalY(barThickness - 1)
        
        If (mOpen - barThicknessScaleY / 2) < mLow Then
            mCanvas.drawRectangle mCanvas.newPoint(mX - mBarSemiWidth, mLow), _
                                    mCanvas.newPoint(mX + barThicknessScaleX / 2, mLow + barThicknessScaleY)
        ElseIf (mOpen + barThicknessScaleY / 2) > mHigh Then
            mCanvas.drawRectangle mCanvas.newPoint(mX - mBarSemiWidth, mHigh - barThicknessScaleY), _
                                    mCanvas.newPoint(mX + barThicknessScaleX / 2, mHigh)
        Else
            mCanvas.drawRectangle mCanvas.newPoint(mX - mBarSemiWidth, mOpen - barThicknessScaleY / 2), _
                                    mCanvas.newPoint(mX + barThicknessScaleX / 2, mOpen + barThicknessScaleY / 2)
        End If
        
        mCanvas.drawRectangle mCanvas.newPoint(mX - barThicknessScaleX / 2, mLow), _
                                mCanvas.newPoint(mX + barThicknessScaleX / 2, mHigh)
        
        If (mClose - barThicknessScaleY / 2) < mLow Then
            mCanvas.drawRectangle mCanvas.newPoint(mX - barThicknessScaleX / 2, mLow), _
                                    mCanvas.newPoint(mX + mBarSemiWidth, mLow + barThicknessScaleY)
        ElseIf (mClose + barThicknessScaleY / 2) > mHigh Then
            mCanvas.drawRectangle mCanvas.newPoint(mX - barThicknessScaleX / 2, mHigh - barThicknessScaleY), _
                                    mCanvas.newPoint(mX + mBarSemiWidth, mHigh)
        Else
            mCanvas.drawRectangle mCanvas.newPoint(mX - barThicknessScaleX / 2, mClose - barThicknessScaleY / 2), _
                                    mCanvas.newPoint(mX + mBarSemiWidth, mClose + barThicknessScaleY / 2)
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
Dim lOutlineThickness As Long

lOutlineThickness = IIf(outlineThickness = 0, 1, outlineThickness)

If upBar And Not solidUpBody Then
    mCanvas.setBrushAttributes mCanvas.BackColor, FillStyles.FillSolid
    barOutlineScaleX = mCanvas.convertPixelsToLogicalX(lOutlineThickness - 1)
    barOutlineScaleY = mCanvas.convertPixelsToLogicalY(lOutlineThickness - 1)
    
    If bottom = mOpen And top = mClose Then
        ' draw the whole of the body
        mCanvas.setPenAttributes outlineColor, _
                                lOutlineThickness, _
                                LineStyles.LineInsideSolid, _
                                DrawModes.DrawModeCopyPen
        mCanvas.drawRectangle mCanvas.newPoint(mX - mBarSemiWidth, bottom), _
                                mCanvas.newPoint(mX + mBarSemiWidth, top)
    Else
        
        ' first blank out the area
        mCanvas.setPenAttributes mCanvas.BackColor, _
                                1, _
                                LineStyles.LineInsideSolid, _
                                DrawModes.DrawModeCopyPen
        
        mCanvas.drawRectangle mCanvas.newPoint(mX - mBarSemiWidth, bottom), _
                        mCanvas.newPoint(mX + mBarSemiWidth, top)
        
        mCanvas.setPenAttributes Color, _
                                1, _
                                LineStyles.LineInsideSolid, _
                                DrawModes.DrawModeCopyPen
        ' we'll draw the outlines of the bar as filled boxes
        ' so set the fill color to the outline color
        mCanvas.setBrushAttributes IIf(Color <> -1, Color, outlineColor), FillStyles.FillSolid
        
        If bottom = mOpen Then
            ' draw the base of the body
            If lOutlineThickness = 1 Then
                mCanvas.drawLine mCanvas.newPoint(mX - mBarSemiWidth, bottom), _
                                mCanvas.newPoint(mX + mBarSemiWidth, bottom)
            Else
                mCanvas.drawRectangle mCanvas.newPoint(mX - mBarSemiWidth, bottom), _
                                mCanvas.newPoint(mX + mBarSemiWidth, bottom + barOutlineScaleY)
            End If
        End If
        If top = mClose Then
            ' draw the top of the body
            If lOutlineThickness = 1 Then
                mCanvas.drawLine mCanvas.newPoint(mX - mBarSemiWidth, top), _
                                mCanvas.newPoint(mX + mBarSemiWidth, top)
            Else
                mCanvas.drawRectangle mCanvas.newPoint(mX - mBarSemiWidth, top - barOutlineScaleY), _
                                mCanvas.newPoint(mX + mBarSemiWidth, top)
            End If
        End If
        If lOutlineThickness = 1 Then
            mCanvas.drawLine mCanvas.newPoint(mX - mBarSemiWidth, bottom), _
                            mCanvas.newPoint(mX - mBarSemiWidth, top)
            mCanvas.drawLine mCanvas.newPoint(mX + mBarSemiWidth, bottom), _
                            mCanvas.newPoint(mX + mBarSemiWidth, top)
        Else
            mCanvas.drawRectangle mCanvas.newPoint(mX - mBarSemiWidth, bottom), _
                                mCanvas.newPoint(mX - mBarSemiWidth + barOutlineScaleX, top)
            mCanvas.drawRectangle mCanvas.newPoint(mX + mBarSemiWidth - barOutlineScaleX, bottom), _
                                mCanvas.newPoint(mX + mBarSemiWidth, top)
        End If
    End If
Else
    mCanvas.setBrushAttributes Color, FillStyles.FillSolid
    mCanvas.setPenAttributes outlineColor, _
                            lOutlineThickness, _
                            LineStyles.LineInsideSolid, _
                            DrawModes.DrawModeCopyPen
    mCanvas.drawRectangle mCanvas.newPoint(mX - mBarSemiWidth, bottom), _
                        mCanvas.newPoint(mX + mBarSemiWidth, top)
End If
End Sub

Friend Sub drawPartialBar( _
                ByRef area As TRectangle, _
                ByVal Color As Long, _
                ByVal outlineColor As Long, _
                ByVal tailColor As Long)
Dim barThicknessScaleX As Double
Dim barThicknessScaleY As Double
If Not mVisible Then Exit Sub
If blank Then Exit Sub

If displayMode = BarDisplayModeCandlestick Then
    If downBar Then
        If area.top < mClose Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), area.top, tailColor
        ElseIf area.top >= mClose And area.top <= mOpen And area.bottom < mClose Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), mClose, tailColor
            drawBody mClose, area.top, Color, outlineColor
        ElseIf area.top > mOpen And area.bottom < mClose Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), mClose, tailColor
            drawBody mClose, mOpen, Color, outlineColor
            drawTail mOpen, IIf(mHigh > area.top, area.top, mHigh), tailColor
        ElseIf area.top <= mOpen And area.bottom >= mClose Then
            drawBody area.bottom, area.top, Color, outlineColor
        ElseIf area.top > mOpen And area.bottom <= mOpen And area.bottom >= mClose Then
            drawTail mOpen, IIf(mHigh > area.top, area.top, mHigh), tailColor
            drawBody area.bottom, mOpen, Color, outlineColor
        Else
            drawTail area.bottom, IIf(mHigh > area.top, area.top, mHigh), tailColor
        End If
    Else
        If area.top < mOpen Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), area.top, tailColor
        ElseIf area.top >= mOpen And area.top <= mClose And area.bottom < mOpen Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), mOpen, tailColor
            drawBody mOpen, area.top, Color, outlineColor
        ElseIf area.top > mClose And area.bottom < mOpen Then
            drawTail IIf(mLow < area.bottom, area.bottom, mLow), mOpen, tailColor
            drawBody mOpen, mClose, Color, outlineColor
            drawTail mClose, IIf(mHigh > area.top, area.top, mHigh), tailColor
        ElseIf area.top <= mClose And area.bottom >= mOpen Then
            drawBody area.bottom, area.top, Color, outlineColor
        ElseIf area.top > mClose And area.bottom <= mClose And area.bottom >= mOpen Then
            drawTail mClose, IIf(mHigh > area.top, area.top, mHigh), tailColor
            drawBody area.bottom, mClose, Color, outlineColor
        Else
            drawTail area.bottom, IIf(mHigh > area.top, area.top, mHigh), tailColor
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
            mCanvas.drawLine mCanvas.newPoint(mX - mBarSemiWidth, mOpen), _
                            mCanvas.newPoint(mX, mOpen)
        End If
        If mClose >= area.bottom And mClose <= area.top Then
            mCanvas.drawLine mCanvas.newPoint(mX, mClose), _
                            mCanvas.newPoint(mX + mBarSemiWidth + onePixel, mClose)
        End If
        mCanvas.drawLine mCanvas.newPoint(mX, IIf(area.bottom > mLow, area.bottom, mLow)), _
                        mCanvas.newPoint(mX, IIf(area.top < mHigh, area.top, mHigh))
    Else
        mCanvas.setBrushAttributes outlineColor, FillStyles.FillSolid
        barThicknessScaleX = mCanvas.convertPixelsToLogicalX(barThickness - 1)
        barThicknessScaleY = mCanvas.convertPixelsToLogicalY(barThickness - 1)
        
        If mOpen >= area.bottom And mOpen <= area.top Then
            If (mOpen - barThicknessScaleY / 2) < mLow Then
                mCanvas.drawRectangle mCanvas.newPoint(mX - mBarSemiWidth, mLow), _
                                mCanvas.newPoint(mX + barThicknessScaleX / 2, mLow + barThicknessScaleY)
            ElseIf (mOpen + barThicknessScaleY / 2) > mHigh Then
                mCanvas.drawRectangle mCanvas.newPoint(mX - mBarSemiWidth, mHigh - barThicknessScaleY), _
                                mCanvas.newPoint(mX + barThicknessScaleX / 2, mHigh)
            Else
                mCanvas.drawRectangle mCanvas.newPoint(mX - mBarSemiWidth, mOpen - barThicknessScaleY / 2), _
                                mCanvas.newPoint(mX + barThicknessScaleX / 2, mOpen + barThicknessScaleY / 2)
            End If
        End If
        
        If mClose >= area.bottom And mClose <= area.top Then
            If (mClose - barThicknessScaleY / 2) < mLow Then
                mCanvas.drawRectangle mCanvas.newPoint(mX - barThicknessScaleX / 2, mLow), _
                                mCanvas.newPoint(mX + mBarSemiWidth, mLow + barThicknessScaleY)
            ElseIf (mClose + barThicknessScaleY / 2) > mHigh Then
                mCanvas.drawRectangle mCanvas.newPoint(mX - barThicknessScaleX / 2, mHigh - barThicknessScaleY), _
                                mCanvas.newPoint(mX + mBarSemiWidth, mHigh)
            Else
                mCanvas.drawRectangle mCanvas.newPoint(mX - barThicknessScaleX / 2, mClose - barThicknessScaleY / 2), _
                                mCanvas.newPoint(mX + mBarSemiWidth, mClose + barThicknessScaleY / 2)
            End If
        End If
        
        mCanvas.drawRectangle mCanvas.newPoint(mX - barThicknessScaleX / 2, IIf(area.bottom > mLow, area.bottom, mLow)), _
                        mCanvas.newPoint(mX + barThicknessScaleX / 2, IIf(area.top < mHigh, area.top, mHigh))
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
    mCanvas.drawRectangle mCanvas.newPoint(mX - tailThicknessScaleX / 2, bottom), _
                    mCanvas.newPoint(mX + tailThicknessScaleX / 2, top)
Else
    mCanvas.drawRectangle mCanvas.newPoint(mX, bottom), mCanvas.newPoint(mX, top)
End If
End Sub

Private Sub setBarSemiWidth()
Select Case displayMode
Case BarDisplayModeBar
'    mBarSemiWidth = (2 * barWidth - 0.2) / 2
'    If mBarSemiWidth > 0.5 Then mBarSemiWidth = 0.5
    mBarSemiWidth = barWidth / 2
Case BarDisplayModeCandlestick
    mBarSemiWidth = barWidth / 2
Case BarDisplayModeLine
    mBarSemiWidth = barWidth / 2
End Select
End Sub

Public Sub tick(ByVal Price As Double)
Dim clearRect As TRectangle

If Price = mClose Then Exit Sub

'If Not mBlank Then
'    With clearRect
'        .isValid = False
'        If upBar Then
'            If Price < mOpen Then
'                ' is now a downbar so redraw the whole thing
'                .left = mX - mBarSemiWidth
'                .right = mX + mBarSemiWidth
'                .bottom = IIf(Price < mLow, Price, mLow)
'                .top = mHigh
'                .isValid = True
'            ElseIf Price < mClose Then
'                .left = mX - mBarSemiWidth
'                .right = mX + mBarSemiWidth
'                .bottom = Price
'                .top = mClose
'                .isValid = True
'            ElseIf Price > mClose Then
'                .left = mX - mBarSemiWidth
'                .right = mX + mBarSemiWidth
'                .bottom = mClose
'                .top = Price
'                .isValid = True
'            End If
'        Else
'            If Price >= mOpen Then
'                ' now an upbar so redraw the whole thing
'                .left = mX - mBarSemiWidth
'                .right = mX + mBarSemiWidth
'                .bottom = mLow
'                .top = IIf(Price > mHigh, Price, mHigh)
'                .isValid = True
'            ElseIf Price > mClose Then
'                .left = mX - mBarSemiWidth
'                .right = mX + mBarSemiWidth
'                .bottom = mClose
'                .top = Price
'                .isValid = True
'            ElseIf Price < mClose Then
'                .left = mX - mBarSemiWidth
'                .right = mX + mBarSemiWidth
'                .bottom = Price
'                .top = mClose
'                .isValid = True
'            End If
'        End If
'        If .isValid Then undraw clearRect
'    End With
'End If

undraw mBoundingRect

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

mGOSP.ObjectChanged mHandle, mBoundingRect
'mGOSP.ObjectChanged mHandle, clearRect

End Sub

Private Sub undraw(ByRef area As TRectangle)
If Not mVisible Then Exit Sub
If Not area.isValid Then
    mGOSP.ObjectUndrawn mHandle, mBoundingRect
Else
    mGOSP.ObjectUndrawn mHandle, area
End If
End Sub



