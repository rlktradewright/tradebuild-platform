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

Private mSurface As PictureBox
Private mLayer As Long
Private mHandle As Long

Private mInScope As Boolean
Private mVisible As Boolean

Private mEventProxy As ChartRegionEventProxy

Private mBoundingRect As TRectangle

Private mHistBarWidth As Single

Private mIncludeInAutoscale  As Boolean

Private mPeriodNumber  As Long
Private mDataValue As Double
Private mBlank As Boolean   ' indicates that no value has been assigned
Private mLineThickness As Long
Private mLineColour As Long
Private mLineStyle As LineStyles
Private mDisplayMode As DisplayModes

Private mPrevDataPoint As DataPoint

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
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
If gaugeChanged Then
    mBoundingRect = calcBoundingRect
End If
IGraphicObject_boundingRectangle = mBoundingRect
End Property

Private Function IGraphicObject_boundingRectanglesAt(ByVal periodNumber As Long) As ChartSkilTypes.TRectangle()
'
End Function

Private Property Get IGraphicObject_capabilities() As Long
IGraphicObject_capabilities = capabilities.PartialRedraw
End Property

Private Sub IGraphicObject_draw(ByRef areas() As TRectangle)
draw
End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = False
End Property

Private Property Get IGraphicObject_gaugeDependent() As Boolean
Select Case mDisplayMode
Case displayAsPoints
    IGraphicObject_gaugeDependent = True
Case DisplayAsLines
    IGraphicObject_gaugeDependent = True
Case DisplayAsSteppedLines
    IGraphicObject_gaugeDependent = True
Case displayAsHistogram
    IGraphicObject_gaugeDependent = False
End Select
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
IGraphicObject_periodNumber = mPeriodNumber
End Property

Private Property Get IGraphicObject_scaleDependent() As Boolean
IGraphicObject_scaleDependent = False
End Property

Private Sub IGraphicObject_undraw(ByRef area As TRectangle)
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

Friend Property Let displayMode(ByVal value As DisplayModes)
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

Friend Property Let layer(ByVal value As Long)
mLayer = value
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Let lineColour(ByVal value As Long)
mLineColour = value
End Property

Friend Property Let lineStyle(ByVal value As LineStyles)
mLineStyle = value
mSurface.DrawStyle = value
End Property

Friend Property Let lineThickness(ByVal value As Long)
mLineThickness = value
mSurface.DrawWidth = value
calcBoundingRect
End Property

Public Property Get inScope() As Boolean
inScope = mInScope
End Property

Public Property Get periodNumber() As Long
periodNumber = mPeriodNumber
End Property

Public Property Let periodNumber(ByVal newValue As Long)
mPeriodNumber = newValue
End Property

Public Property Get prevDataPoint() As DataPoint
Set prevDataPoint = mPrevDataPoint
End Property

Public Property Let prevDataPoint(ByVal newValue As DataPoint)
Set mPrevDataPoint = newValue
End Property

Friend Property Let surface(ByVal value As PictureBox)
Set mSurface = value
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

Private Function calcBoundingRect() As TRectangle
Dim boundingRect As TRectangle
Dim thicknessScaleX As Double
Dim thicknessScaleY As Double
Dim prevPeriod As Long
Dim prevValue As Double

thicknessScaleX = Abs((((mLineThickness + 2) * Screen.TwipsPerPixelX) / mSurface.Width) * mSurface.scaleWidth)
thicknessScaleY = Abs((((mLineThickness + 2) * Screen.TwipsPerPixelY) / mSurface.Height) * mSurface.scaleHeight)

If mDisplayMode = DisplayModes.displayAsHistogram Then
    With boundingRect
        .left = mPeriodNumber - mHistBarWidth / 2
        .right = mPeriodNumber + mHistBarWidth / 2
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
        .left = mPeriodNumber - thicknessScaleX / 2
        .right = mPeriodNumber + thicknessScaleX / 2
        .bottom = mDataValue - thicknessScaleY / 2
        .top = mDataValue + thicknessScaleY / 2
    End With
ElseIf mDisplayMode = DisplayModes.DisplayAsLines Then
    With mPrevDataPoint
        prevPeriod = .periodNumber
        prevValue = .dataValue
    End With
    With boundingRect
        .left = prevPeriod - thicknessScaleX / 2
        .right = mPeriodNumber + thicknessScaleX / 2
        If mDataValue > prevValue Then
            .bottom = prevValue - thicknessScaleY / 2
            .top = mDataValue + thicknessScaleY / 2
        Else
            .bottom = mDataValue - thicknessScaleY / 2
            .top = prevValue + thicknessScaleY / 2
        End If
    End With
ElseIf mDisplayMode = DisplayModes.DisplayAsSteppedLines Then
    With mPrevDataPoint
        prevPeriod = .periodNumber
        prevValue = .dataValue
    End With
    With boundingRect
        .left = prevPeriod - thicknessScaleX / 2
        .right = mPeriodNumber + thicknessScaleX / 2
        If mDataValue > prevValue Then
            .bottom = prevValue - thicknessScaleY / 2
            .top = mDataValue + thicknessScaleY / 2
        Else
            .bottom = mDataValue - thicknessScaleY / 2
            .top = prevValue + thicknessScaleY / 2
        End If
    End With
End If
boundingRect.isvalid = True
calcBoundingRect = boundingRect
End Function

Private Sub draw()
drawPoint mLineColour
End Sub

Private Function drawAsPoint() As Boolean
If mPrevDataPoint Is Nothing Then
    drawAsPoint = True
ElseIf mDisplayMode = DisplayModes.displayAsPoints Then
'ElseIf mDisplayMode = DisplayModes.displayAsPoints Or _
'    (mDisplayMode = DisplayModes.DisplayAsLines And Not mPrevDataPoint.inScope) Or _
'    (mDisplayMode = DisplayModes.DisplayAsSteppedLines And Not mPrevDataPoint.inScope) _
'Then
    drawAsPoint = True
Else
    drawAsPoint = False
End If
End Function

Private Sub drawPoint(ByVal colour As Long)
If blank Then Exit Sub
If Not mVisible Then Exit Sub
mSurface.DrawWidth = mLineThickness
mSurface.DrawStyle = mLineStyle
mSurface.DrawMode = vbCopyPen
If mDisplayMode = DisplayModes.displayAsHistogram Then
    mSurface.Line (mPeriodNumber - mHistBarWidth / 2, 0)-(mPeriodNumber + mHistBarWidth / 2, mDataValue), colour, BF
ElseIf drawAsPoint Then
    mSurface.PSet (mPeriodNumber, mDataValue), colour
ElseIf mDisplayMode = DisplayModes.DisplayAsLines Then
    mSurface.Line (mPrevDataPoint.periodNumber, mPrevDataPoint.dataValue)-(mPeriodNumber, mDataValue), colour
ElseIf mDisplayMode = DisplayModes.DisplayAsSteppedLines Then
    mSurface.Line (mPrevDataPoint.periodNumber, mPrevDataPoint.dataValue)-(mPeriodNumber - 0.5, mPrevDataPoint.dataValue), colour
    mSurface.Line (periodNumber - 0.5, mPrevDataPoint.dataValue)-(mPeriodNumber - 0.5, mDataValue), colour
    mSurface.Line (periodNumber - 0.5, mDataValue)-(mPeriodNumber, mDataValue), colour
End If
End Sub

Private Sub undraw()
If Not mVisible Then Exit Sub
drawPoint mSurface.backColor
mEventProxy.objectUndrawn mHandle, mBoundingRect
End Sub

Private Sub update()
If mHandle = 0 Then
    ' haven't yet been added to chart region
    Exit Sub
End If
mBoundingRect = calcBoundingRect
mEventProxy.objectChanged mHandle, mBoundingRect
If mIncludeInAutoscale And mBoundingRect.isvalid Then
    If mBoundingRect.top >= mSurface.scaleTop Or mBoundingRect.bottom <= (mSurface.scaleTop + mSurface.scaleHeight) Then
        mEventProxy.boundsExceeded
    End If
End If
End Sub

