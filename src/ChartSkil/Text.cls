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

Private Const ProjectName                As String = "ChartSkil26"
Private Const ModuleName                As String = "Text"

'================================================================================
' Member variables and constants
'================================================================================

Private mCanvas As canvas
Private mLayer As Long
Private mHandle As Long

Private mLocalStyle As TextStyle
Private mBaseStyle As TextStyle

Private mInScope As Boolean
Private mVisible As Boolean

Private mEventProxy As ChartRegionEventProxy

Private mNoDraw As Boolean

Private mText As String

Private mRawBoundingRect As TRectangle
Private mBoundingRect As TRectangle ' the rectangle that surrounds the box (if
                                ' any - including the thickness of the lines) or
                                ' the text if no box
Private mTextRect As TRectangle ' the rectangle containing the text
Private mBoxRect As TRectangle  ' the rectangle used to draw the box

Private mIsSetFont As Boolean
Private mIsSetColor As Boolean
Private mIsSetBox As Boolean
Private mIsSetBoxColor As Boolean
Private mIsSetBoxStyle As Boolean
Private mIsSetBoxThickness As Boolean
Private mIsSetBoxFillColor As Boolean
Private mIsSetBoxFillStyle As Boolean
Private mIsSetAlign As Boolean
Private mIsSetFixedX As Boolean
Private mIsSetFixedY As Boolean
Private mIsSetIncludeInAutoscale As Boolean
Private mIsSetExtended As Boolean
Private mIsSetPaddingX As Boolean
Private mIsSetPaddingY As Boolean

Private mFont As StdFont
Private mColor As Long
Private mBox As Boolean
Private mBoxColor As Long
Private mBoxStyle As LineStyles
Private mBoxThickness As Long
Private mBoxFillColor As Long
Private mBoxFillStyle As FillStyles
Private mAlign As TextAlignModes
Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean
Private mPaddingX As Double
Private mPaddingY As Double

Private mBlank As Boolean
Private mDrawn As Boolean

Private mPosition As Point
Private mOffset As Dimension

Private mKey As String

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

Dim failpoint As Long
On Error GoTo Err

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

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_boundingRectangle" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Property

Private Function IGraphicObject_boundingRectanglesAt(ByVal periodNumber As Long) As ChartSkilTypes.TRectangle()
'
End Function

Private Property Get IGraphicObject_capabilities() As Long
IGraphicObject_capabilities = capabilities.BlockUndraw
End Property

Private Sub IGraphicObject_draw( _
                ByRef areas() As TRectangle, _
                ByVal hdc As Long)
draw
End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = extended Or fixedX
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
IGraphicObject_noDraw = mBlank Or mNoDraw
End Property

Private Property Get IGraphicObject_periodNumber() As Long
If extended Or Not mBoundingRect.isValid Then
    IGraphicObject_periodNumber = -1
Else
    Select Case align
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

Public Property Get align() As TextAlignModes
align = mBaseStyle.align
If Not mLocalStyle Is Nothing Then align = mLocalStyle.align
If mIsSetAlign Then align = mAlign
End Property

Public Property Let align(ByVal value As TextAlignModes)
If mAlign = value And mIsSetAlign Then Exit Property
undraw
mAlign = value
mIsSetAlign = True
update
End Property

Public Property Get box() As Boolean
box = mBaseStyle.box
If Not mLocalStyle Is Nothing Then box = mLocalStyle.box
If mIsSetBox Then box = mBox
End Property

Public Property Let box(ByVal value As Boolean)
If mBox = value And mIsSetBox Then Exit Property
undraw
mBox = value
mIsSetBox = True
update
End Property

Public Property Get boxColor() As Long
boxColor = mBaseStyle.boxColor
If Not mLocalStyle Is Nothing Then boxColor = mLocalStyle.boxColor
If mIsSetBoxColor Then boxColor = mBoxColor
End Property

Public Property Let boxColor(ByVal value As Long)
If mBoxColor = value And mIsSetBoxColor Then Exit Property
mBoxColor = value
mIsSetBoxColor = True
update
End Property

Public Property Get boxFillColor() As Long
boxFillColor = mBaseStyle.boxFillColor
If Not mLocalStyle Is Nothing Then boxFillColor = mLocalStyle.boxFillColor
If mIsSetBoxFillColor Then boxFillColor = mBoxFillColor
End Property

Public Property Let boxFillColor(ByVal value As Long)
If mBoxFillColor = value And mIsSetBoxFillColor Then Exit Property
mBoxFillColor = value
mIsSetBoxFillColor = True
update
End Property

Public Property Get boxFillStyle() As FillStyles
boxFillStyle = mBaseStyle.boxFillStyle
If Not mLocalStyle Is Nothing Then boxFillStyle = mLocalStyle.boxFillStyle
If mIsSetBoxFillStyle Then boxFillStyle = mBoxFillStyle
End Property

Public Property Let boxFillStyle(ByVal value As FillStyles)
If mBoxFillStyle = value And mIsSetBoxFillStyle Then Exit Property
undraw
mBoxFillStyle = value
mIsSetBoxFillStyle = True
update
End Property

Public Property Get boxStyle() As LineStyles
boxStyle = mBaseStyle.boxStyle
If Not mLocalStyle Is Nothing Then boxStyle = mLocalStyle.boxStyle
If mIsSetBoxStyle Then boxStyle = mBoxStyle
End Property

Public Property Let boxStyle(ByVal value As LineStyles)
If mBoxStyle = value And mIsSetBoxStyle Then Exit Property
undraw
mBoxStyle = value
mIsSetBoxStyle = True
update
End Property

Public Property Get boxThickness() As Long
boxThickness = mBaseStyle.boxThickness
If Not mLocalStyle Is Nothing Then boxThickness = mLocalStyle.boxThickness
If mIsSetBoxThickness Then boxThickness = mBoxThickness
End Property

Public Property Let boxThickness(ByVal value As Long)
If mBoxThickness = value And mIsSetBoxThickness Then Exit Property
undraw
mBoxThickness = value
mIsSetBoxThickness = True
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

Public Property Get fixedX() As Boolean
fixedX = mBaseStyle.fixedX
If Not mLocalStyle Is Nothing Then fixedX = mLocalStyle.fixedX
If mIsSetFixedX Then fixedX = mFixedX
End Property

Public Property Let fixedX(ByVal value As Boolean)
If mFixedX = value And mIsSetFixedX Then Exit Property
undraw
mFixedX = value
mIsSetFixedX = True
convertPositionX
update
End Property

Public Property Get fixedY() As Boolean
fixedY = mBaseStyle.fixedY
If Not mLocalStyle Is Nothing Then fixedY = mLocalStyle.fixedY
If mIsSetFixedY Then fixedY = mFixedY
End Property

Public Property Let fixedY(ByVal value As Boolean)
If mFixedY = value And mIsSetFixedY Then Exit Property
undraw
mFixedY = value
mIsSetFixedY = True
convertPositionY
update
End Property

Public Property Get font() As StdFont
If mIsSetFont Then
    Set font = gCloneFont(mFont)
ElseIf Not mLocalStyle Is Nothing Then
    Set font = gCloneFont(mBaseStyle.font)
Else
    Set font = gCloneFont(mBaseStyle.font)
End If
End Property

Public Property Let font(ByVal value As StdFont)
If value Is Nothing Then
    If mIsSetFont Then
        undraw
        Set mFont = Nothing
        mIsSetFont = False
        update
    End If
Else
    undraw
    Set mFont = gCloneFont(value)
    mIsSetFont = True
    update
End If
End Property

Friend Property Get handle() As Long
handle = mHandle
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
If mLayer = value Then Exit Property
undraw
mLayer = value
update
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Get localStyle() As TextStyle
Set localStyle = mLocalStyle.clone
End Property

Public Property Let localStyle( _
                ByVal value As TextStyle)

undraw

Set mLocalStyle = value.clone
calcCurrentBoundingRect
mEventProxy.objectChanged mHandle, mBoundingRect

End Property

Public Property Let paddingX(ByVal value As Double)
If mPaddingX = value And mIsSetPaddingX Then Exit Property
undraw
mPaddingX = value
mIsSetPaddingX = True
update
End Property

Public Property Get paddingX() As Double
paddingX = mBaseStyle.paddingX
If Not mLocalStyle Is Nothing Then paddingX = mLocalStyle.paddingX
If mIsSetPaddingX Then paddingX = mPaddingX
End Property

Public Property Let paddingY(ByVal value As Double)
If mPaddingY = value And mIsSetPaddingY Then Exit Property
undraw
mPaddingY = value
mIsSetPaddingY = True
update
End Property

Public Property Get paddingY() As Double
paddingY = mBaseStyle.paddingY
If Not mLocalStyle Is Nothing Then paddingY = mLocalStyle.paddingY
If mIsSetPaddingY Then paddingY = mPaddingY
End Property

Public Property Get position() As Point
Set position = mPosition
End Property

Public Property Let position(ByVal value As Point)
If value Is Nothing Then
    undraw
    Set mPosition = Nothing
    mBoundingRect.isValid = False
    mBlank = True
    Exit Property
ElseIf Not value.isAssigned Then
    undraw
    Set mPosition = Nothing
    mBoundingRect.isValid = False
    mBlank = True
    Exit Property
ElseIf Not mPosition Is Nothing Then
    If value.equals(mPosition) Then Exit Property
End If
undraw
Set mPosition = value
If mText <> "" Then mBlank = False
If mPosition.CoordinateSystemX <> CoordsLogical Then fixedX = True
If mPosition.CoordinateSystemY <> CoordsLogical Then fixedY = True
convertPosition
update
End Property

Public Property Get offset() As Dimension
Set offset = mOffset
End Property

Public Property Let offset(ByVal value As Dimension)
If value Is Nothing Then
    undraw
    Set mOffset = Nothing
ElseIf Not value.isAssigned Then
    undraw
    Set mOffset = Nothing
ElseIf Not mOffset Is Nothing Then
    If value.equals(mOffset) Then Exit Property
Else
    undraw
    Set mOffset = value
End If
update
End Property

Public Property Let text(value As String)
If value = mText Then Exit Property
undraw  ' remove current text
mText = value
mBlank = (value = "" Or mPosition Is Nothing)

update
End Property

Public Property Get text() As String
text = mText
End Property

Public Property Get visible() As Boolean
visible = mVisible
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub draw()
Dim failpoint As Long
On Error GoTo Err

If mBlank Then Exit Sub
If Not mVisible Then Exit Sub
If Not mPosition.isAssigned Then Exit Sub
mDrawn = True

With mCanvas
    If box Then
        .setPenAttributes boxColor, _
                        boxThickness, _
                        boxStyle, _
                        DrawModes.DrawModeCopyPen
        .setBrushAttributes boxFillColor, boxFillStyle
        .drawRectangle .newPoint(mBoxRect.left, mBoxRect.bottom), .newPoint(mBoxRect.right, mBoxRect.top)
    End If
    
    .setTextAttributes Color, font, True
'    If mText = "Volume" Then Stop
    .drawText mText, .newPoint(mTextRect.left, mTextRect.top)
End With

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "draw" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


    
End Sub

Friend Sub initialise( _
                ByVal key As String, _
                ByVal eventProxy As ChartRegionEventProxy, _
                ByVal layer As Long, _
                ByVal canvas As canvas, _
                ByVal baseStyle As TextStyle, _
                ByVal localStyle As TextStyle)
                
Dim failpoint As Long
On Error GoTo Err

mKey = key
Set mEventProxy = eventProxy
mLayer = layer
Set mCanvas = canvas
Set mBaseStyle = baseStyle
If Not localStyle Is Nothing Then Set mLocalStyle = localStyle.clone ' local style cannot be shared
calcCurrentBoundingRect

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "initialise" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcRawBoundingRect()

' NB: calculates the raw dimensions of the bounding box in twips, because
' the scale may not have been finalised when this is called, and is anyway
' subject to change
If mText = "" Then Exit Sub

mCanvas.setTextAttributes Color, font, True
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

If mPosition Is Nothing Then
    mBoundingRect.isValid = False
    Exit Sub
End If

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

If box Then
    ' make the inside of the box 1mm wider and taller than the text, plus the thickness of the lines
    
    Dim padding As Dimension
    Set padding = mCanvas.newDimension((mPaddingX + 0.5) / 10, (IIf(mPaddingY > 0.5, mPaddingY - 0.5, 0)) / 10)
    XIncrement = padding.XLogical
    YIncrement = padding.YLogical
    
    If boxThickness > 1 Then
        XLineThickness = mCanvas.convertPixelsToLogicalX(mBoxThickness - 1)
        YLineThickness = mCanvas.convertPixelsToLogicalY(mBoxThickness - 1)
        If boxStyle = LineInsideSolid Then
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
    
    If boxThickness > 1 And boxStyle <> LineInsideSolid Then
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

    Select Case align
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
        Select Case align
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
If fixedX Then
    If mPosition.CoordinateSystemX = CoordsLogical Then
        mPosition.CoordinateSystemX = CoordsDistance
    End If
Else
    If mPosition.CoordinateSystemX <> CoordsLogical Then
        mPosition.CoordinateSystemX = CoordsLogical
    End If
End If
End Sub

Private Sub convertPositionY()
If mPosition Is Nothing Then Exit Sub
If fixedY Then
    If mPosition.CoordinateSystemY = CoordsLogical Then
        mPosition.CoordinateSystemY = CoordsDistance
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

'With mCanvas
'    .setPenAttributes .backColor, _
'                    1, _
'                    LineStyles.LineInsideSolid, _
'                    DrawModes.DrawModeCopyPen
'    .setBrushAttributes .backColor, FillStyles.FillSolid
'    .drawRectangle mBoundingRect.left, mBoundingRect.bottom, mBoundingRect.right, mBoundingRect.top
'End With
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
End Sub
