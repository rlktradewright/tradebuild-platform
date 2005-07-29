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

Private mSurface As PictureBox
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
Private mColour As Long
Private mBox As Boolean
Private mBoxColour As Long
Private mBoxStyle As LineStyles
Private mBoxThickness As Long
Private mBoxFillColour As Long
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

mRawBoundingRect.isvalid = False
mBoundingRect.isvalid = False
mTextRect.isvalid = False
mBoxRect.isvalid = False
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
'If mText = "Started here" Then Stop
If gaugeChanged Then
    'If mText = "Started here" Then Stop
    calcRawBoundingRect
End If

If gaugeChanged Or _
    (mFixedX And xBoundsChanged) Or _
    (mFixedY And yBoundsChanged) _
Then
    calcCurrentBoundingRect
End If

IGraphicObject_boundingRectangle = mBoundingRect
End Property

Private Function IGraphicObject_boundingRectanglesAt(ByVal periodNumber As Long) As ChartSkilTypes.TRectangle()
'
End Function

Private Property Get IGraphicObject_capabilities() As Long
IGraphicObject_capabilities = capabilities.blockUndraw
End Property

Private Sub IGraphicObject_draw(ByRef areas() As TRectangle)
'If mText = "Started here" Then Stop
draw
End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = mExtended Or mFixedX
End Property

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
If mExtended Or Not mBoundingRect.isvalid Then
    IGraphicObject_periodNumber = -1
Else
    Select Case mAlign
    Case AlignTopLeft, AlignCentreLeft, AlignBottomLeft
        IGraphicObject_periodNumber = Int(mBoundingRect.left + 0.5)
    Case AlignTopCentre, AlignCentreCentre, AlignBottomCentre
        IGraphicObject_periodNumber = Int((mBoundingRect.left + mBoundingRect.right) / 2 + 0.5)
    Case AlignTopRight, AlignCentreRight, AlignBottomRight
        IGraphicObject_periodNumber = Int(mBoundingRect.right + 0.5)
    End Select
End If
End Property

Private Property Get IGraphicObject_scaleDependent() As Boolean
IGraphicObject_scaleDependent = mFixedX Or mFixedY
End Property

Private Sub IGraphicObject_undraw(ByRef area As TRectangle)
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
boxColor = mBoxColour
End Property

Public Property Let boxColor(ByVal newValue As Long)
If mBoxColour = newValue Then Exit Property
undraw
mBoxColour = newValue
update
End Property

Friend Property Get boxFillColor() As Long
boxFillColor = mBoxFillColour
End Property

Public Property Let boxFillColor(ByVal newValue As Long)
If mBoxFillColour = newValue Then Exit Property
undraw
mBoxFillColour = newValue
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

Public Property Get color() As Long
color = mColour
End Property

Public Property Let color(ByVal newValue As Long)
If mColour = newValue Then Exit Property
undraw
mColour = newValue
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
mLayer = value
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
Set mPosition = value.clone
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

Friend Property Let surface(ByVal value As PictureBox)
Set mSurface = value
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
Dim prevFont As StdFont
Dim textLines() As String
Dim i As Long
If mBlank Then Exit Sub
If Not mVisible Then Exit Sub
If Not mPosition.isAssigned Then Exit Sub
mDrawn = True

If Not mFont Is Nothing Then
        Set prevFont = New StdFont
        With mSurface.font
        prevFont.name = .name
        prevFont.Bold = .Bold
        prevFont.Italic = .Italic
        prevFont.Size = .Size
        prevFont.Charset = .Charset
        prevFont.Strikethrough = .Strikethrough
        prevFont.Underline = .Underline
        prevFont.Weight = .Weight
            
        If mFont.Size < 8 Then .Size = mFont.Size
       .name = mFont.name
       .Bold = mFont.Bold
       .Italic = mFont.Italic
       .Size = mFont.Size
       .Charset = mFont.Charset
       .Strikethrough = mFont.Strikethrough
       .Underline = mFont.Underline
       .Weight = mFont.Weight
    End With
End If
    
With mSurface
    .DrawMode = vbCopyPen
    .FontTransparent = True
    If mBox Then
        .DrawWidth = mBoxThickness
        .DrawStyle = mBoxStyle
        .ForeColor = mBoxColour
        .FillStyle = mBoxFillStyle
        .FillColor = mBoxFillColour
        mSurface.Line (mBoxRect.left, mBoxRect.bottom)-(mBoxRect.right, mBoxRect.top), , B
    End If
    .CurrentX = mTextRect.left
    .CurrentY = mTextRect.top
    .ForeColor = mColour
    textLines = Split(mText, vbCrLf)
    For i = 0 To UBound(textLines)
        mSurface.Print textLines(i)
        .CurrentX = mTextRect.left
    Next
End With

If Not prevFont Is Nothing Then
    With mSurface.font
        If prevFont.Size < 8 Then .Size = prevFont.Size
       .name = prevFont.name
       .Bold = prevFont.Bold
       .Italic = prevFont.Italic
       .Size = prevFont.Size
       .Charset = prevFont.Charset
       .Strikethrough = prevFont.Strikethrough
       .Underline = prevFont.Underline
       .Weight = prevFont.Weight
    End With
End If
    
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcRawBoundingRect()
Dim prevFont As StdFont

' NB: calculates the raw dimensions of the bounding box in twips, because
' the scale may not have been finalised when this is called, and is anyway
' subject to change
If mText = "" Then Exit Sub

If Not mFont Is Nothing Then
    Set prevFont = New StdFont
    With mSurface.font
        prevFont.name = .name
        prevFont.Bold = .Bold
        prevFont.Italic = .Italic
        prevFont.Size = .Size
        prevFont.Charset = .Charset
        prevFont.Strikethrough = .Strikethrough
        prevFont.Underline = .Underline
        prevFont.Weight = .Weight
        
        If mFont.Size < 8 Then .Size = mFont.Size
       .name = mFont.name
       .Bold = mFont.Bold
       .Italic = mFont.Italic
       .Size = mFont.Size
       .Charset = mFont.Charset
       .Strikethrough = mFont.Strikethrough
       .Underline = mFont.Underline
       .Weight = mFont.Weight
    End With
End If
    
With mRawBoundingRect
    .left = 0
    .right = mSurface.TextWidth(mText)
    .bottom = 0
    .top = Abs(mSurface.TextHeight(mText))
    .isvalid = True
End With

If Not prevFont Is Nothing Then
    With mSurface.font
        If prevFont.Size < 8 Then .Size = prevFont.Size
       .name = prevFont.name
       .Bold = prevFont.Bold
       .Italic = prevFont.Italic
       .Size = prevFont.Size
       .Charset = prevFont.Charset
       .Strikethrough = prevFont.Strikethrough
       .Underline = prevFont.Underline
       .Weight = prevFont.Weight
    End With
End If

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

'If mText = "Started here" Then Stop

If Not mPosition.isAssigned Then
    mBoundingRect.isvalid = False
    Exit Sub
End If

With mTextRect
    .left = mPosition.XAbsolute
    .right = .left + mRawBoundingRect.right
    .top = mPosition.YAbsolute
    .bottom = .top - mRawBoundingRect.top
    
    ' now add the offset if any
    
    If Not mOffset Is Nothing Then
        xOffset = mOffset.XAbsolute
        yOffset = mOffset.YAbsolute
        .left = .left + xOffset
        .right = .right + xOffset
        .bottom = .bottom + yOffset
        .top = .top + yOffset
    End If
    
    ' now adjust the rectangle to take account of alignment
    
    Select Case mAlign
    Case AlignTopLeft
        xshift = 0
        yshift = 0
    Case AlignCentreLeft
        xshift = 0
        yshift = (.top - .bottom) / 2
    Case AlignBottomLeft
        xshift = 0
        yshift = .top - .bottom
    Case AlignTopCentre
        xshift = (.left - .right) / 2
        yshift = 0
    Case AlignCentreCentre
        xshift = (.left - .right) / 2
        yshift = (.top - .bottom) / 2
    Case AlignBottomCentre
        xshift = (.left - .right) / 2
        yshift = .top - .bottom
    Case AlignTopRight
        xshift = .left - .right
        yshift = 0
    Case AlignCentreRight
        xshift = .left - .right
        yshift = (.top - .bottom) / 2
    Case AlignBottomRight
        xshift = .left - .right
        yshift = .top - .bottom
    End Select
    .left = .left + xshift
    .right = .right + xshift
    .bottom = .bottom + yshift
    .top = .top + yshift

    If mKeepInView Then
        ' now adjust the rectangle to take account of keepInView
    End If
End With

If mBox Then
    ' make the inside of the box 1mm wider and taller than the text, plus the thickness of the lines
    
    XIncrement = (1440 * mPaddingX / 25.4) / mSurface.Width * mSurface.ScaleWidth
    YIncrement = -(1440 * (mPaddingY - 0.5) / 25.4) / mSurface.Height * mSurface.ScaleHeight
    
    If mBoxThickness > 1 Then
        XLineThickness = ((mBoxThickness - 1) * Screen.TwipsPerPixelX) / (mSurface.Width / mSurface.ScaleWidth)
        YLineThickness = -((mBoxThickness - 1) * Screen.TwipsPerPixelY) / (mSurface.Height / mSurface.ScaleHeight)
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
    
    If mBoxThickness > 1 Then
        If mBoxStyle = LineInsideSolid Then
            mBoundingRect = mBoxRect
        Else
            With mBoundingRect
                .left = mBoxRect.left - XLineThickness / 2
                .right = mBoxRect.right + XLineThickness / 2
                .bottom = mBoxRect.bottom - YLineThickness / 2
                .top = mBoxRect.top + YLineThickness / 2
            End With
        End If
    Else
        mBoundingRect = mBoxRect
    End If
Else
    mBoxRect = mTextRect
    mBoundingRect = mTextRect
End If
mTextRect.isvalid = True
mBoxRect.isvalid = True
mBoundingRect.isvalid = True

End Sub

Private Sub convertPosition()
convertPositionX
convertPositionY
End Sub

Private Sub convertPositionX()
If mPosition Is Nothing Then Exit Sub
If mFixedX Then
    If mPosition.PositionTypeX = PositionAbsolute Then
        mPosition.PositionTypeX = PositionRelative
    End If
Else
    If mPosition.PositionTypeX <> PositionAbsolute Then
        mPosition.PositionTypeX = PositionAbsolute
    End If
End If
End Sub

Private Sub convertPositionY()
If mPosition Is Nothing Then Exit Sub
If mFixedY Then
    If mPosition.PositionTypeY = PositionAbsolute Then
        mPosition.PositionTypeY = PositionRelative
    End If
Else
    If mPosition.PositionTypeY <> PositionAbsolute Then
        mPosition.PositionTypeY = PositionAbsolute
    End If
End If
End Sub

Private Sub undraw()

If mBlank Then Exit Sub
If Not mVisible Then Exit Sub
If Not mPosition.isAssigned Then Exit Sub
mDrawn = True

With mSurface
    .DrawMode = vbCopyPen
    .DrawWidth = 1
    .DrawStyle = LineStyles.LineInsideSolid
    .ForeColor = .backColor
    .FillStyle = FillStyles.FillSolid
    .FillColor = .backColor
    mSurface.Line (mBoundingRect.left, mBoundingRect.bottom)-(mBoundingRect.right, mBoundingRect.top), , B
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
If mIncludeInAutoscale And mBoundingRect.isvalid Then
    If mBoundingRect.top >= mSurface.ScaleTop Or mBoundingRect.bottom <= (mSurface.ScaleTop + mSurface.ScaleHeight) Then
        mEventProxy.boundsExceeded
    End If
End If
End Sub
