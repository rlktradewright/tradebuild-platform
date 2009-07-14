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


Private Const ModuleName                As String = "Text"

'================================================================================
' Member variables and constants
'================================================================================

Private mSeries As TextSeries

Private mCanvas As Canvas
Private mLayer As Long
Private mHandle As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private WithEvents mLocalStyle As TextStyle
Attribute mLocalStyle.VB_VarHelpID = -1

Private mInScope As Boolean
Private mVisible As Boolean

Private mGOSP As GraphObjServiceProvider

Private mNoDraw As Boolean

Private mText As String

Private mRawBoundingRect As TRectangle
Private mBoundingRect As TRectangle ' the rectangle that surrounds the box (if
                                ' any - including the thickness of the lines) or
                                ' the Text if no box
Private mTextRect As TRectangle ' the rectangle containing the Text
Private mBoxRect As TRectangle  ' the rectangle used to draw the box

Private mIsSetFont As Boolean
Private mIsSetColor As Boolean
Private mIsSetBox As Boolean
Private mIsSetBoxColor As Boolean
Private mIsSetBoxStyle As Boolean
Private mIsSetBoxThickness As Boolean
Private mIsSetBoxFillColor As Boolean
Private mIsSetBoxFillStyle As Boolean
Private mIsSetBoxFillWithBackgroundColor As Boolean
Private mIsSetAlign As Boolean
Private mIsSetPaddingX As Boolean
Private mIsSetPaddingY As Boolean
Private mIsSetFixedX As Boolean
Private mIsSetFixedY As Boolean
Private mIsSetIncludeInAutoscale As Boolean
Private mIsSetExtended As Boolean
Private mIsSetLayer As Boolean

Private mFont As StdFont
Private mColor As Long
Private mBox As Boolean
Private mBoxColor As Long
Private mBoxStyle As LineStyles
Private mBoxThickness As Long
Private mBoxFillColor As Long
Private mBoxFillStyle As FillStyles
Private mBoxFillWithBackgroundColor As Boolean
Private mAlign As TextAlignModes
Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean
Private mPaddingX As Double
Private mPaddingY As Double

Private mBlank As Boolean

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

Private Sub Class_Terminate()
Debug.Print "Text terminated"
End Sub

'================================================================================
' IGraphicObject Event Handlers
'================================================================================

Private Property Get IGraphicObject_boundingRectangle() As TRectangle
Static prevGaugeX As Double
Static prevGaugeY As Double
Static prevRect As TRectangle
Dim canvasChanges As Long

Dim failpoint As Long
On Error GoTo Err

canvasChanges = mCanvas.CompareTo(prevGaugeX, prevGaugeY, prevRect)
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

prevGaugeX = mCanvas.GaugeX
prevGaugeY = mCanvas.GaugeY
prevRect = mCanvas.Boundary

IGraphicObject_boundingRectangle = mBoundingRect

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_boundingRectangle" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Property

Private Property Get IGraphicObject_BoundingRectanglesAt(ByVal PeriodNumber As Long) As TRectangle()
'
End Property

Private Property Get IGraphicObject_capabilities() As GraphicObjectCapabilities
IGraphicObject_capabilities = GraphicObjectCapabilities.BlockUndraw
End Property

Private Sub IGraphicObject_Click()
RaiseEvent Click
End Sub

Private Sub IGraphicObject_DblCLick()
RaiseEvent DblCLick
End Sub

Private Sub IGraphicObject_draw( _
                ByRef areas() As TRectangle)
Draw
End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = Extended Or FixedX
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

Private Function IGraphicObject_HitTest(ByVal X As Double, ByVal Y As Double) As Boolean
IGraphicObject_HitTest = RectContainsPoint(mBoundingRect, X, Y)
End Function

Private Property Get IGraphicObject_IncludeInAutoscale() As Boolean
IGraphicObject_IncludeInAutoscale = IncludeInAutoscale
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
IGraphicObject_noDraw = mBlank Or mNoDraw
End Property

Private Property Get IGraphicObject_PeriodNumber() As Long
If Extended Or Not mBoundingRect.isValid Then
    IGraphicObject_PeriodNumber = -1
Else
    Select Case Align
    Case AlignTopLeft, AlignCentreLeft, AlignBottomLeft, _
            AlignBoxTopLeft, AlignBoxCentreLeft, AlignBoxBottomLeft
        IGraphicObject_PeriodNumber = Int(mBoundingRect.Left + 0.5)
    Case AlignTopCentre, AlignCentreCentre, AlignBottomCentre, _
            AlignBoxTopCentre, AlignBoxCentreCentre, AlignBoxBottomCentre
        IGraphicObject_PeriodNumber = Int((mBoundingRect.Left + mBoundingRect.Right) / 2 + 0.5)
    Case AlignTopRight, AlignCentreRight, AlignBottomRight, _
            AlignBoxTopRight, AlignBoxCentreRight, AlignBoxBottomRight
        IGraphicObject_PeriodNumber = Int(mBoundingRect.Right + 0.5)
    End Select
End If
End Property

Private Property Get IGraphicObject_scaleDependent() As Boolean
IGraphicObject_scaleDependent = FixedX Or FixedY
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
IGraphicObject_SeriesID = mSeries.Id
End Property

Private Property Get IGraphicObject_Timestamp() As Date
IGraphicObject_Timestamp = 0
End Property

'Private Sub IGraphicObject_undraw( _
'                ByRef area As TRectangle, _
'                ByVal Hdc As Long)
'Dim failpoint As Long
'On Error GoTo Err
'
'undraw
'
'Exit Sub
'
'Err:
'Dim errNumber As Long: errNumber = Err.Number
'Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_undraw" & "." & failpoint & IIf(err.Source <> "", vbCrLf & err.Source, "")
'Dim errDescription As String: errDescription = Err.Description
'gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
'Err.Raise errNumber, errSource, errDescription
'
'
'End Sub

Private Property Let IGraphicObject_visible(ByVal value As Boolean)
mVisible = value
End Property

'================================================================================
' mLocalStyle Event Handlers
'================================================================================

Private Sub mLocalStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
Dim failpoint As Long
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case UCase$("align")
    initiateRedraw True, True
Case UCase$("box")
    initiateRedraw True, True
Case UCase$("boxColor")
    initiateRedraw False, False
Case UCase$("boxFillColor")
    initiateRedraw False, False
Case UCase$("boxFillStyle")
    initiateRedraw True, False
Case UCase$("boxStyle")
    initiateRedraw True, True
Case UCase$("boxThickness")
    initiateRedraw True, True
Case UCase$("color")
    initiateRedraw False, False
Case UCase$("Font")
    initiateRedraw True, True
Case UCase$("paddingX")
    initiateRedraw True, True
Case UCase$("paddingY")
    initiateRedraw True, True
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & "mStyle_PropertyChanged", _
            "Unhandled property change"
End Select

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "mLocalStyle_PropertyChanged" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get Align() As TextAlignModes
Align = mSeries.Align
If Not mLocalStyle Is Nothing Then Align = mLocalStyle.Align
If mIsSetAlign Then Align = mAlign
End Property

Public Property Let Align(ByVal value As TextAlignModes)
If mAlign = value And mIsSetAlign Then Exit Property
mAlign = value
mIsSetAlign = True
initiateRedraw True, True
End Property

Public Property Get Box() As Boolean
Box = mSeries.Box
If Not mLocalStyle Is Nothing Then Box = mLocalStyle.Box
If mIsSetBox Then Box = mBox
End Property

Public Property Let Box(ByVal value As Boolean)
If mBox = value And mIsSetBox Then Exit Property
mBox = value
mIsSetBox = True
initiateRedraw True, True
End Property

Public Property Get BoxColor() As OLE_COLOR
BoxColor = mSeries.BoxColor
If Not mLocalStyle Is Nothing Then BoxColor = mLocalStyle.BoxColor
If mIsSetBoxColor Then BoxColor = mBoxColor
End Property

Public Property Let BoxColor(ByVal value As OLE_COLOR)
If mBoxColor = value And mIsSetBoxColor Then Exit Property
mBoxColor = value
mIsSetBoxColor = True
initiateRedraw False, False
End Property

Public Property Get BoxFillColor() As OLE_COLOR
BoxFillColor = mSeries.BoxFillColor
If Not mLocalStyle Is Nothing Then BoxFillColor = mLocalStyle.BoxFillColor
If mIsSetBoxFillColor Then BoxFillColor = mBoxFillColor
End Property

Public Property Let BoxFillColor(ByVal value As OLE_COLOR)
If mBoxFillColor = value And mIsSetBoxFillColor Then Exit Property
mBoxFillColor = value
mIsSetBoxFillColor = True
initiateRedraw False, False
End Property

Public Property Get BoxFillStyle() As FillStyles
BoxFillStyle = mSeries.BoxFillStyle
If Not mLocalStyle Is Nothing Then BoxFillStyle = mLocalStyle.BoxFillStyle
If mIsSetBoxFillStyle Then BoxFillStyle = mBoxFillStyle
End Property

Public Property Let BoxFillStyle(ByVal value As FillStyles)
If mBoxFillStyle = value And mIsSetBoxFillStyle Then Exit Property
mBoxFillStyle = value
mIsSetBoxFillStyle = True
initiateRedraw True, False
End Property

Public Property Get BoxFillWithBackgroundColor() As Boolean
BoxFillWithBackgroundColor = mSeries.BoxFillWithBackgroundColor
If Not mLocalStyle Is Nothing Then BoxFillWithBackgroundColor = mLocalStyle.BoxFillWithBackgroundColor
If mIsSetBoxFillWithBackgroundColor Then BoxFillWithBackgroundColor = mBoxFillWithBackgroundColor
End Property

Public Property Let BoxFillWithBackgroundColor(ByVal value As Boolean)
If mBoxFillWithBackgroundColor = value And mIsSetBoxFillWithBackgroundColor Then Exit Property
mBoxFillWithBackgroundColor = value
mIsSetBoxFillWithBackgroundColor = True
initiateRedraw True, True
End Property

Public Property Get BoxStyle() As LineStyles
BoxStyle = mSeries.BoxStyle
If Not mLocalStyle Is Nothing Then BoxStyle = mLocalStyle.BoxStyle
If mIsSetBoxStyle Then BoxStyle = mBoxStyle
End Property

Public Property Let BoxStyle(ByVal value As LineStyles)
If mBoxStyle = value And mIsSetBoxStyle Then Exit Property
mBoxStyle = value
mIsSetBoxStyle = True
initiateRedraw True, True
End Property

Public Property Get BoxThickness() As Long
BoxThickness = mSeries.BoxThickness
If Not mLocalStyle Is Nothing Then BoxThickness = mLocalStyle.BoxThickness
If mIsSetBoxThickness Then BoxThickness = mBoxThickness
End Property

Public Property Let BoxThickness(ByVal value As Long)
If mBoxThickness = value And mIsSetBoxThickness Then Exit Property
mBoxThickness = value
mIsSetBoxThickness = True
initiateRedraw True, True
End Property

Public Property Get Color() As OLE_COLOR
Color = mSeries.Color
If Not mLocalStyle Is Nothing Then Color = mLocalStyle.Color
If mIsSetColor Then Color = mColor
End Property

Public Property Let Color(ByVal value As OLE_COLOR)
If mColor = value And mIsSetColor Then Exit Property
mColor = value
mIsSetColor = True
initiateRedraw False, False
End Property

Public Property Get Extended() As Boolean
Extended = mSeries.Extended
If mIsSetExtended Then Extended = mExtended
End Property

Public Property Let Extended(ByVal value As Boolean)
If mExtended = value And mIsSetExtended Then Exit Property
mExtended = value
mIsSetExtended = True
initiateRedraw True, True
End Property

Public Property Get FixedX() As Boolean
FixedX = mSeries.FixedX
If mIsSetFixedX Then FixedX = mFixedX
End Property

Public Property Let FixedX(ByVal value As Boolean)
If mFixedX = value And mIsSetFixedX Then Exit Property
mFixedX = value
mIsSetFixedX = True
ConvertPositionX
initiateRedraw True, True
End Property

Public Property Get FixedY() As Boolean
FixedY = mSeries.FixedY
If mIsSetFixedY Then FixedY = mFixedY
End Property

Public Property Let FixedY(ByVal value As Boolean)
If mFixedY = value And mIsSetFixedY Then Exit Property
mFixedY = value
mIsSetFixedY = True
ConvertPositionY
initiateRedraw True, True
End Property

Public Property Get Font() As StdFont
If mIsSetFont Then
    Set Font = mFont
ElseIf Not mLocalStyle Is Nothing Then
    Set Font = mLocalStyle.Font
Else
    Set Font = mSeries.Font
End If
End Property

Public Property Let Font(ByVal value As StdFont)
Set mFont = value
If mFont Is Nothing Then
    mIsSetFont = False
Else
    mIsSetFont = True
End If
initiateRedraw True, True
End Property

Friend Property Get Handle() As Long
Handle = mHandle
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mSeries.IncludeInAutoscale
If mIsSetIncludeInAutoscale Then IncludeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value And mIsSetIncludeInAutoscale Then Exit Property
mIncludeInAutoscale = value
mIsSetIncludeInAutoscale = True
initiateRedraw False, False
End Property

Public Property Get InScope() As Boolean
InScope = mInScope
End Property

Public Property Let IsSelectable(ByVal value As Boolean)
mIsSelectable = value
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

Public Property Let Layer(ByVal value As LayerNumbers)
If mLayer = value And mIsSetLayer Then Exit Property
mLayer = value
mIsSetLayer = True
initiateRedraw False, False
End Property

Public Property Get Layer() As LayerNumbers
Layer = mSeries.Layer
If mIsSetLayer Then Layer = mLayer
End Property

Public Property Get LocalStyle() As TextStyle
Set LocalStyle = mLocalStyle
End Property

Public Property Let LocalStyle( _
                ByVal value As TextStyle)
Set mLocalStyle = value
initiateRedraw True, True
End Property

Public Property Let PaddingX(ByVal value As Double)
If mPaddingX = value And mIsSetPaddingX Then Exit Property
mPaddingX = value
mIsSetPaddingX = True
initiateRedraw True, True
End Property

Public Property Get PaddingX() As Double
PaddingX = mSeries.PaddingX
If Not mLocalStyle Is Nothing Then PaddingX = mLocalStyle.PaddingX
If mIsSetPaddingX Then PaddingX = mPaddingX
End Property

Public Property Let PaddingY(ByVal value As Double)
If mPaddingY = value And mIsSetPaddingY Then Exit Property
mPaddingY = value
mIsSetPaddingY = True
initiateRedraw True, True
End Property

Public Property Get PaddingY() As Double
PaddingY = mSeries.PaddingY
If Not mLocalStyle Is Nothing Then PaddingY = mLocalStyle.PaddingY
If mIsSetPaddingY Then PaddingY = mPaddingY
End Property

Public Property Get position() As Point
Set position = mPosition
End Property

Public Property Let position(ByVal value As Point)
If value Is Nothing Then
    Set mPosition = Nothing
    mBoundingRect.isValid = False
    mBlank = True
    initiateRedraw True, True
    Exit Property
ElseIf Not value.IsAssigned Then
    Set mPosition = Nothing
    mBoundingRect.isValid = False
    mBlank = True
    initiateRedraw True, True
    Exit Property
ElseIf Not mPosition Is Nothing Then
    If value.Equals(mPosition) Then Exit Property
End If

Set mPosition = value
If mText <> "" Then mBlank = False
If mPosition.CoordinateSystemX <> CoordsLogical Then FixedX = True
If mPosition.CoordinateSystemY <> CoordsLogical Then FixedY = True
ConvertPosition
initiateRedraw True, True
End Property

Public Property Get Offset() As Dimension
Set Offset = mOffset
End Property

Public Property Let Offset(ByVal value As Dimension)
If value Is Nothing Then
    Set mOffset = Nothing
ElseIf Not value.IsAssigned Then
    Set mOffset = Nothing
ElseIf Not mOffset Is Nothing Then
    If value.Equals(mOffset) Then Exit Property
Else
    Set mOffset = value
End If
initiateRedraw True, True
End Property

Public Property Let Text(value As String)
If value = mText Then Exit Property
mText = value
mBlank = (value = "" Or mPosition Is Nothing)
initiateRedraw True, True
End Property

Public Property Get Text() As String
Text = mText
End Property

Public Property Get Visible() As Boolean
Visible = mVisible
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Draw()
Dim failpoint As Long
On Error GoTo Err

If Not canBeDrawn Then Exit Sub
If Box Then drawBox
DrawText

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "draw" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


    
End Sub

Friend Sub Initialise( _
                ByVal series As TextSeries, _
                ByVal Key As String, _
                ByVal GOSP As GraphObjServiceProvider, _
                ByVal Canvas As Canvas, _
                ByVal LocalStyle As TextStyle)
                
Dim failpoint As Long
On Error GoTo Err

Set mSeries = series
mKey = Key
Set mGOSP = GOSP
Set mCanvas = Canvas
Set mLocalStyle = LocalStyle

calcCurrentBoundingRect

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Initialise" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
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

mCanvas.SetTextAttributes Color, Font, True
With mRawBoundingRect
    .Left = 0
    .Right = mCanvas.TextWidth(mText)
    .Bottom = 0
    .Top = mCanvas.TextHeight(mText)
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

If mPosition Is Nothing Then
    mBoundingRect.isValid = False
    Exit Sub
End If

If Not mPosition.IsAssigned Then
    mBoundingRect.isValid = False
    Exit Sub
End If

With mTextRect
    .Left = mPosition.XLogical
    .Right = .Left + mRawBoundingRect.Right
    .Top = mPosition.YLogical
    .Bottom = .Top - mRawBoundingRect.Top
    
    ' now Add the Offset if any
    
    If Not mOffset Is Nothing Then
        xOffset = mOffset.XLogical
        yOffset = mOffset.YLogical
        .Left = .Left + xOffset
        .Right = .Right + xOffset
        .Bottom = .Bottom + yOffset
        .Top = .Top + yOffset
    End If
End With

If Box Then
    ' make the inside of the box 1mm wider and taller than the Text, plus the thickness of the lines
    
    Dim padding As Dimension
    Set padding = mCanvas.NewDimension((mPaddingX + 0.5) / 10, (IIf(mPaddingY > 0.5, mPaddingY - 0.5, 0)) / 10)
    XIncrement = padding.XLogical
    YIncrement = padding.YLogical
    
    If BoxThickness > 1 Then
        XLineThickness = mCanvas.ConvertPixelsToLogicalX(mBoxThickness - 1)
        YLineThickness = mCanvas.ConvertPixelsToLogicalY(mBoxThickness - 1)
        If BoxStyle = LineInsideSolid Then
            XIncrement = XIncrement + XLineThickness
            YIncrement = YIncrement + YLineThickness
        Else
            XIncrement = XIncrement + XLineThickness / 2
            YIncrement = YIncrement + YLineThickness / 2
        End If
    End If
    
    With mBoxRect
        .Left = mTextRect.Left - XIncrement
        .Right = mTextRect.Right + XIncrement
        .Bottom = mTextRect.Bottom - YIncrement
        .Top = mTextRect.Top + YIncrement
    End With
    
    If BoxThickness > 1 And BoxStyle <> LineInsideSolid Then
        With mBoundingRect
            .Left = mBoxRect.Left - XLineThickness / 2
            .Right = mBoxRect.Right + XLineThickness / 2
            .Bottom = mBoxRect.Bottom - YLineThickness / 2
            .Top = mBoxRect.Top + YLineThickness / 2
        End With
    Else
        mBoundingRect = mBoxRect
    End If
    ' now adjust the rectangle to take account of alignment

    Select Case Align
    Case AlignTopLeft
        xshift = 0
        yshift = 0
    Case AlignCentreLeft
        xshift = 0
        yshift = (mTextRect.Top - mTextRect.Bottom) / 2
    Case AlignBottomLeft
        xshift = 0
        yshift = mTextRect.Top - mTextRect.Bottom
    Case AlignTopCentre
        xshift = (mTextRect.Left - mTextRect.Right) / 2
        yshift = 0
    Case AlignCentreCentre
        xshift = (mTextRect.Left - mTextRect.Right) / 2
        yshift = (mTextRect.Top - mTextRect.Bottom) / 2
    Case AlignBottomCentre
        xshift = (mTextRect.Left - mTextRect.Right) / 2
        yshift = mTextRect.Top - mTextRect.Bottom
    Case AlignTopRight
        xshift = mTextRect.Left - mTextRect.Right
        yshift = 0
    Case AlignCentreRight
        xshift = mTextRect.Left - mTextRect.Right
        yshift = (mTextRect.Top - mTextRect.Bottom) / 2
    Case AlignBottomRight
        xshift = mTextRect.Left - mTextRect.Right
        yshift = mTextRect.Top - mTextRect.Bottom
    Case AlignBoxTopLeft
        xshift = mTextRect.Left - mBoundingRect.Left
        yshift = mTextRect.Top - mBoundingRect.Top
    Case AlignBoxCentreLeft
        xshift = mTextRect.Left - mBoundingRect.Left
        yshift = (mTextRect.Top - mTextRect.Bottom) / 2
    Case AlignBoxBottomLeft
        xshift = mTextRect.Left - mBoundingRect.Left
        yshift = mTextRect.Top - mTextRect.Bottom + mTextRect.Top - mBoundingRect.Top
    Case AlignBoxTopCentre
        xshift = (mTextRect.Left - mTextRect.Right) / 2
        yshift = mTextRect.Top - mBoundingRect.Top
    Case AlignBoxCentreCentre
        xshift = (mTextRect.Left - mTextRect.Right) / 2
        yshift = (mTextRect.Top - mTextRect.Bottom) / 2
    Case AlignBoxBottomCentre
        xshift = (mTextRect.Left - mTextRect.Right) / 2
        yshift = mTextRect.Top - mTextRect.Bottom + mTextRect.Top - mBoundingRect.Top
    Case AlignBoxTopRight
        xshift = mTextRect.Left - mTextRect.Right - (mBoundingRect.Right - mTextRect.Right)
        yshift = mTextRect.Top - mBoundingRect.Top
    Case AlignBoxCentreRight
        xshift = mTextRect.Left - mTextRect.Right - (mBoundingRect.Right - mTextRect.Right)
        yshift = (mTextRect.Top - mTextRect.Bottom) / 2
    Case AlignBoxBottomRight
        xshift = mTextRect.Left - mTextRect.Right - (mBoundingRect.Right - mTextRect.Right)
        yshift = mTextRect.Top - mTextRect.Bottom + mTextRect.Top - mBoundingRect.Top
    End Select

    With mTextRect
        .Left = .Left + xshift
        .Right = .Right + xshift
        .Bottom = .Bottom + yshift
        .Top = .Top + yshift
    End With

    With mBoxRect
        .Left = .Left + xshift
        .Right = .Right + xshift
        .Bottom = .Bottom + yshift
        .Top = .Top + yshift
    End With

    With mBoundingRect
        .Left = .Left + xshift
        .Right = .Right + xshift
        .Bottom = .Bottom + yshift
        .Top = .Top + yshift
    End With

Else
    ' now adjust the rectangle to take account of alignment
    
    With mTextRect
        Select Case Align
        Case AlignTopLeft, AlignBoxTopLeft
            xshift = 0
            yshift = 0
        Case AlignCentreLeft, AlignBoxCentreLeft
            xshift = 0
            yshift = (.Top - .Bottom) / 2
        Case AlignBottomLeft, AlignBoxBottomLeft
            xshift = 0
            yshift = .Top - .Bottom
        Case AlignTopCentre, AlignBoxTopCentre
            xshift = (.Left - .Right) / 2
            yshift = 0
        Case AlignCentreCentre, AlignBoxCentreCentre
            xshift = (.Left - .Right) / 2
            yshift = (.Top - .Bottom) / 2
        Case AlignBottomCentre, AlignBoxBottomCentre
            xshift = (.Left - .Right) / 2
            yshift = .Top - .Bottom
        Case AlignTopRight, AlignBoxTopRight
            xshift = .Left - .Right
            yshift = 0
        Case AlignCentreRight, AlignBoxCentreRight
            xshift = .Left - .Right
            yshift = (.Top - .Bottom) / 2
        Case AlignBottomRight, AlignBoxBottomRight
            xshift = .Left - .Right
            yshift = .Top - .Bottom
        End Select
    
        .Left = .Left + xshift
        .Right = .Right + xshift
        .Bottom = .Bottom + yshift
        .Top = .Top + yshift
    End With
    
    mBoxRect = mTextRect
    mBoundingRect = mTextRect
End If

mTextRect.isValid = True
mBoxRect.isValid = True
mBoundingRect.isValid = True

End Sub

Private Function canBeDrawn() As Boolean
If mBlank Then Exit Function
If Not mVisible Then Exit Function
If Not mPosition.IsAssigned Then Exit Function
canBeDrawn = True
End Function

Private Sub ConvertPosition()
ConvertPositionX
ConvertPositionY
End Sub

Private Sub ConvertPositionX()
If mPosition Is Nothing Then Exit Sub
If FixedX Then
    If mPosition.CoordinateSystemX = CoordsLogical Then
        mPosition.CoordinateSystemX = CoordsDistance
    End If
Else
    If mPosition.CoordinateSystemX <> CoordsLogical Then
        mPosition.CoordinateSystemX = CoordsLogical
    End If
End If
End Sub

Private Sub ConvertPositionY()
If mPosition Is Nothing Then Exit Sub
If FixedY Then
    If mPosition.CoordinateSystemY = CoordsLogical Then
        mPosition.CoordinateSystemY = CoordsDistance
    End If
Else
    If mPosition.CoordinateSystemY <> CoordsLogical Then
        mPosition.CoordinateSystemY = CoordsLogical
    End If
End If
End Sub

Private Sub drawBox()
If BoxFillWithBackgroundColor Then
    drawBoxWithBackgroundFillColor
Else
    drawBoxWithSpecifiedFillStyle BoxFillStyle
End If
End Sub

Private Sub drawBoxWithSpecifiedFillStyle( _
                ByVal pFillStyle As FillStyles)
With mCanvas
    .SetPenAttributes BoxColor, _
                    BoxThickness, _
                    BoxStyle, _
                    DrawModes.DrawModeCopyPen
    .SetBrushAttributes BoxFillColor, pFillStyle
    .DrawRectangle .NewPoint(mBoxRect.Left, mBoxRect.Bottom), .NewPoint(mBoxRect.Right, mBoxRect.Top)
End With
End Sub

Private Sub drawBoxWithBackgroundFillColor()
With mCanvas
    .PaintBackground
    drawBoxWithSpecifiedFillStyle FillTransparent
End With
End Sub

Private Sub DrawText()
With mCanvas
    .SetTextAttributes Color, Font, True
    .DrawText mText, .NewPoint(mTextRect.Left, mTextRect.Top)
End With
End Sub

Private Sub initiateRedraw( _
                ByVal undrawCurrentImage As Boolean, _
                ByVal recalcBoundaries As Boolean)
If mHandle = 0 Then
    ' haven't yet been added to chart region
    Exit Sub
End If
If undrawCurrentImage Then undraw
If recalcBoundaries Then
    calcRawBoundingRect
    calcCurrentBoundingRect
End If
mGOSP.ObjectChanged mHandle, mBoundingRect
End Sub

Private Sub undraw()
If canBeDrawn Then mGOSP.ObjectUndrawn mHandle, mBoundingRect
End Sub

