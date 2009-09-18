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

'Event Click()
'
'Event DblCLick()
'
'Event SelectionStateChanged()

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
Private mSeriesUpdateNumber As Long

Private mViewport As Viewport
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

Private mPropertyOverrideFlags As TextPropertyOverrideFlags

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
gLogger.Log LogLevelHighDetail, "Text terminated"
Debug.Print "Text terminated"
End Sub

'================================================================================
' IGraphicObject Event Handlers
'================================================================================

Private Property Get IGraphicObject_boundingRectangle() As TRectangle
Static prevGaugeX As Double
Static prevGaugeY As Double
Static prevRect As TRectangle
Dim viewportChanges As Long

Dim failpoint As Long
On Error GoTo Err

viewportChanges = mViewport.CompareTo(prevGaugeX, prevGaugeY, prevRect)
If (viewportChanges And ViewportComparisonCodes.GaugeChangedX) Or _
    (viewportChanges And ViewportComparisonCodes.GaugeChangedY) _
Then
    calcRawBoundingRect
End If

If (viewportChanges And ViewportComparisonCodes.GaugeChangedX) Or _
    (viewportChanges And ViewportComparisonCodes.GaugeChangedY) Or _
    (mFixedX And (viewportChanges And ViewportComparisonCodes.BoundsChangedX)) Or _
    (mFixedY And (viewportChanges And ViewportComparisonCodes.BoundsChangedY)) Or _
    mSeriesUpdateNumber <> mSeries.SeriesUpdateNumber _
Then
    mSeriesUpdateNumber = mSeries.SeriesUpdateNumber
    calcCurrentBoundingRect
End If

prevGaugeX = mViewport.GaugeX
prevGaugeY = mViewport.GaugeY
prevRect = mViewport.Boundary

IGraphicObject_boundingRectangle = mBoundingRect

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "IGraphicObject_boundingRectangle" & "." & failpoint
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
'RaiseEvent Click
End Sub

Private Sub IGraphicObject_DblCLick()
'RaiseEvent DblCLick
End Sub

Private Sub IGraphicObject_draw( _
                ByRef areas() As TRectangle)
' note that
mCanvas.SetClippingRegion mViewport.ConvertTRectangleToRect(areas(0))
Draw
mCanvas.ClearClippingRegion
End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = Extended Or FixedX
End Property

Private Sub IGraphicObject_finish()
Set mCanvas = Nothing
Set mViewport = Nothing
End Sub

Private Property Get IGraphicObject_gaugeDependent() As Boolean
IGraphicObject_gaugeDependent = True
End Property

Private Property Let IGraphicObject_Handle(ByVal value As Long)
mHandle = value
End Property

Private Property Get IGraphicObject_Handle() As Long
IGraphicObject_Handle = mHandle
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
    'RaiseEvent SelectionStateChanged
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
'Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbcrlf, "") & ProjectName & "." & ModuleName & ":" & "IGraphicObject_undraw" & "." & failpoint
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
Case "ALIGN"
    initiateRedraw True, True
Case "BOX"
    initiateRedraw True, True
Case "BOXCOLOR"
    initiateRedraw False, False
Case "BOXFILLCOLOR"
    initiateRedraw False, False
Case "BOXFILLSTYLE"
    initiateRedraw True, False
Case "BOXSTYLE"
    initiateRedraw True, True
Case "BOXTHICKNESS"
    initiateRedraw True, True
Case "COLOR"
    initiateRedraw False, False
Case "FONT"
    initiateRedraw True, True
Case "PADDINGX"
    initiateRedraw True, True
Case "PADDINGY"
    initiateRedraw True, True
Case "EXTENDED"
    initiateRedraw True, True
Case "FIXEDX"
    initiateRedraw True, True
Case "FIXEDY"
    initiateRedraw True, True
Case "INCLUDEINAUTOSCALE"
    initiateRedraw False, False
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
If isPropertyOverrideFlagSet(TextIsSetAlign) Then Align = mAlign
End Property

Public Property Let Align(ByVal value As TextAlignModes)
If mAlign = value And isPropertyOverrideFlagSet(TextIsSetAlign) Then Exit Property
mAlign = value
setPropertyOverrideFlag TextIsSetAlign
initiateRedraw True, True
End Property

Public Property Get Box() As Boolean
Box = mSeries.Box
If Not mLocalStyle Is Nothing Then Box = mLocalStyle.Box
If isPropertyOverrideFlagSet(TextIsSetBox) Then Box = mBox
End Property

Public Property Let Box(ByVal value As Boolean)
If mBox = value And isPropertyOverrideFlagSet(TextIsSetBox) Then Exit Property
mBox = value
setPropertyOverrideFlag TextIsSetBox
initiateRedraw True, True
End Property

Public Property Get BoxColor() As Long
BoxColor = mSeries.BoxColor
If Not mLocalStyle Is Nothing Then BoxColor = mLocalStyle.BoxColor
If isPropertyOverrideFlagSet(TextIsSetBoxColor) Then BoxColor = mBoxColor
End Property

Public Property Let BoxColor(ByVal value As Long)
If mBoxColor = value And isPropertyOverrideFlagSet(TextIsSetBoxColor) Then Exit Property
mBoxColor = value
setPropertyOverrideFlag TextIsSetBoxColor
initiateRedraw False, False
End Property

Public Property Get BoxFillColor() As Long
BoxFillColor = mSeries.BoxFillColor
If Not mLocalStyle Is Nothing Then BoxFillColor = mLocalStyle.BoxFillColor
If isPropertyOverrideFlagSet(TextIsSetBoxFillColor) Then BoxFillColor = mBoxFillColor
End Property

Public Property Let BoxFillColor(ByVal value As Long)
If mBoxFillColor = value And isPropertyOverrideFlagSet(TextIsSetBoxFillColor) Then Exit Property
mBoxFillColor = value
setPropertyOverrideFlag TextIsSetBoxFillColor
initiateRedraw False, False
End Property

Public Property Get BoxFillStyle() As FillStyles
BoxFillStyle = mSeries.BoxFillStyle
If Not mLocalStyle Is Nothing Then BoxFillStyle = mLocalStyle.BoxFillStyle
If isPropertyOverrideFlagSet(TextIsSetBoxFillStyle) Then BoxFillStyle = mBoxFillStyle
End Property

Public Property Let BoxFillStyle(ByVal value As FillStyles)
If mBoxFillStyle = value And isPropertyOverrideFlagSet(TextIsSetBoxFillStyle) Then Exit Property
mBoxFillStyle = value
setPropertyOverrideFlag TextIsSetBoxFillStyle
initiateRedraw True, False
End Property

Public Property Get BoxFillWithBackgroundColor() As Boolean
BoxFillWithBackgroundColor = mSeries.BoxFillWithBackgroundColor
If Not mLocalStyle Is Nothing Then BoxFillWithBackgroundColor = mLocalStyle.BoxFillWithBackgroundColor
If isPropertyOverrideFlagSet(TextIsSetBoxFillWithBackGroundColor) Then BoxFillWithBackgroundColor = mBoxFillWithBackgroundColor
End Property

Public Property Let BoxFillWithBackgroundColor(ByVal value As Boolean)
If mBoxFillWithBackgroundColor = value And isPropertyOverrideFlagSet(TextIsSetBoxFillWithBackGroundColor) Then Exit Property
mBoxFillWithBackgroundColor = value
setPropertyOverrideFlag TextIsSetBoxFillWithBackGroundColor
initiateRedraw True, True
End Property

Public Property Get BoxStyle() As LineStyles
BoxStyle = mSeries.BoxStyle
If Not mLocalStyle Is Nothing Then BoxStyle = mLocalStyle.BoxStyle
If isPropertyOverrideFlagSet(TextIsSetBoxStyle) Then BoxStyle = mBoxStyle
End Property

Public Property Let BoxStyle(ByVal value As LineStyles)
If mBoxStyle = value And isPropertyOverrideFlagSet(TextIsSetBoxStyle) Then Exit Property
mBoxStyle = value
setPropertyOverrideFlag TextIsSetBoxStyle
initiateRedraw True, True
End Property

Public Property Get BoxThickness() As Long
BoxThickness = mSeries.BoxThickness
If Not mLocalStyle Is Nothing Then BoxThickness = mLocalStyle.BoxThickness
If isPropertyOverrideFlagSet(TextIsSetBoxThickness) Then BoxThickness = mBoxThickness
End Property

Public Property Let BoxThickness(ByVal value As Long)
If mBoxThickness = value And isPropertyOverrideFlagSet(TextIsSetBoxThickness) Then Exit Property
mBoxThickness = value
setPropertyOverrideFlag TextIsSetBoxThickness
initiateRedraw True, True
End Property

Public Property Get Color() As Long
Color = mSeries.Color
If Not mLocalStyle Is Nothing Then Color = mLocalStyle.Color
If isPropertyOverrideFlagSet(TextIsSetColor) Then Color = mColor
End Property

Public Property Let Color(ByVal value As Long)
If mColor = value And isPropertyOverrideFlagSet(TextIsSetColor) Then Exit Property
mColor = value
setPropertyOverrideFlag TextIsSetColor
initiateRedraw False, False
End Property

Public Property Get Extended() As Boolean
Extended = mSeries.Extended
If Not mLocalStyle Is Nothing Then Extended = mLocalStyle.Extended
If isPropertyOverrideFlagSet(TextIsSetExtended) Then Extended = mExtended
End Property

Public Property Let Extended(ByVal value As Boolean)
If mExtended = value And isPropertyOverrideFlagSet(TextIsSetExtended) Then Exit Property
mExtended = value
setPropertyOverrideFlag TextIsSetExtended
initiateRedraw True, True
End Property

Public Property Get FixedX() As Boolean
FixedX = mSeries.FixedX
If Not mLocalStyle Is Nothing Then FixedX = mLocalStyle.FixedX
If isPropertyOverrideFlagSet(TextIsSetFixedX) Then FixedX = mFixedX
End Property

Public Property Let FixedX(ByVal value As Boolean)
If mFixedX = value And isPropertyOverrideFlagSet(TextIsSetFixedX) Then Exit Property
mFixedX = value
setPropertyOverrideFlag TextIsSetFixedX
initiateRedraw True, True
End Property

Public Property Get FixedY() As Boolean
FixedY = mSeries.FixedY
If Not mLocalStyle Is Nothing Then FixedY = mLocalStyle.FixedY
If isPropertyOverrideFlagSet(TextIsSetFixedY) Then FixedY = mFixedY
End Property

Public Property Let FixedY(ByVal value As Boolean)
If mFixedY = value And isPropertyOverrideFlagSet(TextIsSetFixedY) Then Exit Property
mFixedY = value
setPropertyOverrideFlag TextIsSetFixedY
initiateRedraw True, True
End Property

Public Property Get Font() As StdFont
If isPropertyOverrideFlagSet(TextIsSetFont) Then
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
    clearPropertyOverrideFlag TextIsSetFont
Else
    setPropertyOverrideFlag TextIsSetFont
End If
initiateRedraw True, True
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mSeries.IncludeInAutoscale
If isPropertyOverrideFlagSet(TextIsSetIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value And isPropertyOverrideFlagSet(TextIsSetIncludeInAutoscale) Then Exit Property
mIncludeInAutoscale = value
setPropertyOverrideFlag TextIsSetIncludeInAutoscale
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
If mLayer = value And isPropertyOverrideFlagSet(TextIsSetLayer) Then Exit Property
mLayer = value
setPropertyOverrideFlag TextIsSetLayer
initiateRedraw False, False
End Property

Public Property Get Layer() As LayerNumbers
Layer = mSeries.Layer
If isPropertyOverrideFlagSet(TextIsSetLayer) Then Layer = mLayer
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
If mPaddingX = value And isPropertyOverrideFlagSet(TextIsSetPaddingX) Then Exit Property
mPaddingX = value
setPropertyOverrideFlag TextIsSetPaddingX
initiateRedraw True, True
End Property

Public Property Get PaddingX() As Double
PaddingX = mSeries.PaddingX
If Not mLocalStyle Is Nothing Then PaddingX = mLocalStyle.PaddingX
If isPropertyOverrideFlagSet(TextIsSetPaddingX) Then PaddingX = mPaddingX
End Property

Public Property Let PaddingY(ByVal value As Double)
If mPaddingY = value And isPropertyOverrideFlagSet(TextIsSetPaddingY) Then Exit Property
mPaddingY = value
setPropertyOverrideFlag TextIsSetPaddingY
initiateRedraw True, True
End Property

Public Property Get PaddingY() As Double
PaddingY = mSeries.PaddingY
If Not mLocalStyle Is Nothing Then PaddingY = mLocalStyle.PaddingY
If isPropertyOverrideFlagSet(TextIsSetPaddingY) Then PaddingY = mPaddingY
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
    If value.Equals(mPosition) Then
        If mPosition.CoordinateSystemX = CoordsLogical And mPosition.CoordinateSystemY = CoordsLogical Then Exit Property
    End If
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

Public Sub ClearOverrides()
mPropertyOverrideFlags = 0
initiateRedraw True, True
End Sub

Friend Sub Draw()
Dim failpoint As Long
On Error GoTo Err

If Not canBeDrawn Then Exit Sub
If Box Then drawBox
DrawText

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "draw" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


    
End Sub

Friend Sub Initialise( _
                ByVal pSeries As TextSeries, _
                ByVal pKey As String, _
                ByVal pGOSP As GraphObjServiceProvider, _
                ByVal pviewport As Viewport, _
                ByVal pLocalStyle As TextStyle)
                
Dim failpoint As Long
On Error GoTo Err

Set mSeries = pSeries
mKey = pKey
Set mGOSP = pGOSP
Set mViewport = pviewport
Set mCanvas = mViewport.Canvas
Set mLocalStyle = pLocalStyle

calcCurrentBoundingRect

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Initialise" & "." & failpoint
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
Dim alignmentOffset As TPoint
Dim lTextRect As RECT
Dim lBoxRect As RECT
Dim lBoundingRect As RECT

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
    .isValid = True
End With
    
If Not mOffset Is Nothing Then
    RectOffsetDim mTextRect, mOffset
End If
    
lTextRect = mViewport.ConvertTRectangleToRect(mTextRect)

If Box Then
    
    lBoxRect = lTextRect
    
    TWWin32API.InflateRect lBoxRect, _
                            mViewport.ConvertDistanceToPixelsX(PaddingX / 10), _
                            mViewport.ConvertDistanceToPixelsY(PaddingY / 10)
    
    Dim lBoxThickness As Long
    lBoxThickness = BoxThickness
    If (lBoxThickness > 1 And BoxStyle = LineInsideSolid) Or lBoxThickness = 1 Then
        TWWin32API.InflateRect lBoxRect, lBoxThickness, lBoxThickness
        lBoundingRect = lBoxRect
        TWWin32API.InflateRect lBoundingRect, 1, 1
    Else
        TWWin32API.InflateRect lBoxRect, _
                                lBoxThickness - Int(lBoxThickness / 2), _
                                lBoxThickness - Int(lBoxThickness / 2)
        lBoundingRect = lBoxRect
        TWWin32API.InflateRect lBoundingRect, Int(lBoxThickness / 2) + 1, Int(lBoxThickness / 2) + 1
    End If
Else
    lBoxRect = lTextRect
    lBoundingRect = lTextRect
End If

mTextRect = mViewport.ConvertRectToTRectangle(lTextRect)
mBoxRect = mViewport.ConvertRectToTRectangle(lBoxRect)
alignmentOffset = getAlignmentOffset(mTextRect, mBoxRect)

Dim alignmentOffsetPixelsX As Long
Dim alignmentOffsetPixelsY As Long
alignmentOffsetPixelsX = mViewport.ConvertLogicalToPixelsX(alignmentOffset.X)
alignmentOffsetPixelsY = mViewport.ConvertLogicalToPixelsY(alignmentOffset.Y)

TWWin32API.OffsetRect lTextRect, alignmentOffsetPixelsX, -alignmentOffsetPixelsY
TWWin32API.OffsetRect lBoxRect, alignmentOffsetPixelsX, -alignmentOffsetPixelsY
TWWin32API.OffsetRect lBoundingRect, alignmentOffsetPixelsX, -alignmentOffsetPixelsY

mTextRect = mViewport.ConvertRectToTRectangle(lTextRect)
mBoxRect = mViewport.ConvertRectToTRectangle(lBoxRect)
mBoundingRect = mViewport.ConvertRectToTRectangle(lBoundingRect)

End Sub

Private Function canBeDrawn() As Boolean
If mBlank Then Exit Function
If Not mVisible Then Exit Function
If Not mPosition.IsAssigned Then Exit Function
canBeDrawn = True
End Function

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As TextPropertyOverrideFlags)
mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)
End Sub

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
With mViewport
    .Canvas.SetPenAttributes BoxColor, _
                    BoxThickness, _
                    BoxStyle, _
                    DrawModes.DrawModeCopyPen
    .Canvas.SetBrushAttributes BoxFillColor, pFillStyle
    .Canvas.DrawRectangle .NewPoint(mBoxRect.Left, mBoxRect.Bottom), .NewPoint(mBoxRect.Right, mBoxRect.Top)
End With
End Sub

Private Sub drawBoxWithBackgroundFillColor()
mViewport.PaintBackground
drawBoxWithSpecifiedFillStyle FillTransparent
End Sub

Private Sub DrawText()
With mViewport
    .Canvas.SetTextAttributes Color, Font, True
    .Canvas.DrawText mText, .NewPoint(mTextRect.Left, mTextRect.Top)
End With
End Sub

Private Function getAlignmentOffset( _
                ByRef pTextBoundary As TRectangle, _
                ByRef pBoxBoundary As TRectangle) As TPoint
Dim textTopLeft As TPoint
Dim alignPoint As TPoint

textTopLeft = RectTopLeft(pTextBoundary)
Select Case Align
Case AlignTopLeft
    alignPoint = textTopLeft
Case AlignCentreLeft
    alignPoint = RectCentreLeft(pTextBoundary)
Case AlignBottomLeft
    alignPoint = RectBottomLeft(pTextBoundary)
Case AlignTopCentre
    alignPoint = RectTopCentre(pTextBoundary)
Case AlignCentreCentre
    alignPoint = RectCentreCentre(pTextBoundary)
Case AlignBottomCentre
    alignPoint = RectBottomCentre(pTextBoundary)
Case AlignTopRight
    alignPoint = RectTopRight(pTextBoundary)
Case AlignCentreRight
    alignPoint = RectCentreRight(pTextBoundary)
Case AlignBottomRight
    alignPoint = RectBottomRight(pTextBoundary)
Case AlignBoxTopLeft
    alignPoint = RectTopLeft(pBoxBoundary)
Case AlignBoxCentreLeft
    alignPoint = RectCentreLeft(pBoxBoundary)
Case AlignBoxBottomLeft
    alignPoint = RectBottomLeft(pBoxBoundary)
Case AlignBoxTopCentre
    alignPoint = RectTopCentre(pBoxBoundary)
Case AlignBoxCentreCentre
    alignPoint = RectCentreCentre(pBoxBoundary)
Case AlignBoxBottomCentre
    alignPoint = RectBottomCentre(pBoxBoundary)
Case AlignBoxTopRight
    alignPoint = RectTopRight(pBoxBoundary)
Case AlignBoxCentreRight
    alignPoint = RectCentreRight(pBoxBoundary)
Case AlignBoxBottomRight
    alignPoint = RectBottomRight(pBoxBoundary)
End Select

getAlignmentOffset = PointSubtract(textTopLeft, alignPoint)
End Function

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

Private Function isPropertyOverrideFlagSet( _
                ByVal flag As TextPropertyOverrideFlags) As Boolean
isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)
End Function

Private Sub setPropertyOverrideFlag( _
                ByVal flag As TextPropertyOverrideFlags)
mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)
End Sub

Private Sub undraw()
If canBeDrawn Then mGOSP.ObjectUndrawn mHandle, mBoundingRect
End Sub

