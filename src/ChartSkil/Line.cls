VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Line"
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


Private Const ModuleName                    As String = "Line"

'================================================================================
' Member variables
'================================================================================

Private mSeries As LineSeries
Private mSeriesUpdateNumber As Long

Private mCanvas  As Canvas
Private mLayer As Long
Private mHandle As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private WithEvents mLocalStyle As LineStyle
Attribute mLocalStyle.VB_VarHelpID = -1

Private mInScope As Boolean
Private mVisible As Boolean

Private mGOSP As GraphObjServiceProvider

Private mBoundingRect As TRectangle
Private mStartArrowRect As TRectangle
Private mEndArrowRect As TRectangle
Private mRegionRect As TRectangle

Private mEndArrowX1Offset As Double   ' horiz Offset from apex of upper Arrow endpoint
Private mEndArrowY1Offset As Double   ' vert Offset from apex of upper Arrow endpoint
Private mEndArrowX2Offset As Double   ' horiz Offset from apex of lower Arrow endpoint
Private mEndArrowY2Offset As Double   ' vert Offset from apex of lower Arrow endpoint

Private mStartArrowX1Offset As Double   ' horiz Offset from apex of upper Arrow endpoint
Private mStartArrowY1Offset As Double   ' vert Offset from apex of upper Arrow endpoint
Private mStartArrowX2Offset As Double   ' horiz Offset from apex of lower Arrow endpoint
Private mStartArrowY2Offset As Double   ' vert Offset from apex of lower Arrow endpoint

Private mGaugeX As Double
Private mGaugeY As Double

Private mThicknessOffsetX As Double
Private mThicknessOffsetY As Double

Private mBlank As Boolean
Private mHidden As Boolean

Private mColor As Long
Private mThickness As Long
Private mLineStyle As LineStyles
Private mExtendBefore As Boolean
Private mExtendAfter As Boolean
Private mArrowStartStyle As ArrowStyles
Private mArrowStartLength As Long
Private mArrowStartWidth As Long
Private mArrowStartColor As Long
Private mArrowStartFillColor As Long
Private mArrowStartfillStyle As FillStyles
Private mArrowEndStyle As ArrowStyles
Private mArrowEndLength As Long
Private mArrowEndWidth As Long
Private mArrowEndColor As Long
Private mArrowEndFillColor As Long
Private mArrowEndFillStyle As FillStyles
Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

Private mPoint1 As Point
Private mPoint2 As Point
Private mDirection As Quadrants
Private mVertical As Verticals
Private mLeftToRight As Boolean
Private mBottomToTop As Boolean
Private mA As Double    ' from the equation ...
Private mB As Double    ' ... Y = ax + b
Private mLineAngle As Double    ' slope in radians
Private mLineAngleSin As Double
Private mLineAngleCos As Double

Private mKey As String

Private mPropertyOverrideFlags As LinePropertyOverrideFlags

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mBlank = True

mBoundingRect.isValid = False
mStartArrowRect.isValid = False
mEndArrowRect.isValid = False
mRegionRect.isValid = False

mColor = vbBlack
mThickness = 1
mLineStyle = LineStyles.LineSolid
mExtendBefore = False
mExtendAfter = False
mArrowStartStyle = ArrowStyles.ArrowNone
mArrowStartLength = 10
mArrowStartWidth = 10
mArrowStartColor = vbBlack
mArrowStartFillColor = vbBlack
mArrowStartfillStyle = FillStyles.FillSolid
mArrowEndStyle = ArrowStyles.ArrowNone
mArrowEndLength = 10
mArrowEndWidth = 10
mArrowEndColor = vbBlack
mArrowEndFillColor = vbBlack
mArrowEndFillStyle = FillStyles.FillSolid
End Sub

Private Sub Class_Terminate()
gLogger.Log LogLevelHighDetail, "Line terminated"
Debug.Print "Line terminated"
End Sub

'================================================================================
' IGraphicObject Interface Members
'================================================================================

Private Property Get IGraphicObject_boundingRectangle() As TRectangle

Dim canvasChanges As Long

Dim failpoint As Long
On Error GoTo Err

canvasChanges = mCanvas.CompareTo(mGaugeX, mGaugeY, mRegionRect)

mRegionRect = mCanvas.Boundary
mGaugeX = mCanvas.GaugeX
mGaugeY = mCanvas.GaugeY

If (canvasChanges And CanvasComparisonCodes.GaugeChangedX) Or _
    (canvasChanges And CanvasComparisonCodes.GaugeChangedY) Or _
    (canvasChanges And CanvasComparisonCodes.BoundsChangedX) Or _
    (canvasChanges And CanvasComparisonCodes.BoundsChangedY) Or _
    mSeriesUpdateNumber <> mSeries.SeriesUpdateNumber _
Then
    mSeriesUpdateNumber = mSeries.SeriesUpdateNumber
    calcConstants
    calcEndArrowConstants
    calcStartArrowConstants
    calcCurrentBoundingRect
End If

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
Dim periodrect As TRectangle
Dim rects() As TRectangle

Dim failpoint As Long
On Error GoTo Err

If mBlank Then Exit Property

With periodrect
    .Left = PeriodNumber
    .Right = PeriodNumber + 1
    .Bottom = mRegionRect.Bottom
    .Top = mRegionRect.Top
    .isValid = True
End With

If mVertical <> VerticalNot Then
    Dim vertRect As TRectangle
    
    vertRect = rectIntersection(mBoundingRect, periodrect)
    If vertRect.isValid Then
        
        ReDim rects(0) As TRectangle

        rects(0) = vertRect
    End If
Else
    If PeriodNumber = Int(mRegionRect.Right) Or _
        (Not mLeftToRight And PeriodNumber = mPoint2.X And Not ExtendAfter) Or _
        (Not mLeftToRight And PeriodNumber = mPoint1.X And Not ExtendBefore) _
    Then
        If Thickness <= 1 Then Exit Property
    End If
    
    ReDim rects(0) As TRectangle
    
    With rects(0)
        .Left = PeriodNumber
        If mLeftToRight Then
            If PeriodNumber < mPoint1.X And _
                Not ExtendBefore _
            Then
                .Left = mPoint1.X - 1
            End If
        Else
            If PeriodNumber < mPoint2.X And _
                Not ExtendAfter _
            Then
                .Left = mPoint2.X - 1
            End If
        End If
        If mLeftToRight Then
            If (Not ExtendAfter And PeriodNumber = mPoint2.X) Or _
                PeriodNumber = Int(mRegionRect.Right) _
            Then
                .Right = .Left
            Else
                .Right = PeriodNumber + 1
            End If
        Else
            If (Not ExtendBefore And PeriodNumber = mPoint1.X) Or _
                PeriodNumber = Int(mRegionRect.Right) _
            Then
                .Right = .Left
            Else
                .Right = PeriodNumber + 1
            End If
        End If
        
        If mA > 0 Then
            .Bottom = mA * .Left + mB - mThicknessOffsetY
            .Top = mA * .Right + mB + mThicknessOffsetY
        Else
            .Top = mA * .Left + mB + mThicknessOffsetY
            .Bottom = mA * .Right + mB - mThicknessOffsetY
        End If
        If mLeftToRight Then
            If (Not ExtendBefore And PeriodNumber = mPoint1.X) Then .Left = .Left - mThicknessOffsetX
            If (Not ExtendAfter And PeriodNumber = mPoint2.X) Or _
                PeriodNumber = Int(mRegionRect.Right) _
            Then
                .Right = .Right + mThicknessOffsetX
            End If
        Else
            If (Not ExtendAfter And PeriodNumber = mPoint2.X) Then .Left = .Left - mThicknessOffsetX
            If (Not ExtendBefore And PeriodNumber = mPoint1.X) Or _
                PeriodNumber = Int(mRegionRect.Right) _
            Then
                .Right = .Right + mThicknessOffsetX
            End If
        End If
        .isValid = True
    End With
End If

If Not ArrowEndStyle = ArrowNone Then
    If rectOverlaps(mEndArrowRect, periodrect) Then
        rects(0) = rectUnion(rects(0), mEndArrowRect)
    End If
End If
If Not ArrowStartStyle = ArrowNone Then
    If rectOverlaps(mStartArrowRect, periodrect) Then
        rects(0) = rectUnion(rects(0), mStartArrowRect)
    End If
End If
IGraphicObject_BoundingRectanglesAt = rects

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "IGraphicObject_boundingRectanglesAt" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Property

Private Property Get IGraphicObject_capabilities() As GraphicObjectCapabilities
'IGraphicObject_capabilities = GraphicObjectCapabilities.ComplexBounding Or _
'                            GraphicObjectCapabilities.PartialRedraw
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
Dim j As Long
Dim areasToRedraw() As TRectangle

Dim failpoint As Long
On Error GoTo Err

ReDim areasToRedraw(UBound(areas)) As TRectangle

' amalagamate any adjacent areas
areasToRedraw(0) = areas(0)
For i = 1 To UBound(areas)
    If rectOverlaps(areasToRedraw(j), areas(i)) Then
        areasToRedraw(j) = rectUnion(areasToRedraw(j), areas(i))
    Else
        j = j + 1
        areasToRedraw(j) = areas(i)
    End If
Next

For i = j To 0 Step -1
    Draw areasToRedraw(i)
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "IGraphicObject_draw" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = isExtended
End Property

Private Sub IGraphicObject_finish()
Set mCanvas = Nothing
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
Dim lineX As Double
Dim lineY As Double
Dim hitTestRect As TRectangle

hitTestRect = mBoundingRect
hitTestRect.Left = hitTestRect.Left - mCanvas.ConvertPixelsToLogicalX(HitTestTolerancePixels)
hitTestRect.Right = hitTestRect.Right + mCanvas.ConvertPixelsToLogicalX(HitTestTolerancePixels)
hitTestRect.Bottom = hitTestRect.Bottom - mCanvas.ConvertPixelsToLogicalY(HitTestTolerancePixels)
hitTestRect.Top = hitTestRect.Top + mCanvas.ConvertPixelsToLogicalY(HitTestTolerancePixels)

If Not RectContainsPoint(hitTestRect, X, Y) Then Exit Function

If mA <= 1 Then
    lineY = mA * X + mB
    If Abs(Y - lineY) <= mCanvas.ConvertPixelsToLogicalY(HitTestTolerancePixels) Then IGraphicObject_HitTest = True
Else
    lineX = (Y - mB) / mA
    If Abs(X - lineX) <= mCanvas.ConvertPixelsToLogicalX(HitTestTolerancePixels) Then IGraphicObject_HitTest = True
End If

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
IGraphicObject_noDraw = mHidden Or mBlank
End Property

Private Property Get IGraphicObject_PeriodNumber() As Long
IGraphicObject_PeriodNumber = -1
If isExtended Then Exit Sub
If mPoint1 Is Nothing Then Exit Sub
If Not mPoint1.IsAssigned Then Exit Sub
IGraphicObject_PeriodNumber = Int(mPoint1.XLogical + 0.5)
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
Case "ARROWENDCOLOR"
    initiateRedraw False, False
Case "ARROWENDFILLCOLOR"
    initiateRedraw False, False
Case "ARROWENDFILLSTYLE"
    initiateRedraw False, False
Case "ARROWENDLENGTH"
    initiateRedraw True, True
Case "ARROWENDSTYLE"
    initiateRedraw True, True
Case "ARROWENDWIDTH"
    initiateRedraw True, True
Case "ARROWSTARTCOLOR"
    initiateRedraw False, False
Case "ARROWSTARTFILLCOLOR"
    initiateRedraw False, False
Case "ARROWSTARTFILLSTYLE"
    initiateRedraw False, False
Case "ARROWSTARTLENGTH"
    initiateRedraw True, True
Case "ARROWSTARTSTYLE"
    initiateRedraw True, True
Case "ARROWSTARTWIDTH"
    initiateRedraw True, True
Case "COLOR"
    initiateRedraw False, False
Case "EXTENDAFTER"
    initiateRedraw True, True
Case "EXTENDBEFORE"
    initiateRedraw True, True
Case "EXTENDED"
    initiateRedraw True, True
Case "FIXEDX"
    initiateRedraw True, True
Case "FIXEDY"
    initiateRedraw True, True
Case "INCLUDEINAUTOSCALE"
    initiateRedraw False, False
Case "LINESTYLE"
    initiateRedraw True, False
Case "THICKNESS"
    initiateRedraw True, True
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & "mLocalStyle_PropertyChanged", _
            "Unhandled property change"
End Select

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "mStyle_PropertyChanged" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get ArrowEndColor() As Long
If isPropertyOverrideFlagSet(LineIsSetArrowEndColor) Then
    ArrowEndColor = mArrowEndColor
ElseIf Not mLocalStyle Is Nothing Then
    ArrowEndColor = mLocalStyle.ArrowEndColor
Else
    ArrowEndColor = mSeries.ArrowEndColor
End If
End Property

Public Property Let ArrowEndColor(ByVal value As Long)
If mArrowEndColor = value And isPropertyOverrideFlagSet(LineIsSetArrowEndColor) Then Exit Property
mArrowEndColor = value
setPropertyOverrideFlag LineIsSetArrowEndColor
initiateRedraw False, False
End Property

Public Property Get ArrowEndFillColor() As Long
ArrowEndFillColor = mSeries.ArrowEndFillColor
If Not mLocalStyle Is Nothing Then ArrowEndFillColor = mLocalStyle.ArrowEndFillColor
If isPropertyOverrideFlagSet(LineIsSetArrowEndFillColor) Then ArrowEndFillColor = mArrowEndFillColor
End Property

Public Property Let ArrowEndFillColor(ByVal value As Long)
If mArrowEndFillColor = value And isPropertyOverrideFlagSet(LineIsSetArrowEndFillColor) Then Exit Property
mArrowEndFillColor = value
setPropertyOverrideFlag LineIsSetArrowEndFillColor
initiateRedraw False, False
End Property

Public Property Get ArrowEndFillStyle() As FillStyles
ArrowEndFillStyle = mSeries.ArrowEndFillStyle
If Not mLocalStyle Is Nothing Then ArrowEndFillStyle = mLocalStyle.ArrowEndFillStyle
If isPropertyOverrideFlagSet(LineIsSetArrowEndFillStyle) Then ArrowEndFillStyle = mArrowEndFillStyle
End Property

Public Property Let ArrowEndFillStyle(ByVal value As FillStyles)
If mArrowEndFillStyle = value And isPropertyOverrideFlagSet(LineIsSetArrowEndFillStyle) Then Exit Property
mArrowEndFillStyle = value
setPropertyOverrideFlag LineIsSetArrowEndFillStyle
initiateRedraw False, False
End Property

Public Property Get ArrowEndLength() As Long
ArrowEndLength = mSeries.ArrowEndLength
If Not mLocalStyle Is Nothing Then ArrowEndLength = mLocalStyle.ArrowEndLength
If isPropertyOverrideFlagSet(LineIsSetArrowEndLength) Then ArrowEndLength = mArrowEndLength
End Property

Public Property Let ArrowEndLength(ByVal value As Long)
If mArrowEndLength = value And isPropertyOverrideFlagSet(LineIsSetArrowEndLength) Then Exit Property
mArrowEndLength = value
setPropertyOverrideFlag LineIsSetArrowEndLength
initiateRedraw True, True
End Property

Public Property Get ArrowEndStyle() As ArrowStyles
ArrowEndStyle = mSeries.ArrowEndStyle
If Not mLocalStyle Is Nothing Then ArrowEndStyle = mLocalStyle.ArrowEndStyle
If isPropertyOverrideFlagSet(LineIsSetArrowEndStyle) Then ArrowEndStyle = mArrowEndStyle
End Property

Public Property Let ArrowEndStyle(ByVal value As ArrowStyles)
If mArrowEndStyle = value And isPropertyOverrideFlagSet(LineIsSetArrowEndStyle) Then Exit Property
mArrowEndStyle = value
setPropertyOverrideFlag LineIsSetArrowEndStyle
initiateRedraw True, True
End Property

Public Property Get ArrowEndWidth() As Long
ArrowEndWidth = mSeries.ArrowEndWidth
If Not mLocalStyle Is Nothing Then ArrowEndWidth = mLocalStyle.ArrowEndWidth
If isPropertyOverrideFlagSet(LineIsSetArrowEndWidth) Then ArrowEndWidth = mArrowEndWidth
End Property

Public Property Let ArrowEndWidth(ByVal value As Long)
If mArrowEndWidth = value And isPropertyOverrideFlagSet(LineIsSetArrowEndWidth) Then Exit Property
mArrowEndWidth = value
setPropertyOverrideFlag LineIsSetArrowEndWidth
initiateRedraw True, True
End Property

Public Property Get ArrowStartColor() As Long
ArrowStartColor = mSeries.ArrowStartColor
If Not mLocalStyle Is Nothing Then ArrowStartColor = mLocalStyle.ArrowStartColor
If isPropertyOverrideFlagSet(LineIsSetArrowStartColor) Then ArrowStartColor = mArrowStartColor
End Property

Public Property Let ArrowStartColor(ByVal value As Long)
If mArrowStartColor = value And isPropertyOverrideFlagSet(LineIsSetArrowStartColor) Then Exit Property
mArrowStartColor = value
setPropertyOverrideFlag LineIsSetArrowStartColor
initiateRedraw False, False
End Property

Public Property Get ArrowStartFillColor() As Long
ArrowStartFillColor = mSeries.ArrowStartFillColor
If Not mLocalStyle Is Nothing Then ArrowStartFillColor = mLocalStyle.ArrowStartFillColor
If isPropertyOverrideFlagSet(LineIsSetArrowStartFillColor) Then ArrowStartFillColor = mArrowStartFillColor
End Property

Public Property Let ArrowStartFillColor(ByVal value As Long)
If mArrowStartFillColor = value And isPropertyOverrideFlagSet(LineIsSetArrowStartFillColor) Then Exit Property
mArrowStartFillColor = value
setPropertyOverrideFlag LineIsSetArrowStartFillColor
initiateRedraw False, False
End Property

Public Property Get ArrowStartFillStyle() As FillStyles
ArrowStartFillStyle = mSeries.ArrowStartFillStyle
If Not mLocalStyle Is Nothing Then ArrowStartFillStyle = mLocalStyle.ArrowStartFillStyle
If isPropertyOverrideFlagSet(LineIsSetArrowStartFillStyle) Then ArrowStartFillStyle = mArrowStartfillStyle
End Property

Public Property Let ArrowStartFillStyle(ByVal value As FillStyles)
If mArrowStartfillStyle = value And isPropertyOverrideFlagSet(LineIsSetArrowStartFillStyle) Then Exit Property
mArrowStartfillStyle = value
setPropertyOverrideFlag LineIsSetArrowStartFillStyle
initiateRedraw False, False
End Property

Public Property Get ArrowStartLength() As Long
ArrowStartLength = mSeries.ArrowStartLength
If Not mLocalStyle Is Nothing Then ArrowStartLength = mLocalStyle.ArrowStartLength
If isPropertyOverrideFlagSet(LineIsSetArrowStartLength) Then ArrowStartLength = mArrowStartLength
End Property

Public Property Let ArrowStartLength(ByVal value As Long)
If mArrowStartLength = value And isPropertyOverrideFlagSet(LineIsSetArrowStartLength) Then Exit Property
mArrowStartLength = value
setPropertyOverrideFlag LineIsSetArrowStartLength
initiateRedraw True, True
End Property

Public Property Get ArrowStartStyle() As ArrowStyles
ArrowStartStyle = mSeries.ArrowStartStyle
If Not mLocalStyle Is Nothing Then ArrowStartStyle = mLocalStyle.ArrowStartStyle
If isPropertyOverrideFlagSet(LineIsSetArrowStartStyle) Then ArrowStartStyle = mArrowStartStyle
End Property

Public Property Let ArrowStartStyle(ByVal value As ArrowStyles)
If mArrowStartStyle = value And isPropertyOverrideFlagSet(LineIsSetArrowStartStyle) Then Exit Property
mArrowStartStyle = value
setPropertyOverrideFlag LineIsSetArrowStartStyle
initiateRedraw True, True
End Property

Public Property Get ArrowStartWidth() As Long
ArrowStartWidth = mSeries.ArrowStartWidth
If Not mLocalStyle Is Nothing Then ArrowStartWidth = mLocalStyle.ArrowStartWidth
If isPropertyOverrideFlagSet(LineIsSetArrowStartWidth) Then ArrowStartWidth = mArrowStartWidth
End Property

Public Property Let ArrowStartWidth(ByVal value As Long)
If mArrowStartWidth = value And isPropertyOverrideFlagSet(LineIsSetArrowStartWidth) Then Exit Property
mArrowStartWidth = value
setPropertyOverrideFlag LineIsSetArrowStartWidth
initiateRedraw True, True
End Property

Public Property Get Color() As Long
Color = mSeries.Color
If Not mLocalStyle Is Nothing Then Color = mLocalStyle.Color
If isPropertyOverrideFlagSet(LineIsSetColor) Then Color = mColor
End Property

Public Property Let Color(ByVal value As Long)
If mColor = value And isPropertyOverrideFlagSet(LineIsSetColor) Then Exit Property
mColor = value
setPropertyOverrideFlag LineIsSetColor
initiateRedraw False, False
End Property

Public Property Get Direction() As Quadrants
Direction = mDirection
End Property

Public Property Get ExtendAfter() As Boolean
ExtendAfter = mSeries.ExtendAfter
If Not mLocalStyle Is Nothing Then ExtendAfter = mLocalStyle.ExtendAfter
If isPropertyOverrideFlagSet(LineIsSetExtendAfter) Then ExtendAfter = mExtendAfter
End Property

Public Property Let ExtendAfter(ByVal value As Boolean)
If mExtendAfter = value And isPropertyOverrideFlagSet(LineIsSetExtendAfter) Then Exit Property
mExtendAfter = value
setPropertyOverrideFlag LineIsSetExtendAfter
initiateRedraw True, True
End Property

Public Property Get ExtendBefore() As Boolean
ExtendBefore = mSeries.ExtendBefore
If Not mLocalStyle Is Nothing Then ExtendBefore = mLocalStyle.ExtendBefore
If isPropertyOverrideFlagSet(LineIsSetExtendBefore) Then ExtendBefore = mExtendBefore
End Property

Public Property Let ExtendBefore(ByVal value As Boolean)
If mExtendBefore = value And isPropertyOverrideFlagSet(LineIsSetExtendBefore) Then Exit Property
mExtendBefore = value
setPropertyOverrideFlag LineIsSetExtendBefore
initiateRedraw True, True
End Property

Public Property Get Extended() As Boolean
Extended = mSeries.Extended
If Not mLocalStyle Is Nothing Then Extended = mLocalStyle.Extended
If isPropertyOverrideFlagSet(LineIsSetExtended) Then Extended = mExtended
End Property

Public Property Let Extended(ByVal value As Boolean)
If mExtended = value And isPropertyOverrideFlagSet(LineIsSetExtended) Then Exit Property
mExtended = value
setPropertyOverrideFlag LineIsSetExtended
initiateRedraw True, True
End Property

Public Property Get FixedX() As Boolean
FixedX = mSeries.FixedX
If Not mLocalStyle Is Nothing Then FixedX = mLocalStyle.FixedX
If isPropertyOverrideFlagSet(LineIsSetFixedX) Then FixedX = mFixedX
End Property

Public Property Let FixedX(ByVal value As Boolean)
If mFixedX = value And isPropertyOverrideFlagSet(LineIsSetFixedX) Then Exit Property
mFixedX = value
setPropertyOverrideFlag LineIsSetFixedX
initiateRedraw True, True
End Property

Public Property Get FixedY() As Boolean
FixedY = mSeries.FixedY
If Not mLocalStyle Is Nothing Then FixedY = mLocalStyle.FixedY
If isPropertyOverrideFlagSet(LineIsSetFixedY) Then FixedY = mFixedY
End Property

Public Property Let FixedY(ByVal value As Boolean)
If mFixedY = value And isPropertyOverrideFlagSet(LineIsSetFixedY) Then Exit Property
mFixedY = value
setPropertyOverrideFlag LineIsSetFixedY
initiateRedraw True, True
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Get Hidden() As Boolean
Hidden = mHidden
End Property

Public Property Let Hidden(ByVal value As Boolean)
If mHidden = value Then Exit Property
mHidden = value
initiateRedraw IIf(mHidden, True, False), False
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mSeries.IncludeInAutoscale
If isPropertyOverrideFlagSet(LineIsSetIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value And isPropertyOverrideFlagSet(LineIsSetIncludeInAutoscale) Then Exit Property
mIncludeInAutoscale = value
setPropertyOverrideFlag LineIsSetIncludeInAutoscale
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
If mLayer = value And isPropertyOverrideFlagSet(LineIsSetLayer) Then Exit Property
mLayer = value
setPropertyOverrideFlag LineIsSetLayer
initiateRedraw False, False
End Property

Public Property Get Layer() As LayerNumbers
Layer = mSeries.Layer
If isPropertyOverrideFlagSet(LineIsSetLayer) Then Layer = mLayer
End Property

Public Property Get LineStyle() As LineStyles
LineStyle = mSeries.LineStyle
If Not mLocalStyle Is Nothing Then LineStyle = mLocalStyle.LineStyle
If isPropertyOverrideFlagSet(LineIsSetLineStyle) Then LineStyle = mLineStyle
End Property

Public Property Let LineStyle(ByVal value As LineStyles)
If mLineStyle = value And isPropertyOverrideFlagSet(LineIsSetLineStyle) Then Exit Property
mLineStyle = value
setPropertyOverrideFlag LineIsSetLineStyle
initiateRedraw True, False
End Property

Public Property Get LocalStyle() As LineStyle
Set LocalStyle = mLocalStyle
End Property

Public Property Let LocalStyle( _
                ByVal value As LineStyle)

Set mLocalStyle = value
initiateRedraw True, True
End Property

Public Property Get Point1() As Point
Set Point1 = mPoint1
End Property

Public Property Let Point1(ByVal value As Point)
If value Is Nothing Then
    Set mPoint1 = Nothing
    mBlank = True
    initiateRedraw True, True
    Exit Property
ElseIf Not value.IsAssigned Then
    Set mPoint1 = Nothing
    mBlank = True
    initiateRedraw True, True
    Exit Property
ElseIf Not mPoint1 Is Nothing Then
    If value.Equals(mPoint1) Then Exit Property
End If

Set mPoint1 = value
If mFixedX Then ConvertPositionX
If mFixedY Then ConvertPositionY

If mPoint2 Is Nothing Then Exit Property
If Not mPoint2.IsAssigned Then Exit Property
setDirection
mBlank = False
initiateRedraw True, True
End Property

Public Property Get Point2() As Point
Set Point2 = mPoint2
End Property

Public Property Let Point2(ByVal value As Point)
If value Is Nothing Then
    Set mPoint2 = Nothing
    mBlank = True
    initiateRedraw True, True
    Exit Property
ElseIf Not value.IsAssigned Then
    Set mPoint2 = Nothing
    mBlank = True
    initiateRedraw True, True
    Exit Property
ElseIf Not mPoint2 Is Nothing Then
    If value.Equals(mPoint2) Then Exit Property
End If

Set mPoint2 = value
If mFixedX Then ConvertPositionX
If mFixedY Then ConvertPositionY

If mPoint1 Is Nothing Then Exit Property
If Not mPoint1.IsAssigned Then Exit Property
setDirection
mBlank = False
initiateRedraw True, True
End Property

Public Property Get Thickness() As Long
Thickness = mSeries.Thickness
If Not mLocalStyle Is Nothing Then Thickness = mLocalStyle.Thickness
If isPropertyOverrideFlagSet(LineIsSetThickness) Then Thickness = mThickness
End Property

Public Property Let Thickness(ByVal value As Long)
If mThickness = value And isPropertyOverrideFlagSet(LineIsSetThickness) Then Exit Property
mThickness = value
setPropertyOverrideFlag LineIsSetThickness
initiateRedraw True, True
End Property

Public Property Get Vertical() As Verticals
Vertical = mVertical
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

Friend Sub Initialise( _
                ByVal series As LineSeries, _
                ByVal Key As String, _
                ByVal GOSP As GraphObjServiceProvider, _
                ByVal Canvas As Canvas, _
                ByVal LocalStyle As LineStyle)
                
Set mSeries = series
mKey = Key
Set mGOSP = GOSP
Set mCanvas = Canvas
Set mLocalStyle = LocalStyle

calcCurrentBoundingRect
End Sub

Public Sub SetPosition( _
                ByVal Point1 As Point, _
                ByVal Point2 As Point)
Dim failpoint As Long
On Error GoTo Err

If Point1 Is Nothing Or Point2 Is Nothing Then

    failpoint = 100

    Set mPoint1 = Point1
    Set mPoint2 = Point2
    mBlank = True
    initiateRedraw True, True
    Exit Sub
ElseIf (Not Point1.IsAssigned) Then

    failpoint = 200

    Set mPoint1 = Nothing
    Set mPoint2 = Point2
    mBlank = True
    initiateRedraw True, True
    Exit Sub
ElseIf (Not Point2.IsAssigned) Then

    failpoint = 300

    Set mPoint1 = Point1
    Set mPoint2 = Nothing
    mBlank = True
    initiateRedraw True, True
    Exit Sub
ElseIf (Not mPoint1 Is Nothing) And (Not mPoint2 Is Nothing) Then

    failpoint = 400
    
    If Point1.Equals(mPoint1) And Point2.Equals(mPoint2) Then Exit Sub
End If

failpoint = 500

Set mPoint1 = Point1
Set mPoint2 = Point2

failpoint = 600

If mFixedX Then ConvertPositionX
If mFixedY Then ConvertPositionY

failpoint = 800

setDirection
mBlank = False

failpoint = 900

initiateRedraw True, True

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "setPosition" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcConstants()
Dim x1 As Double
Dim x2 As Double
Dim y1 As Double
Dim y2 As Double

If mBlank Then Exit Sub

mGaugeX = mCanvas.GaugeX
mGaugeY = mCanvas.GaugeY

mThicknessOffsetX = mGaugeX * Thickness / 2
mThicknessOffsetY = mGaugeY * Thickness / 2

If mVertical = VerticalUp Then
    mLineAngle = Pi / 2
    Exit Sub
ElseIf mVertical = VerticalDown Then
    mLineAngle = 3 * Pi / 2
    Exit Sub
End If

With mPoint1
    x1 = .XLogical
    y1 = .YLogical
End With
With mPoint2
    x2 = .XLogical
    y2 = .YLogical
End With

mA = (y2 - y1) / (x2 - x1)
mB = (x2 * y1 - x1 * y2) / (x2 - x1)

Select Case mDirection
Case NE
    mLineAngle = Atn(mA * mGaugeX / mGaugeY)
Case NW
    mLineAngle = Pi + Atn(mA * mGaugeX / mGaugeY)
Case SW
    mLineAngle = Pi + Atn(mA * mGaugeX / mGaugeY)
Case SE
    mLineAngle = 2 * Pi + Atn(mA * mGaugeX / mGaugeY)
End Select

mLineAngleSin = Sin(mLineAngle)
mLineAngleCos = Cos(mLineAngle)

End Sub

Private Sub calcCurrentBoundingRect()
Dim lineInt As TInterval
Dim regionInt As TInterval

Dim x1 As Double
Dim x2 As Double
Dim y1 As Double
Dim y2 As Double

mBoundingRect.isValid = False
If mBlank Then Exit Sub

With mPoint1
    x1 = .XLogical
    y1 = .YLogical
End With
With mPoint2
    x2 = .XLogical
    y2 = .YLogical
End With

With mBoundingRect
    If mVertical <> VerticalNot Then
        If mBottomToTop Then
            If ExtendBefore Then
                lineInt.startValue = MinusInfinityDouble
            Else
                lineInt.startValue = y1
            End If
            If ExtendAfter Then
                lineInt.endValue = PlusInfinityDouble
            Else
                lineInt.endValue = y2
            End If
        Else
            If ExtendAfter Then
                lineInt.startValue = MinusInfinityDouble
            Else
                lineInt.startValue = y2
            End If
            If ExtendBefore Then
                lineInt.endValue = PlusInfinityDouble
            Else
                lineInt.endValue = y1
            End If
        End If
        lineInt.isValid = True
        regionInt = rectGetYInterval(mRegionRect)
        rectSetYInterval mBoundingRect, intIntersection(regionInt, lineInt)
        
        .Left = x1 - mThicknessOffsetX
        .Right = x1 + mThicknessOffsetX
        .Bottom = .Bottom - mThicknessOffsetY
        .Top = .Top + mThicknessOffsetY
        .isValid = True
    Else
        If mLeftToRight Then
            If ExtendBefore Then
                lineInt.startValue = MinusInfinityDouble
            Else
                lineInt.startValue = x1
            End If
            If ExtendAfter Then
                lineInt.endValue = PlusInfinityDouble
            Else
                lineInt.endValue = x2
            End If
        Else
            If ExtendAfter Then
                lineInt.startValue = MinusInfinityDouble
            Else
                lineInt.startValue = x2
            End If
            If ExtendBefore Then
                lineInt.endValue = PlusInfinityDouble
            Else
                lineInt.endValue = x1
            End If
        End If
        lineInt.isValid = True
        regionInt = rectGetXInterval(mRegionRect)
        rectSetXInterval mBoundingRect, intIntersection(regionInt, lineInt)
        
        If Not .isValid Then
            .Bottom = MinusInfinityDouble
            .Top = PlusInfinityDouble
            .Left = MinusInfinityDouble
            .Right = PlusInfinityDouble
            Exit Sub
        End If
            
        Select Case mDirection
        Case NE
            .Top = mA * .Right + mB + mThicknessOffsetY
            .Bottom = mA * .Left + mB - mThicknessOffsetY
        Case NW
            .Top = mA * .Left + mB + mThicknessOffsetY
            .Bottom = mA * .Right + mB - mThicknessOffsetY
        Case SW
            .Bottom = mA * .Left + mB - mThicknessOffsetY
            .Top = mA * .Right + mB + mThicknessOffsetY
        Case SE
            .Bottom = mA * .Right + mB - mThicknessOffsetY
            .Top = mA * .Left + mB + mThicknessOffsetY
        End Select
        
        If Thickness > 1 Then
            .Right = .Right + mThicknessOffsetX
            .Left = .Left - mThicknessOffsetX
        End If
        .isValid = True
    End If
'Debug.Print "Calc bounds : (" & .Left & "," & (mA * .Left + mB) & ")-(" & .Right & "," & (mA * .Right + mB) & ")"
End With

If ArrowEndStyle <> ArrowNone Then mBoundingRect = rectUnion(mBoundingRect, mEndArrowRect)
If ArrowStartStyle <> ArrowNone Then mBoundingRect = rectUnion(mBoundingRect, mStartArrowRect)

End Sub

Private Sub calcEndArrowConstants()
Dim hyp As Double
Dim halfAngle As Double
Dim X As Double
Dim Y As Double

If ArrowEndStyle = ArrowNone Then Exit Sub
If mBlank Then Exit Sub

hyp = Sqr(ArrowEndLength * ArrowEndLength + (ArrowEndWidth * ArrowEndWidth) / 4#)
halfAngle = Atn(ArrowEndWidth / (2 * ArrowEndLength))
mEndArrowX1Offset = -hyp * Cos(mLineAngle - halfAngle) * mGaugeX
mEndArrowY1Offset = -hyp * Sin(mLineAngle - halfAngle) * mGaugeY
mEndArrowX2Offset = -hyp * Sin(Pi / 2 - mLineAngle - halfAngle) * mGaugeX
mEndArrowY2Offset = -hyp * Cos(Pi / 2 - mLineAngle - halfAngle) * mGaugeY

With mPoint2
    X = .XLogical
    Y = .YLogical
End With

With mEndArrowRect
    Select Case mDirection
    Case NE
        .Bottom = Y + mEndArrowY2Offset - mThicknessOffsetY
        .Left = X + mEndArrowX1Offset - mThicknessOffsetX
        .Top = Y + IIf(mEndArrowY1Offset > 0, mEndArrowY1Offset, 0) + mThicknessOffsetY
        .Right = X + IIf(mEndArrowX2Offset > 0, mEndArrowX2Offset, 0) + mThicknessOffsetX
        .isValid = True
    Case NW
        .Bottom = Y + mEndArrowY1Offset - mThicknessOffsetY
        .Left = X + IIf(mEndArrowX1Offset < 0, mEndArrowX1Offset, 0) - mThicknessOffsetX
        .Top = Y + IIf(mEndArrowY2Offset > 0, mEndArrowY2Offset, 0) + mThicknessOffsetY
        .Right = X + mEndArrowX2Offset + mThicknessOffsetX
        .isValid = True
    Case SW
        .Bottom = Y + IIf(mEndArrowY1Offset < 0, mEndArrowY1Offset, 0) - mThicknessOffsetY
        .Left = X + IIf(mEndArrowX2Offset < 0, mEndArrowX2Offset, 0) - mThicknessOffsetX
        .Top = Y + mEndArrowY2Offset + mThicknessOffsetY
        .Right = X + mEndArrowX1Offset + mThicknessOffsetX
        .isValid = True
    Case SE
        .Bottom = Y + IIf(mEndArrowY2Offset < 0, mEndArrowY2Offset, 0) - mThicknessOffsetY
        .Left = X + mEndArrowX2Offset - mThicknessOffsetX
        .Top = Y + mEndArrowY1Offset + mThicknessOffsetY
        .Right = X + IIf(mEndArrowX1Offset > 0, mEndArrowX1Offset, 0) + mThicknessOffsetX
        .isValid = True
    End Select
    'Debug.Assert (.Right - .Left) < 5 And (.Top - .Bottom) < 5
End With
End Sub

Private Sub calcStartArrowConstants()
Dim hyp As Double
Dim halfAngle As Double
Dim X As Double
Dim Y As Double

If ArrowStartStyle = ArrowNone Then Exit Sub
If mBlank Then Exit Sub

hyp = Sqr(ArrowStartLength * ArrowStartLength + ArrowStartWidth * ArrowStartWidth / 4)
halfAngle = Atn(ArrowStartWidth / (2 * ArrowStartLength))
mStartArrowX1Offset = hyp * Cos(mLineAngle - halfAngle) * mGaugeX
mStartArrowY1Offset = hyp * Sin(mLineAngle - halfAngle) * mGaugeY
mStartArrowX2Offset = hyp * Sin(Pi / 2 - mLineAngle - halfAngle) * mGaugeX
mStartArrowY2Offset = hyp * Cos(Pi / 2 - mLineAngle - halfAngle) * mGaugeY

With mPoint1
    X = .XLogical
    Y = .YLogical
End With

With mStartArrowRect
    Select Case mDirection
    Case NE
        .Bottom = Y + IIf(mStartArrowY1Offset < 0, mStartArrowY1Offset, 0) - mThicknessOffsetY
        .Left = X + IIf(mStartArrowX2Offset < 0, mStartArrowX2Offset, 0) - mThicknessOffsetX
        .Top = Y + mStartArrowY2Offset + mThicknessOffsetY
        .Right = X + mStartArrowX1Offset + mThicknessOffsetY
        .isValid = True
    Case NW
        .Bottom = Y + IIf(mEndArrowY2Offset < 0, mEndArrowY2Offset, 0) - mThicknessOffsetY
        .Left = X + mEndArrowX2Offset - mThicknessOffsetX
        .Top = Y + mEndArrowY1Offset + mThicknessOffsetY
        .Right = X + IIf(mEndArrowX1Offset > 0, mEndArrowX1Offset, 0) + mThicknessOffsetX
        .isValid = True
    Case SW
        .Bottom = Y + mStartArrowY2Offset - mThicknessOffsetY
        .Left = X + mStartArrowX1Offset - mThicknessOffsetX
        .Top = Y + IIf(mStartArrowY1Offset > 0, mStartArrowY1Offset, 0) + mThicknessOffsetY
        .Right = X + IIf(mStartArrowX2Offset > 0, mStartArrowX2Offset, 0) + mThicknessOffsetX
        .isValid = True
    Case SE
        .Bottom = Y + mStartArrowY1Offset - mThicknessOffsetY
        .Left = X + IIf(mStartArrowX1Offset < 0, mStartArrowX1Offset, 0) - mThicknessOffsetX
        .Top = Y + IIf(mStartArrowY2Offset > 0, mStartArrowY2Offset, 0) + mThicknessOffsetY
        .Right = X + mStartArrowX2Offset + mThicknessOffsetX
        .isValid = True
        
    End Select
    'Debug.Assert (.Right - .Left) < 5 And (.Top - .Bottom) < 5
End With
End Sub

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As LinePropertyOverrideFlags)
mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)
End Sub

Private Sub ConvertPositionX()
ConvertPosition1X
ConvertPosition2X
End Sub

Private Sub ConvertPositionY()
ConvertPosition1Y
ConvertPosition2Y
End Sub

Private Sub ConvertPosition1X()
If mPoint1 Is Nothing Then Exit Sub
If mFixedX Then
    If mPoint1.CoordinateSystemX = CoordsLogical Then
        mPoint1.CoordinateSystemX = CoordsRelative
    End If
Else
    If mPoint1.CoordinateSystemX <> CoordsLogical Then
        mPoint1.CoordinateSystemX = CoordsLogical
    End If
End If
End Sub

Private Sub ConvertPosition1Y()
If mPoint1 Is Nothing Then Exit Sub
If mFixedY Then
    If mPoint1.CoordinateSystemY = CoordsLogical Then
        mPoint1.CoordinateSystemY = CoordsRelative
    End If
Else
    If mPoint1.CoordinateSystemY <> CoordsLogical Then
        mPoint1.CoordinateSystemY = CoordsLogical
    End If
End If
End Sub

Private Sub ConvertPosition2X()
If mPoint2 Is Nothing Then Exit Sub
If FixedX Then
    If mPoint2.CoordinateSystemX = CoordsLogical Then
        mPoint2.CoordinateSystemX = CoordsRelative
    End If
Else
    If mPoint2.CoordinateSystemX <> CoordsLogical Then
        mPoint2.CoordinateSystemX = CoordsLogical
    End If
End If
End Sub

Private Sub ConvertPosition2Y()
If mPoint2 Is Nothing Then Exit Sub
If FixedY Then
    If mPoint2.CoordinateSystemY = CoordsLogical Then
        mPoint2.CoordinateSystemY = CoordsRelative
    End If
Else
    If mPoint2.CoordinateSystemY <> CoordsLogical Then
        mPoint2.CoordinateSystemY = CoordsLogical
    End If
End If
End Sub

Private Sub Draw(ByRef area As TRectangle)
Dim Right As Double
Dim Left As Double
Dim Bottom As Double
Dim Top As Double
Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double
Dim vertices(3) As Point
Dim slopeThicknessY As Double
Dim xBottom As Double
Dim xTop As Double
Dim includeStartArrow As Boolean
Dim includeEndArrow As Boolean

Dim failpoint As Long
On Error GoTo Err

If mBlank Or mHidden Then Exit Sub
If Not mVisible Then
    Exit Sub
End If

failpoint = 100

If Thickness = 1 Then
    mCanvas.SetPenAttributes Color, _
                            Thickness, _
                            LineStyle, _
                            DrawModes.DrawModeCopyPen
Else
    mCanvas.SetPenAttributes Color, _
                            Thickness, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
End If

failpoint = 200

With mPoint1
    x1 = .XLogical
    y1 = .YLogical
End With

With mPoint2
    x2 = .XLogical
    y2 = .YLogical
End With

If mVertical <> VerticalNot Then

    failpoint = 300

    With area
        If mBottomToTop Then
            If ExtendAfter Then
                Top = .Top
            Else
                If .Top >= y2 Then
                    Top = y2
                    If .Bottom <= y2 Then includeEndArrow = True
                Else
                    Top = .Top
                End If
            End If
            If ExtendBefore Then
                Bottom = .Bottom
            Else
                If .Bottom <= y1 Then
                    Bottom = y1
                    If .Top >= y1 Then includeStartArrow = True
                Else
                    Bottom = .Bottom
                End If
            End If
        Else
            If ExtendAfter Then
                Bottom = .Bottom
            Else
                If .Bottom <= y2 Then
                    Bottom = y2
                    If .Top >= y2 Then includeEndArrow = True
                Else
                    Bottom = .Bottom
                End If
            End If
            If ExtendBefore Then
                Top = .Top
            Else
                If .Top >= y1 Then
                    Top = y1
                    If .Bottom <= y1 Then includeStartArrow = True
                Else
                    Top = .Top
                End If
            End If
        End If
    End With

    failpoint = 400

    mCanvas.DrawLine mCanvas.NewPoint(x1, Bottom), mCanvas.NewPoint(x1, Top)

    failpoint = 500

    mCanvas.DrawPoint mCanvas.NewPoint(x1, Top)
Else
    With area
        
        failpoint = 600

        Left = .Left
        Right = .Right
        
        If mLeftToRight And Not ExtendBefore Then
            If Left < x1 Then Left = x1
            If Right < x1 Then Right = x1
            If Left <= x1 And Right >= x1 Then includeStartArrow = True
        ElseIf Not mLeftToRight And Not ExtendAfter Then
            If Left < x2 Then Left = x2
            If Right < x2 Then Right = x2
            If Left <= x2 And Right >= x2 Then includeEndArrow = True
        End If
            
        failpoint = 700

        If mLeftToRight And Not ExtendAfter Then
            If Left > x2 Then Left = x2
            If Right > x2 Then Right = x2
            If Left <= x2 And Right >= x2 Then includeEndArrow = True
        ElseIf Not mLeftToRight And Not ExtendBefore Then
            If Left > x1 Then Left = x1
            If Right > x1 Then Right = x1
            If Left <= x1 And Right >= x1 Then includeStartArrow = True
        End If
    
        failpoint = 800

        If Left > Int(mRegionRect.Right) Then
            Left = Int(mRegionRect.Right)
        End If
        
        failpoint = 900

        If Right > Int(mRegionRect.Right) Then
            Right = Int(mRegionRect.Right)
        End If
    
    End With
    
    With mCanvas

        failpoint = 1000

        Debug.Assert Left <= Right
        If Left <> Right Then
            If Thickness = 1 Then

                failpoint = 1100

                ' now calculate the X coordinates that ensure the line
                ' does not go outside the Top or Bottom of the area
                ' to be drawn
                If mA = 0 Then
                    xBottom = MinusInfinityDouble
                Else
                    xBottom = (area.Bottom - mB) / mA
                End If
                If mA = 0 Then
                    xTop = PlusInfinityDouble
                Else
                    xTop = (area.Top - mB) / mA
                End If
                If xBottom > Left And xBottom < Right Then
                    If mA >= 0 Then
                        Left = xBottom
                    Else
                        Right = xBottom
                    End If
                End If
                If xTop > Left And xTop < Right Then
                    If mA >= 0 Then
                        Right = xTop
                    Else
                        Left = xTop
                    End If
                End If

                failpoint = 1200

                .DrawLine mCanvas.NewPoint(Left, mA * Left + mB), mCanvas.NewPoint(Right, mA * Right + mB)
            Else

                failpoint = 1300

                If Left = x1 Then
                    .DrawPoint mCanvas.NewPoint(Left, y1)
                ElseIf Left = x2 Then
                    .DrawPoint mCanvas.NewPoint(Left, y2)
                End If
                
                slopeThicknessY = Abs(Thickness / mLineAngleCos)
                                
                failpoint = 1400

                If Left = x1 Or Left = x2 Then
                    Set vertices(0) = .NewPoint(Left + .ConvertPixelsToLogicalX(Thickness) * mLineAngleSin / 2, _
                                            mA * Left + mB - .ConvertPixelsToLogicalY(Thickness) * mLineAngleCos / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                Else
                    Set vertices(0) = .NewPoint(Left, _
                                            mA * Left + mB - .ConvertPixelsToLogicalY(slopeThicknessY) / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                End If
                
                If Left = x1 Or Left = x2 Then
                    Set vertices(1) = .NewPoint(Left - .ConvertPixelsToLogicalX(Thickness) * mLineAngleSin / 2, _
                                            mA * Left + mB + .ConvertPixelsToLogicalY(Thickness) * mLineAngleCos / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                Else
                    Set vertices(1) = .NewPoint(Left, _
                                            mA * Left + mB + .ConvertPixelsToLogicalY(slopeThicknessY) / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                End If
                
                If Right = x1 Or Right = x2 Then
                    Set vertices(2) = .NewPoint(Right - .ConvertPixelsToLogicalX(Thickness) * mLineAngleSin / 2, _
                                            mA * Right + mB + .ConvertPixelsToLogicalY(Thickness) * mLineAngleCos / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                Else
                    Set vertices(2) = .NewPoint(Right, _
                                            mA * Right + mB + .ConvertPixelsToLogicalY(slopeThicknessY) / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                End If
                
                If Right = x1 Or Right = x2 Then
                    Set vertices(3) = .NewPoint(Right + .ConvertPixelsToLogicalX(Thickness) * mLineAngleSin / 2, _
                                            mA * Right + mB - .ConvertPixelsToLogicalY(Thickness) * mLineAngleCos / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                Else
                    Set vertices(3) = .NewPoint(Right, _
                                            mA * Right + mB - .ConvertPixelsToLogicalY(slopeThicknessY) / 2, _
                                            CoordsLogical, _
                                            CoordsLogical)
                End If

                failpoint = 1500

                .SetPenAttributes Color, _
                                1, _
                                LineStyles.LineSolid, _
                                DrawModes.DrawModeCopyPen

                failpoint = 1600

                .SetBrushAttributes Color, _
                                FillStyles.FillSolid
                

                failpoint = 1700

                .DrawPolygon vertices
                
                failpoint = 1800

                .SetPenAttributes Color, _
                                Thickness, _
                                LineStyles.LineSolid, _
                                DrawModes.DrawModeCopyPen

                failpoint = 1900

                If Right = x1 Then
                    .DrawPoint mCanvas.NewPoint(Right, y1)
                ElseIf Right = x2 Then
                    .DrawPoint mCanvas.NewPoint(Right, y2)
                End If
                
                Debug.Assert IIf(mLeftToRight, Right <= x2, True)
                Debug.Assert IIf(mLeftToRight, True, Right <= x1)
            End If
        End If
    End With
End If

failpoint = 2000

If ArrowEndStyle <> ArrowNone And includeEndArrow Then
    drawArrowEnd area
End If

failpoint = 2100

If ArrowStartStyle <> ArrowNone And includeStartArrow Then
    drawArrowStart area
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "draw" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


    
End Sub

Private Sub drawArrowEnd(ByRef area As TRectangle)
Dim X As Double
Dim Y As Double
Dim vertices(2) As Point

With mPoint2
    X = .XLogical
    Y = .YLogical
End With

If Not rectOverlaps(mEndArrowRect, area) Then Exit Sub

Select Case ArrowEndStyle
Case ArrowNone

Case ArrowSingleOpen
    mCanvas.SetPenAttributes ArrowEndColor, _
                            Thickness, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
    
    mCanvas.DrawLine mCanvas.NewPoint(X, Y), mCanvas.NewPoint(X + mEndArrowX1Offset, Y + mEndArrowY1Offset)
    mCanvas.DrawLine mCanvas.NewPoint(X, Y), mCanvas.NewPoint(X + mEndArrowX2Offset, Y + mEndArrowY2Offset)
Case ArrowDoubleOpen

Case ArrowClosed
    mCanvas.SetPenAttributes ArrowEndColor, _
                            1, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
    
    With mCanvas
        .SetBrushAttributes ArrowEndFillColor, ArrowEndFillStyle
        
        Set vertices(0) = .NewPoint(X, Y)
        Set vertices(1) = .NewPoint(X + mEndArrowX1Offset, _
                                    Y + mEndArrowY1Offset)
        Set vertices(2) = .NewPoint(X + mEndArrowX2Offset, _
                                    Y + mEndArrowY2Offset)
        .DrawPolygon vertices
    End With
Case ArrowSingleBar

Case ArrowDoubleBar

Case ArrowLollipop

Case ArrowDiamond

Case ArrowBarb

End Select

End Sub

Private Sub drawArrowStart(ByRef area As TRectangle)
Dim X As Double
Dim Y As Double
Dim vertices(2) As Point

With mPoint1
    X = .XLogical
    Y = .YLogical
End With

If Not rectOverlaps(mStartArrowRect, area) Then Exit Sub

Select Case ArrowStartStyle
Case ArrowNone

Case ArrowSingleOpen
    mCanvas.SetPenAttributes ArrowStartColor, _
                            Thickness, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
    
    mCanvas.DrawLine mCanvas.NewPoint(X, Y), mCanvas.NewPoint(X + mStartArrowX1Offset, Y + mStartArrowY1Offset)
    mCanvas.DrawLine mCanvas.NewPoint(X, Y), mCanvas.NewPoint(X + mStartArrowX2Offset, Y + mStartArrowY2Offset)
Case ArrowDoubleOpen

Case ArrowClosed
    mCanvas.SetPenAttributes ArrowStartColor, _
                            1, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
    
    With mCanvas
        .SetBrushAttributes ArrowStartFillColor, ArrowStartFillStyle
        
        Set vertices(0) = .NewPoint(X, Y)
        Set vertices(1) = .NewPoint(X + mStartArrowX1Offset, _
                                    Y + mStartArrowY1Offset)
        Set vertices(2) = .NewPoint(X + mStartArrowX2Offset, _
                                    Y + mStartArrowY2Offset)
        .DrawPolygon vertices
    End With
Case ArrowSingleBar

Case ArrowDoubleBar

Case ArrowLollipop

Case ArrowDiamond

Case ArrowBarb

End Select

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
    calcConstants
    calcEndArrowConstants
    calcStartArrowConstants
    calcCurrentBoundingRect
End If
mGOSP.ObjectChanged mHandle, mBoundingRect
End Sub

Private Function isExtended() As Boolean
isExtended = Extended Or FixedX
End Function

Private Function isPropertyOverrideFlagSet( _
                ByVal flag As LinePropertyOverrideFlags) As Boolean
isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)
End Function

Private Sub setDirection()
Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double

With mPoint1
    x1 = .XLogical
    y1 = .YLogical
End With

With mPoint2
    x2 = .XLogical
    y2 = .YLogical
End With

If x1 = x2 Then
    If y1 <= y2 Then
        mVertical = VerticalUp
    Else
        mVertical = VerticalDown
    End If
Else
    mVertical = VerticalNot
End If
If x1 <= x2 Then
    If y1 <= y2 Then
        mDirection = NE
        mLeftToRight = True
        mBottomToTop = True
    Else
        mDirection = SE
        mLeftToRight = True
        mBottomToTop = False
    End If
Else
    If y1 <= y2 Then
        mDirection = NW
        mLeftToRight = False
        mBottomToTop = True
    Else
        mDirection = SW
        mLeftToRight = False
        mBottomToTop = False
    End If
End If
End Sub

Private Sub setPropertyOverrideFlag( _
                ByVal flag As LinePropertyOverrideFlags)
mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)
End Sub

Private Sub undraw()

If Not mVisible Then Exit Sub
mGOSP.ObjectUndrawn mHandle, mBoundingRect
Exit Sub

End Sub


