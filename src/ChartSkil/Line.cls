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

Private mViewport As Viewport
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

Private mBoundingRectInViewport As TRectangle

Private mViewportBoundary As TRectangle
Private mTruncatedViewportBoundary As TRectangle

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
Private mOffset1 As Size
Private mOffset2 As Size

Private mPoint1 As Point
Private mPoint2 As Point
Private mX1 As Double
Private mY1 As Double
Private mX2 As Double
Private mY2 As Double

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

Private mPropertyOverrideFlags As LinePropertyFlags

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
mTruncatedViewportBoundary.isValid = False

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
'gLogger.Log LogLevelHighDetail, "Line terminated"
Debug.Print "Line terminated"
End Sub

'================================================================================
' IGraphicObject Interface Members
'================================================================================

Private Property Get IGraphicObject_boundingRectangle() As TRectangle

Dim viewportChanges As ViewportComparisonCodes

Const ProcName As String = "IGraphicObject_boundingRectangle"
Dim failpoint As String
On Error GoTo Err

viewportChanges = mViewport.CompareTo(mGaugeX, mGaugeY, mViewportBoundary)

mViewportBoundary = mViewport.Boundary
mTruncatedViewportBoundary = getTruncatedViewportBoundary
mGaugeX = mViewport.GaugeX
mGaugeY = mViewport.GaugeY

If (viewportChanges And ViewportComparisonCodes.GaugeChangedX) Or _
    (viewportChanges And ViewportComparisonCodes.GaugeChangedY) Or _
    (viewportChanges And ViewportComparisonCodes.BoundsChangedX) Or _
    (viewportChanges And ViewportComparisonCodes.BoundsChangedY) Or _
    mSeriesUpdateNumber <> mSeries.SeriesUpdateNumber _
Then
    mSeriesUpdateNumber = mSeries.SeriesUpdateNumber
    getCoordinatesLogical1
    getCoordinatesLogical2
    calcConstants
    calcEndArrowConstants
    calcStartArrowConstants
    calcCurrentBoundingRect
    calcBoundingRectInViewport
End If

IGraphicObject_boundingRectangle = mBoundingRect

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get IGraphicObject_BoundingRectangleInViewport() As TRectangle
Const ProcName As String = "IGraphicObject_BoundingRectangleInViewport"
Dim failpoint As String
On Error GoTo Err

IGraphicObject_BoundingRectangleInViewport = mBoundingRectInViewport

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get IGraphicObject_BoundingRectanglesAt(ByVal PeriodNumber As Long) As TRectangle()
Dim periodrect As TRectangle
Dim rects() As TRectangle

Const ProcName As String = "IGraphicObject_BoundingRectanglesAt"
Dim failpoint As String
On Error GoTo Err

If mBlank Then Exit Property

With periodrect
    .Left = PeriodNumber
    .Right = PeriodNumber + 1
    .Bottom = mTruncatedViewportBoundary.Bottom
    .Top = mTruncatedViewportBoundary.Top
    .isValid = True
End With

If mVertical <> VerticalNot Then
    Dim vertRect As TRectangle
    
    vertRect = RectIntersection(mBoundingRect, periodrect)
    If vertRect.isValid Then
        
        ReDim rects(0) As TRectangle

        rects(0) = vertRect
    End If
Else
    If PeriodNumber = Int(mTruncatedViewportBoundary.Right) Or _
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
                PeriodNumber = Int(mTruncatedViewportBoundary.Right) _
            Then
                .Right = .Left
            Else
                .Right = PeriodNumber + 1
            End If
        Else
            If (Not ExtendBefore And PeriodNumber = mPoint1.X) Or _
                PeriodNumber = Int(mTruncatedViewportBoundary.Right) _
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
                PeriodNumber = Int(mTruncatedViewportBoundary.Right) _
            Then
                .Right = .Right + mThicknessOffsetX
            End If
        Else
            If (Not ExtendAfter And PeriodNumber = mPoint2.X) Then .Left = .Left - mThicknessOffsetX
            If (Not ExtendBefore And PeriodNumber = mPoint1.X) Or _
                PeriodNumber = Int(mTruncatedViewportBoundary.Right) _
            Then
                .Right = .Right + mThicknessOffsetX
            End If
        End If
        .isValid = True
    End With
End If

If Not ArrowEndStyle = ArrowNone Then
    If RectOverlaps(mEndArrowRect, periodrect) Then
        rects(0) = RectUnion(rects(0), mEndArrowRect)
    End If
End If
If Not ArrowStartStyle = ArrowNone Then
    If RectOverlaps(mStartArrowRect, periodrect) Then
        rects(0) = RectUnion(rects(0), mStartArrowRect)
    End If
End If
IGraphicObject_BoundingRectanglesAt = rects

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get IGraphicObject_capabilities() As GraphicObjectCapabilities
'IGraphicObject_capabilities = GraphicObjectCapabilities.ComplexBounding Or _
'                            GraphicObjectCapabilities.PartialRedraw
Const ProcName As String = "IGraphicObject_capabilities"
Dim failpoint As String
On Error GoTo Err

IGraphicObject_capabilities = GraphicObjectCapabilities.PartialRedraw

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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

Const ProcName As String = "IGraphicObject_draw"
Dim failpoint As String
On Error GoTo Err

ReDim areasToRedraw(UBound(areas)) As TRectangle

' amalagamate any adjacent areas
areasToRedraw(0) = areas(0)
For i = 1 To UBound(areas)
    If RectOverlaps(areasToRedraw(j), areas(i)) Then
        areasToRedraw(j) = RectUnion(areasToRedraw(j), areas(i))
    Else
        j = j + 1
        areasToRedraw(j) = areas(i)
    End If
Next

For i = j To 0 Step -1
'    mCanvas.SetClippingRegion mViewport.ConvertTRectangleToRect(areasToRedraw(i))
    Draw areasToRedraw(i)
'    mCanvas.ClearClippingRegion
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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

Private Property Let IGraphicObject_Handle(ByVal Value As Long)
mHandle = Value
End Property

Private Property Get IGraphicObject_Handle() As Long
IGraphicObject_Handle = mHandle
End Property

Private Function IGraphicObject_HitTest(ByVal X As Double, ByVal Y As Double) As Boolean
Dim lineX As Double
Dim lineY As Double
Dim hitTestRect As TRectangle

Const ProcName As String = "IGraphicObject_HitTest"
Dim failpoint As String
On Error GoTo Err

hitTestRect = mBoundingRectInViewport
hitTestRect.Left = hitTestRect.Left - mViewport.ConvertPixelsToLogicalX(HitTestTolerancePixels)
hitTestRect.Right = hitTestRect.Right + mViewport.ConvertPixelsToLogicalX(HitTestTolerancePixels)
hitTestRect.Bottom = hitTestRect.Bottom - mViewport.ConvertPixelsToLogicalY(HitTestTolerancePixels)
hitTestRect.Top = hitTestRect.Top + mViewport.ConvertPixelsToLogicalY(HitTestTolerancePixels)

If Not RectContainsPoint(hitTestRect, X, Y) Then Exit Function

If mA <= 1 Then
    lineY = mA * X + mB
    If Abs(Y - lineY) <= mViewport.ConvertPixelsToLogicalY(HitTestTolerancePixels) Then IGraphicObject_HitTest = True
Else
    lineX = (Y - mB) / mA
    If Abs(X - lineX) <= mViewport.ConvertPixelsToLogicalX(HitTestTolerancePixels) Then IGraphicObject_HitTest = True
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Private Property Get IGraphicObject_IncludeInAutoscale() As Boolean
IGraphicObject_IncludeInAutoscale = IncludeInAutoscale
End Property

Private Property Let IGraphicObject_inScope(ByVal Value As Boolean)
mInScope = Value
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
IGraphicObject_PeriodNumber = &H80000000
If isExtended Then Exit Sub
If mPoint1 Is Nothing Then Exit Sub
IGraphicObject_PeriodNumber = Int(mX1)
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
Const ProcName As String = "IGraphicObject_SeriesID"
Dim failpoint As String
On Error GoTo Err

IGraphicObject_SeriesID = mSeries.Id

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get IGraphicObject_Timestamp() As Date
IGraphicObject_Timestamp = 0
End Property

Private Property Get IGraphicObject_VerticalRangeInXinterval(pXInterval As TInterval) As TInterval
IGraphicObject_VerticalRangeInXinterval = RectGetYInterval(calcBoundingRectForXInterval(pXInterval))
End Property

Private Property Let IGraphicObject_visible(ByVal Value As Boolean)
mVisible = Value
End Property

'================================================================================
' mLocalStyle Event Handlers
'================================================================================

Private Sub mLocalStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
Const ProcName As String = "mLocalStyle_PropertyChanged"
Dim failpoint As String
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
Case "LAYER"
    initiateRedraw False, False
Case "LINESTYLE"
    initiateRedraw True, False
Case "OFFSET1"
    initiateRedraw True, True
Case "OFFSET2"
    initiateRedraw True, True
Case "THICKNESS"
    initiateRedraw True, True
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Unhandled property change"
End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get ArrowEndColor() As Long
Const ProcName As String = "ArrowEndColor"
Dim failpoint As String
On Error GoTo Err

ArrowEndColor = mSeries.ArrowEndColor
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyArrowEndColor) Then ArrowEndColor = mLocalStyle.ArrowEndColor
If IsPropertySet(LinePropertyArrowEndColor) Then ArrowEndColor = mArrowEndColor
    
Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndColor(ByVal Value As Long)
Const ProcName As String = "ArrowEndColor"
Dim failpoint As String
On Error GoTo Err

If mArrowEndColor = Value And IsPropertySet(LinePropertyArrowEndColor) Then Exit Property
mArrowEndColor = Value
setPropertyFlag LinePropertyArrowEndColor
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndFillColor() As Long
Const ProcName As String = "ArrowEndFillColor"
Dim failpoint As String
On Error GoTo Err

ArrowEndFillColor = mSeries.ArrowEndFillColor
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyArrowEndFillColor) Then ArrowEndFillColor = mLocalStyle.ArrowEndFillColor
If IsPropertySet(LinePropertyArrowEndFillColor) Then ArrowEndFillColor = mArrowEndFillColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndFillColor(ByVal Value As Long)
Const ProcName As String = "ArrowEndFillColor"
Dim failpoint As String
On Error GoTo Err

If mArrowEndFillColor = Value And IsPropertySet(LinePropertyArrowEndFillColor) Then Exit Property
mArrowEndFillColor = Value
setPropertyFlag LinePropertyArrowEndFillColor
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndFillStyle() As FillStyles
Const ProcName As String = "ArrowEndFillStyle"
Dim failpoint As String
On Error GoTo Err

ArrowEndFillStyle = mSeries.ArrowEndFillStyle
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyArrowEndFillStyle) Then ArrowEndFillStyle = mLocalStyle.ArrowEndFillStyle
If IsPropertySet(LinePropertyArrowEndFillStyle) Then ArrowEndFillStyle = mArrowEndFillStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndFillStyle(ByVal Value As FillStyles)
Const ProcName As String = "ArrowEndFillStyle"
Dim failpoint As String
On Error GoTo Err

If mArrowEndFillStyle = Value And IsPropertySet(LinePropertyArrowEndFillStyle) Then Exit Property
mArrowEndFillStyle = Value
setPropertyFlag LinePropertyArrowEndFillStyle
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndLength() As Long
Const ProcName As String = "ArrowEndLength"
Dim failpoint As String
On Error GoTo Err

ArrowEndLength = mSeries.ArrowEndLength
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyArrowEndLength) Then ArrowEndLength = mLocalStyle.ArrowEndLength
If IsPropertySet(LinePropertyArrowEndLength) Then ArrowEndLength = mArrowEndLength

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndLength(ByVal Value As Long)
Const ProcName As String = "ArrowEndLength"
Dim failpoint As String
On Error GoTo Err

If mArrowEndLength = Value And IsPropertySet(LinePropertyArrowEndLength) Then Exit Property
mArrowEndLength = Value
setPropertyFlag LinePropertyArrowEndLength
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndStyle() As ArrowStyles
Const ProcName As String = "ArrowEndStyle"
Dim failpoint As String
On Error GoTo Err

ArrowEndStyle = mSeries.ArrowEndStyle
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyArrowEndStyle) Then ArrowEndStyle = mLocalStyle.ArrowEndStyle
If IsPropertySet(LinePropertyArrowEndStyle) Then ArrowEndStyle = mArrowEndStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndStyle(ByVal Value As ArrowStyles)
Const ProcName As String = "ArrowEndStyle"
Dim failpoint As String
On Error GoTo Err

If mArrowEndStyle = Value And IsPropertySet(LinePropertyArrowEndStyle) Then Exit Property
mArrowEndStyle = Value
setPropertyFlag LinePropertyArrowEndStyle
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndWidth() As Long
Const ProcName As String = "ArrowEndWidth"
Dim failpoint As String
On Error GoTo Err

ArrowEndWidth = mSeries.ArrowEndWidth
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyArrowEndWidth) Then ArrowEndWidth = mLocalStyle.ArrowEndWidth
If IsPropertySet(LinePropertyArrowEndWidth) Then ArrowEndWidth = mArrowEndWidth

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowEndWidth(ByVal Value As Long)
Const ProcName As String = "ArrowEndWidth"
Dim failpoint As String
On Error GoTo Err

If mArrowEndWidth = Value And IsPropertySet(LinePropertyArrowEndWidth) Then Exit Property
mArrowEndWidth = Value
setPropertyFlag LinePropertyArrowEndWidth
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartColor() As Long
Const ProcName As String = "ArrowStartColor"
Dim failpoint As String
On Error GoTo Err

ArrowStartColor = mSeries.ArrowStartColor
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyArrowStartColor) Then ArrowStartColor = mLocalStyle.ArrowStartColor
If IsPropertySet(LinePropertyArrowStartColor) Then ArrowStartColor = mArrowStartColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartColor(ByVal Value As Long)
Const ProcName As String = "ArrowStartColor"
Dim failpoint As String
On Error GoTo Err

If mArrowStartColor = Value And IsPropertySet(LinePropertyArrowStartColor) Then Exit Property
mArrowStartColor = Value
setPropertyFlag LinePropertyArrowStartColor
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartFillColor() As Long
Const ProcName As String = "ArrowStartFillColor"
Dim failpoint As String
On Error GoTo Err

ArrowStartFillColor = mSeries.ArrowStartFillColor
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyArrowStartFillColor) Then ArrowStartFillColor = mLocalStyle.ArrowStartFillColor
If IsPropertySet(LinePropertyArrowStartFillColor) Then ArrowStartFillColor = mArrowStartFillColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartFillColor(ByVal Value As Long)
Const ProcName As String = "ArrowStartFillColor"
Dim failpoint As String
On Error GoTo Err

If mArrowStartFillColor = Value And IsPropertySet(LinePropertyArrowStartFillColor) Then Exit Property
mArrowStartFillColor = Value
setPropertyFlag LinePropertyArrowStartFillColor
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartFillStyle() As FillStyles
Const ProcName As String = "ArrowStartFillStyle"
Dim failpoint As String
On Error GoTo Err

ArrowStartFillStyle = mSeries.ArrowStartFillStyle
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyArrowStartFillStyle) Then ArrowStartFillStyle = mLocalStyle.ArrowStartFillStyle
If IsPropertySet(LinePropertyArrowStartFillStyle) Then ArrowStartFillStyle = mArrowStartfillStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartFillStyle(ByVal Value As FillStyles)
Const ProcName As String = "ArrowStartFillStyle"
Dim failpoint As String
On Error GoTo Err

If mArrowStartfillStyle = Value And IsPropertySet(LinePropertyArrowStartFillStyle) Then Exit Property
mArrowStartfillStyle = Value
setPropertyFlag LinePropertyArrowStartFillStyle
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartLength() As Long
Const ProcName As String = "ArrowStartLength"
Dim failpoint As String
On Error GoTo Err

ArrowStartLength = mSeries.ArrowStartLength
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyArrowStartLength) Then ArrowStartLength = mLocalStyle.ArrowStartLength
If IsPropertySet(LinePropertyArrowStartLength) Then ArrowStartLength = mArrowStartLength

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartLength(ByVal Value As Long)
Const ProcName As String = "ArrowStartLength"
Dim failpoint As String
On Error GoTo Err

If mArrowStartLength = Value And IsPropertySet(LinePropertyArrowStartLength) Then Exit Property
mArrowStartLength = Value
setPropertyFlag LinePropertyArrowStartLength
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartStyle() As ArrowStyles
Const ProcName As String = "ArrowStartStyle"
Dim failpoint As String
On Error GoTo Err

ArrowStartStyle = mSeries.ArrowStartStyle
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyArrowStartStyle) Then ArrowStartStyle = mLocalStyle.ArrowStartStyle
If IsPropertySet(LinePropertyArrowStartStyle) Then ArrowStartStyle = mArrowStartStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartStyle(ByVal Value As ArrowStyles)
Const ProcName As String = "ArrowStartStyle"
Dim failpoint As String
On Error GoTo Err

If mArrowStartStyle = Value And IsPropertySet(LinePropertyArrowStartStyle) Then Exit Property
mArrowStartStyle = Value
setPropertyFlag LinePropertyArrowStartStyle
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartWidth() As Long
Const ProcName As String = "ArrowStartWidth"
Dim failpoint As String
On Error GoTo Err

ArrowStartWidth = mSeries.ArrowStartWidth
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyArrowStartWidth) Then ArrowStartWidth = mLocalStyle.ArrowStartWidth
If IsPropertySet(LinePropertyArrowStartWidth) Then ArrowStartWidth = mArrowStartWidth

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ArrowStartWidth(ByVal Value As Long)
Const ProcName As String = "ArrowStartWidth"
Dim failpoint As String
On Error GoTo Err

If mArrowStartWidth = Value And IsPropertySet(LinePropertyArrowStartWidth) Then Exit Property
mArrowStartWidth = Value
setPropertyFlag LinePropertyArrowStartWidth
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Color() As Long
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

Color = mSeries.Color
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyColor) Then Color = mLocalStyle.Color
If IsPropertySet(LinePropertyColor) Then Color = mColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Color(ByVal Value As Long)
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

If mColor = Value And IsPropertySet(LinePropertyColor) Then Exit Property
mColor = Value
setPropertyFlag LinePropertyColor
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Direction() As Quadrants
Direction = mDirection
End Property

Public Property Get ExtendAfter() As Boolean
Const ProcName As String = "ExtendAfter"
Dim failpoint As String
On Error GoTo Err

ExtendAfter = mSeries.ExtendAfter
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyExtendAfter) Then ExtendAfter = mLocalStyle.ExtendAfter
If IsPropertySet(LinePropertyExtendAfter) Then ExtendAfter = mExtendAfter

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ExtendAfter(ByVal Value As Boolean)
Const ProcName As String = "ExtendAfter"
Dim failpoint As String
On Error GoTo Err

If mExtendAfter = Value And IsPropertySet(LinePropertyExtendAfter) Then Exit Property
mExtendAfter = Value
setPropertyFlag LinePropertyExtendAfter
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ExtendBefore() As Boolean
Const ProcName As String = "ExtendBefore"
Dim failpoint As String
On Error GoTo Err

ExtendBefore = mSeries.ExtendBefore
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyExtendBefore) Then ExtendBefore = mLocalStyle.ExtendBefore
If IsPropertySet(LinePropertyExtendBefore) Then ExtendBefore = mExtendBefore

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ExtendBefore(ByVal Value As Boolean)
Const ProcName As String = "ExtendBefore"
Dim failpoint As String
On Error GoTo Err

If mExtendBefore = Value And IsPropertySet(LinePropertyExtendBefore) Then Exit Property
mExtendBefore = Value
setPropertyFlag LinePropertyExtendBefore
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Extended() As Boolean
Const ProcName As String = "Extended"
Dim failpoint As String
On Error GoTo Err

Extended = mSeries.Extended
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyExtended) Then Extended = mLocalStyle.Extended
If IsPropertySet(LinePropertyExtended) Then Extended = mExtended

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Extended(ByVal Value As Boolean)
Const ProcName As String = "Extended"
Dim failpoint As String
On Error GoTo Err

If mExtended = Value And IsPropertySet(LinePropertyExtended) Then Exit Property
mExtended = Value
setPropertyFlag LinePropertyExtended
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get FixedX() As Boolean
Const ProcName As String = "FixedX"
Dim failpoint As String
On Error GoTo Err

FixedX = mSeries.FixedX
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyFixedX) Then FixedX = mLocalStyle.FixedX
If IsPropertySet(LinePropertyFixedX) Then FixedX = mFixedX

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let FixedX(ByVal Value As Boolean)
Const ProcName As String = "FixedX"
Dim failpoint As String
On Error GoTo Err

If mFixedX = Value And IsPropertySet(LinePropertyFixedX) Then Exit Property
mFixedX = Value
setPropertyFlag LinePropertyFixedX
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get FixedY() As Boolean
Const ProcName As String = "FixedY"
Dim failpoint As String
On Error GoTo Err

FixedY = mSeries.FixedY
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyFixedY) Then FixedY = mLocalStyle.FixedY
If IsPropertySet(LinePropertyFixedY) Then FixedY = mFixedY

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let FixedY(ByVal Value As Boolean)
Const ProcName As String = "FixedY"
Dim failpoint As String
On Error GoTo Err

If mFixedY = Value And IsPropertySet(LinePropertyFixedY) Then Exit Property
mFixedY = Value
setPropertyFlag LinePropertyFixedY
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Get Hidden() As Boolean
Hidden = mHidden
End Property

Public Property Let Hidden(ByVal Value As Boolean)
Const ProcName As String = "Hidden"
Dim failpoint As String
On Error GoTo Err

If mHidden = Value Then Exit Property
mHidden = Value
initiateRedraw IIf(mHidden, True, False), False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

IncludeInAutoscale = mSeries.IncludeInAutoscale
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyIncludeInAutoscale) Then IncludeInAutoscale = mLocalStyle.IncludeInAutoscale
If IsPropertySet(LinePropertyIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

If mIncludeInAutoscale = Value And IsPropertySet(LinePropertyIncludeInAutoscale) Then Exit Property
mIncludeInAutoscale = Value
setPropertyFlag LinePropertyIncludeInAutoscale
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get InScope() As Boolean
InScope = mInScope
End Property

Public Property Let IsSelectable(ByVal Value As Boolean)
mIsSelectable = Value
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

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"
Dim failpoint As String
On Error GoTo Err

If mLayer = Value And IsPropertySet(LinePropertyLayer) Then Exit Property
mLayer = Value
setPropertyFlag LinePropertyLayer
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
Dim failpoint As String
On Error GoTo Err

Layer = mSeries.Layer
If IsPropertySet(LinePropertyLayer) Then Layer = mLayer

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LineStyle() As LineStyles
Const ProcName As String = "LineStyle"
Dim failpoint As String
On Error GoTo Err

LineStyle = mSeries.LineStyle
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyLineStyle) Then LineStyle = mLocalStyle.LineStyle
If IsPropertySet(LinePropertyLineStyle) Then LineStyle = mLineStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let LineStyle(ByVal Value As LineStyles)
Const ProcName As String = "LineStyle"
Dim failpoint As String
On Error GoTo Err

If mLineStyle = Value And IsPropertySet(LinePropertyLineStyle) Then Exit Property
mLineStyle = Value
setPropertyFlag LinePropertyLineStyle
initiateRedraw True, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LocalStyle() As LineStyle
Set LocalStyle = mLocalStyle
End Property

Public Property Let LocalStyle( _
                ByVal Value As LineStyle)

Const ProcName As String = "LocalStyle"
Dim failpoint As String
On Error GoTo Err

Set mLocalStyle = Value
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Offset1() As Size
Set Offset1 = mOffset1
End Property

Public Property Let Offset1(ByVal Value As Size)
Const ProcName As String = "Offset1"
Dim failpoint As String
On Error GoTo Err

If Value Is Nothing Then
    Set mOffset1 = Nothing
ElseIf Not mOffset1 Is Nothing Then
    If Value.Equals(mOffset1) Then Exit Property
Else
    Set mOffset1 = Value
End If
If Not mPoint1 Is Nothing Then getCoordinatesLogical1
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Offset2() As Size
Set Offset2 = mOffset2
End Property

Public Property Let Offset2(ByVal Value As Size)
Const ProcName As String = "Offset2"
Dim failpoint As String
On Error GoTo Err

If Value Is Nothing Then
    Set mOffset2 = Nothing
ElseIf Not mOffset2 Is Nothing Then
    If Value.Equals(mOffset2) Then Exit Property
Else
    Set mOffset2 = Value
End If
If Not mPoint2 Is Nothing Then getCoordinatesLogical2
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Point1() As Point
Set Point1 = mPoint1
End Property

Public Property Let Point1(ByVal Value As Point)
Const ProcName As String = "Point1"
On Error GoTo Err

If Value Is Nothing Then
    Set mPoint1 = Nothing
    mBlank = True
    initiateRedraw True, True
    Exit Property
ElseIf Not mPoint1 Is Nothing Then
    If Value.Equals(mPoint1) Then
        If mPoint1.CoordinateSystemX = CoordsLogical And mPoint1.CoordinateSystemY = CoordsLogical Then
            Exit Property
        End If
    End If
End If

Set mPoint1 = Value

If mPoint2 Is Nothing Then Exit Property
mBlank = False
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Point2() As Point
Set Point2 = mPoint2
End Property

Public Property Let Point2(ByVal Value As Point)
Const ProcName As String = "Point2"
On Error GoTo Err

If Value Is Nothing Then
    Set mPoint2 = Nothing
    mBlank = True
    initiateRedraw True, True
    Exit Property
ElseIf Not mPoint2 Is Nothing Then
    If Value.Equals(mPoint2) Then
        If mPoint2.CoordinateSystemX = CoordsLogical And mPoint2.CoordinateSystemY = CoordsLogical Then
            Exit Property
        End If
    End If
End If

Set mPoint2 = Value

If mPoint1 Is Nothing Then Exit Property
mBlank = False
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Thickness() As Long
Const ProcName As String = "Thickness"
Dim failpoint As String
On Error GoTo Err

Thickness = mSeries.Thickness
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(LinePropertyThickness) Then Thickness = mLocalStyle.Thickness
If IsPropertySet(LinePropertyThickness) Then Thickness = mThickness

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Thickness(ByVal Value As Long)
Const ProcName As String = "Thickness"
Dim failpoint As String
On Error GoTo Err

If mThickness = Value And IsPropertySet(LinePropertyThickness) Then Exit Property
mThickness = Value
setPropertyFlag LinePropertyThickness
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "ClearOverrides"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = 0
initiateRedraw True, True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Initialise( _
                ByVal pSeries As LineSeries, _
                ByVal pKey As String, _
                ByVal pGOSP As GraphObjServiceProvider, _
                ByVal pViewport As Viewport, _
                ByVal pLocalStyle As LineStyle)
                
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mSeries = pSeries
mKey = pKey
Set mGOSP = pGOSP
Set mViewport = pViewport
Set mCanvas = mViewport.Canvas
Set mLocalStyle = pLocalStyle

calcCurrentBoundingRect

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub SetPosition( _
                ByVal pPoint1 As Point, _
                ByVal pPoint2 As Point, _
                Optional ByVal pOffset1 As Size, _
                Optional ByVal pOffset2 As Size)
Const ProcName As String = "SetPosition"
On Error GoTo Err

Set mOffset1 = pOffset1
Set mOffset2 = pOffset2

If pPoint1 Is Nothing Or pPoint2 Is Nothing Then
    Set mPoint1 = pPoint1
    Set mPoint2 = pPoint2
    mBlank = True
    initiateRedraw True, True
    
    Exit Sub
    
ElseIf (Not mPoint1 Is Nothing) And (Not mPoint2 Is Nothing) Then
    If pPoint1.Equals(mPoint1) And pPoint2.Equals(mPoint2) Then
        If mPoint1.CoordinateSystemX = CoordsLogical And mPoint1.CoordinateSystemY = CoordsLogical And mPoint2.CoordinateSystemX = CoordsLogical And mPoint2.CoordinateSystemY = CoordsLogical Then
            Exit Sub
        End If
    End If
End If

Set mPoint1 = pPoint1
Set mPoint2 = pPoint2

mBlank = False

initiateRedraw True, True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub adjustRectForLineThickness( _
                ByRef pTRect As TRectangle, _
                ByRef pXInterval As TInterval)
Dim lRect As RECT
Dim t As Long
Const ProcName As String = "adjustRectForLineThickness"
Dim failpoint As String
On Error GoTo Err

lRect = mViewport.ConvertTRectangleToRect(pTRect)
t = Thickness
With lRect
    If .Left <> MinusInfinityLong Then .Left = .Left - Int(t / 2)
    If .Right <> PlusInfinityLong Then .Right = .Right + Int(t / 2) + 1
    If .Bottom <> PlusInfinityLong Then .Bottom = .Bottom + Int(t / 2) + 1
    If .Top <> MinusInfinityLong Then .Top = .Top - Int(t / 2)
End With

pTRect = mViewport.ConvertRectToTRectangle(lRect)
RectSetXInterval pTRect, IntIntersection(pXInterval, RectGetXInterval(pTRect))

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub calcBoundingRectInViewport()
Const ProcName As String = "calcBoundingRectInviewport"
Dim failpoint As String
On Error GoTo Err

mBoundingRectInViewport = calcBoundingRectForXInterval(RectGetXInterval(mViewport.Boundary))
RectSetYInterval mBoundingRectInViewport, IntIntersection(RectGetYInterval(mBoundingRectInViewport), RectGetYInterval(mViewport.Boundary))
Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function calcBoundingRectForXInterval( _
                ByRef pXInterval As TInterval) As TRectangle
Const ProcName As String = "calcBoundingRectForXInterval"
Dim failpoint As String
On Error GoTo Err

If mBlank Then Exit Function

calcBoundingRectForXInterval = calcBoundingRectInXIntervalForLine(pXInterval)
If Not calcBoundingRectForXInterval.isValid Then Exit Function

adjustRectForLineThickness calcBoundingRectForXInterval, pXInterval
If Not calcBoundingRectForXInterval.isValid Then Exit Function
    
If ArrowEndStyle <> ArrowNone Then calcBoundingRectForXInterval = RectUnion(calcBoundingRectForXInterval, mEndArrowRect)
If ArrowStartStyle <> ArrowNone Then calcBoundingRectForXInterval = RectUnion(calcBoundingRectForXInterval, mStartArrowRect)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Private Function calcBoundingRectInXIntervalForLine( _
                ByRef pXInterval As TInterval) As TRectangle
Const ProcName As String = "calcBoundingRectInXIntervalForLine"
Dim failpoint As String
On Error GoTo Err

calcBoundingRectInXIntervalForLine.isValid = True
RectSetXInterval calcBoundingRectInXIntervalForLine, IntIntersection(pXInterval, getHorizontalTInterval)
If Not calcBoundingRectInXIntervalForLine.isValid Then Exit Function

If Vertical Then
    RectSetYInterval calcBoundingRectInXIntervalForLine, getVerticalTInterval
Else
    RectSetYInterval calcBoundingRectInXIntervalForLine, mapDomainInterval(RectGetXInterval(calcBoundingRectInXIntervalForLine))
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Private Sub calcConstants()

Const ProcName As String = "calcConstants"
Dim failpoint As String
On Error GoTo Err

If mBlank Then Exit Sub

mGaugeX = mViewport.GaugeX
mGaugeY = mViewport.GaugeY

mThicknessOffsetX = mGaugeX * Thickness / 2
mThicknessOffsetY = mGaugeY * Thickness / 2

If mVertical = VerticalUp Then
    mLineAngle = Pi / 2
    Exit Sub
ElseIf mVertical = VerticalDown Then
    mLineAngle = 3 * Pi / 2
    Exit Sub
End If

mA = (mY2 - mY1) / (mX2 - mX1)
mB = (mX2 * mY1 - mX1 * mY2) / (mX2 - mX1)

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub calcCurrentBoundingRect()
Dim lTInt As TInterval

Const ProcName As String = "calcCurrentBoundingRect"
Dim failpoint As String
On Error GoTo Err

lTInt.startValue = MinusInfinityDouble
lTInt.endValue = PlusInfinityDouble
lTInt.isValid = True
mBoundingRect = calcBoundingRectForXInterval(lTInt)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub calcEndArrowConstants()
Dim hyp As Double
Dim halfAngle As Double

Const ProcName As String = "calcEndArrowConstants"
Dim failpoint As String
On Error GoTo Err

If ArrowEndStyle = ArrowNone Then Exit Sub
If mBlank Then Exit Sub

hyp = Sqr(ArrowEndLength * ArrowEndLength + (ArrowEndWidth * ArrowEndWidth) / 4#)
halfAngle = Atn(ArrowEndWidth / (2 * ArrowEndLength))
mEndArrowX1Offset = -hyp * Cos(mLineAngle - halfAngle) * mGaugeX
mEndArrowY1Offset = -hyp * Sin(mLineAngle - halfAngle) * mGaugeY
mEndArrowX2Offset = -hyp * Sin(Pi / 2 - mLineAngle - halfAngle) * mGaugeX
mEndArrowY2Offset = -hyp * Cos(Pi / 2 - mLineAngle - halfAngle) * mGaugeY

With mEndArrowRect
    Select Case mDirection
    Case NE
        .Bottom = mY2 + mEndArrowY2Offset - mThicknessOffsetY
        .Left = mX2 + mEndArrowX1Offset - mThicknessOffsetX
        .Top = mY2 + IIf(mEndArrowY1Offset > 0, mEndArrowY1Offset, 0) + mThicknessOffsetY
        .Right = mX2 + IIf(mEndArrowX2Offset > 0, mEndArrowX2Offset, 0) + mThicknessOffsetX
        .isValid = True
    Case NW
        .Bottom = mY2 + mEndArrowY1Offset - mThicknessOffsetY
        .Left = mX2 + IIf(mEndArrowX1Offset < 0, mEndArrowX1Offset, 0) - mThicknessOffsetX
        .Top = mY2 + IIf(mEndArrowY2Offset > 0, mEndArrowY2Offset, 0) + mThicknessOffsetY
        .Right = mX2 + mEndArrowX2Offset + mThicknessOffsetX
        .isValid = True
    Case SW
        .Bottom = mY2 + IIf(mEndArrowY1Offset < 0, mEndArrowY1Offset, 0) - mThicknessOffsetY
        .Left = mX2 + IIf(mEndArrowX2Offset < 0, mEndArrowX2Offset, 0) - mThicknessOffsetX
        .Top = mY2 + mEndArrowY2Offset + mThicknessOffsetY
        .Right = mX2 + mEndArrowX1Offset + mThicknessOffsetX
        .isValid = True
    Case SE
        .Bottom = mY2 + IIf(mEndArrowY2Offset < 0, mEndArrowY2Offset, 0) - mThicknessOffsetY
        .Left = mX2 + mEndArrowX2Offset - mThicknessOffsetX
        .Top = mY2 + mEndArrowY1Offset + mThicknessOffsetY
        .Right = mX2 + IIf(mEndArrowX1Offset > 0, mEndArrowX1Offset, 0) + mThicknessOffsetX
        .isValid = True
    End Select
    'Debug.Assert (.Right - .Left) < 5 And (.Top - .Bottom) < 5
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub calcStartArrowConstants()
Dim hyp As Double
Dim halfAngle As Double

Const ProcName As String = "calcStartArrowConstants"
Dim failpoint As String
On Error GoTo Err

If ArrowStartStyle = ArrowNone Then Exit Sub
If mBlank Then Exit Sub

hyp = Sqr(ArrowStartLength * ArrowStartLength + ArrowStartWidth * ArrowStartWidth / 4)
halfAngle = Atn(ArrowStartWidth / (2 * ArrowStartLength))
mStartArrowX1Offset = hyp * Cos(mLineAngle - halfAngle) * mGaugeX
mStartArrowY1Offset = hyp * Sin(mLineAngle - halfAngle) * mGaugeY
mStartArrowX2Offset = hyp * Sin(Pi / 2 - mLineAngle - halfAngle) * mGaugeX
mStartArrowY2Offset = hyp * Cos(Pi / 2 - mLineAngle - halfAngle) * mGaugeY

With mStartArrowRect
    Select Case mDirection
    Case NE
        .Bottom = mY1 + IIf(mStartArrowY1Offset < 0, mStartArrowY1Offset, 0) - mThicknessOffsetY
        .Left = mX1 + IIf(mStartArrowX2Offset < 0, mStartArrowX2Offset, 0) - mThicknessOffsetX
        .Top = mY1 + mStartArrowY2Offset + mThicknessOffsetY
        .Right = mX1 + mStartArrowX1Offset + mThicknessOffsetY
        .isValid = True
    Case NW
        .Bottom = mY1 + IIf(mEndArrowY2Offset < 0, mEndArrowY2Offset, 0) - mThicknessOffsetY
        .Left = mX1 + mEndArrowX2Offset - mThicknessOffsetX
        .Top = mY1 + mEndArrowY1Offset + mThicknessOffsetY
        .Right = mX1 + IIf(mEndArrowX1Offset > 0, mEndArrowX1Offset, 0) + mThicknessOffsetX
        .isValid = True
    Case SW
        .Bottom = mY1 + mStartArrowY2Offset - mThicknessOffsetY
        .Left = mX1 + mStartArrowX1Offset - mThicknessOffsetX
        .Top = mY1 + IIf(mStartArrowY1Offset > 0, mStartArrowY1Offset, 0) + mThicknessOffsetY
        .Right = mX1 + IIf(mStartArrowX2Offset > 0, mStartArrowX2Offset, 0) + mThicknessOffsetX
        .isValid = True
    Case SE
        .Bottom = mY1 + mStartArrowY1Offset - mThicknessOffsetY
        .Left = mX1 + IIf(mStartArrowX1Offset < 0, mStartArrowX1Offset, 0) - mThicknessOffsetX
        .Top = mY1 + IIf(mStartArrowY2Offset > 0, mStartArrowY2Offset, 0) + mThicknessOffsetY
        .Right = mX1 + mStartArrowX2Offset + mThicknessOffsetX
        .isValid = True
        
    End Select
    'Debug.Assert (.Right - .Left) < 5 And (.Top - .Bottom) < 5
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub ClearProperty( _
                ByVal flag As LinePropertyFlags)
Const ProcName As String = "ClearProperty"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function convertPosition(ByVal pPoint As Point) As Point
Const ProcName As String = "convertPosition"
On Error GoTo Err

If pPoint Is Nothing Then Exit Function
Set convertPosition = pPoint.ChangeCoordinateSystem( _
                                            IIf(FixedX, CoordinateSystems.CoordsRelative, CoordinateSystems.CoordsLogical), _
                                            IIf(FixedY, CoordinateSystems.CoordsRelative, CoordinateSystems.CoordsLogical), _
                                            mViewport)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub Draw(ByRef Area As TRectangle)
Dim includeStartArrow As Boolean
Dim includeEndArrow As Boolean

Const ProcName As String = "Draw"
Dim failpoint As String
On Error GoTo Err

If mBlank Or mHidden Then Exit Sub
If Not mVisible Then
    Exit Sub
End If

includeStartArrow = RectContainsPoint(mTruncatedViewportBoundary, mX1, mY1)
includeEndArrow = RectContainsPoint(mTruncatedViewportBoundary, mX2, mY2)

If isVertical Then
    drawVerticalLine
ElseIf isHorizontal Then
    DrawLine
Else
    If Thickness = 1 Then
        DrawLine
    Else
        drawGeometricLine
    End If
End If

failpoint = 2000

If ArrowEndStyle <> ArrowNone And includeEndArrow Then
    drawArrowEnd Area
End If

failpoint = 2100

If ArrowStartStyle <> ArrowNone And includeStartArrow Then
    drawArrowStart Area
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub drawArrowEnd(ByRef Area As TRectangle)
Dim vertices(2) As Point

Const ProcName As String = "drawArrowEnd"
Dim failpoint As String
On Error GoTo Err

If Not RectOverlaps(mEndArrowRect, Area) Then Exit Sub

Select Case ArrowEndStyle
Case ArrowNone

Case ArrowSingleOpen
    mCanvas.SetPenAttributes ArrowEndColor, _
                            Thickness, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
    
    mCanvas.DrawLine mX2, mY2, mX2 + mEndArrowX1Offset, mY2 + mEndArrowY1Offset
    mCanvas.DrawLine mX2, mY2, mX2 + mEndArrowX2Offset, mY2 + mEndArrowY2Offset
Case ArrowDoubleOpen

Case ArrowClosed
    mCanvas.SetPenAttributes ArrowEndColor, _
                            1, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
    
    With mViewport
        .Canvas.SetBrushAttributes ArrowEndFillColor, ArrowEndFillStyle
        
        Set vertices(0) = gNewPoint(mX2, mY2)
        Set vertices(1) = gNewPoint(mX2 + mEndArrowX1Offset, _
                                    mY2 + mEndArrowY1Offset)
        Set vertices(2) = gNewPoint(mX2 + mEndArrowX2Offset, _
                                    mY2 + mEndArrowY2Offset)
        .Canvas.DrawPolygon .ConvertPointsToWin32Points(vertices)
    End With
Case ArrowSingleBar

Case ArrowDoubleBar

Case ArrowLollipop

Case ArrowDiamond

Case ArrowBarb

End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub drawArrowStart(ByRef Area As TRectangle)
Dim vertices(2) As Point

Const ProcName As String = "drawArrowStart"
Dim failpoint As String
On Error GoTo Err

If Not RectOverlaps(mStartArrowRect, Area) Then Exit Sub

Select Case ArrowStartStyle
Case ArrowNone

Case ArrowSingleOpen
    mCanvas.SetPenAttributes ArrowStartColor, _
                            Thickness, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
    
    mCanvas.DrawLine mX1, mY1, mX1 + mStartArrowX1Offset, mY1 + mStartArrowY1Offset
    mCanvas.DrawLine mX1, mY1, mX1 + mStartArrowX2Offset, mY1 + mStartArrowY2Offset
Case ArrowDoubleOpen

Case ArrowClosed
    mCanvas.SetPenAttributes ArrowStartColor, _
                            1, _
                            LineStyles.LineSolid, _
                            DrawModes.DrawModeCopyPen
    
    With mViewport
        .Canvas.SetBrushAttributes ArrowStartFillColor, ArrowStartFillStyle
        
        Set vertices(0) = gNewPoint(mX1, mY1)
        Set vertices(1) = gNewPoint(mX1 + mStartArrowX1Offset, _
                                    mY1 + mStartArrowY1Offset)
        Set vertices(2) = gNewPoint(mX1 + mStartArrowX2Offset, _
                                    mY1 + mStartArrowY2Offset)
        .Canvas.DrawPolygon .ConvertPointsToWin32Points(vertices)
    End With
Case ArrowSingleBar

Case ArrowDoubleBar

Case ArrowLollipop

Case ArrowDiamond

Case ArrowBarb

End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub drawGeometricLine()
Dim x1 As Double
Dim x2 As Double
Dim vertices(3) As Point

Const ProcName As String = "drawGeometricLine"
Dim failpoint As String
On Error GoTo Err

setDrawingAttributes

x1 = getStartX
x2 = getEndX

mViewport.Canvas.DrawLine x1, mapPoint(x1), _
                        x2, mapPoint(x2)

'With mViewport
'    .Canvas.DrawPoint .NewPoint(x1, mapPoint(x1))
'
'    Set vertices(0) = .NewPoint(x1 + .ConvertPixelsToLogicalX(Thickness) * mLineAngleSin / 2, _
'                            mapPoint(x1) - .ConvertPixelsToLogicalY(Thickness) * mLineAngleCos / 2, _
'                            CoordsLogical, _
'                            CoordsLogical)
'
'    Set vertices(1) = .NewPoint(x1 - .ConvertPixelsToLogicalX(Thickness) * mLineAngleSin / 2, _
'                            mapPoint(x1) + .ConvertPixelsToLogicalY(Thickness) * mLineAngleCos / 2, _
'                            CoordsLogical, _
'                            CoordsLogical)
'
'    Set vertices(2) = .NewPoint(x2 - .ConvertPixelsToLogicalX(Thickness) * mLineAngleSin / 2, _
'                            mapPoint(x2) + .ConvertPixelsToLogicalY(Thickness) * mLineAngleCos / 2, _
'                            CoordsLogical, _
'                            CoordsLogical)
'
'    Set vertices(3) = .NewPoint(x2 + .ConvertPixelsToLogicalX(Thickness) * mLineAngleSin / 2, _
'                            mapPoint(x2) - .ConvertPixelsToLogicalY(Thickness) * mLineAngleCos / 2, _
'                            CoordsLogical, _
'                            CoordsLogical)
'
'    .Canvas.SetPenAttributes Color, _
'                            1, _
'                            LineStyles.LineSolid, _
'                            DrawModes.DrawModeCopyPen
'
'    .Canvas.SetBrushAttributes Color, _
'                                FillStyles.FillSolid
'
'
'    .Canvas.DrawPolygon .ConvertPointsToWin32Points(vertices)
'
'    .Canvas.SetPenAttributes Color, _
'                            Thickness, _
'                            LineStyles.LineSolid, _
'                            DrawModes.DrawModeCopyPen
'
'    .Canvas.DrawPoint .NewPoint(x2, mapPoint(x2))
'
'End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub DrawLine()
Dim x1 As Double
Dim x2 As Double

Const ProcName As String = "DrawLine"
Dim failpoint As String
On Error GoTo Err

setDrawingAttributes

x1 = getStartX
x2 = getEndX

mViewport.Canvas.DrawLine x1, mapPoint(x1), _
                        x2, mapPoint(x2)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub drawVerticalLine()
Dim y1 As Double
Dim y2 As Double

Const ProcName As String = "drawVerticalLine"
Dim failpoint As String
On Error GoTo Err

If Not IntContains(RectGetXInterval(mTruncatedViewportBoundary), mX1) Then Exit Sub

setDrawingAttributes

If ExtendBefore Then
    If mVertical = VerticalUp Then
        y1 = mTruncatedViewportBoundary.Bottom
    Else
        y1 = mTruncatedViewportBoundary.Top
    End If
Else
    y1 = Point1.yLogical(mViewport)
End If

If ExtendAfter Then
    If mVertical = VerticalUp Then
        y2 = mTruncatedViewportBoundary.Top
    Else
        y2 = mTruncatedViewportBoundary.Bottom
    End If
Else
    y2 = Point2.yLogical(mViewport)
End If

mViewport.Canvas.DrawLine mX1, y1, _
                        mX1, y2

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub getCoordinatesLogical1()
Const ProcName As String = "getCoordinatesLogical1"
Dim failpoint As String
On Error GoTo Err

If mPoint1 Is Nothing Then Exit Sub
mX1 = mPoint1.XLogical(mViewport)
mY1 = mPoint1.yLogical(mViewport)
If Not mOffset1 Is Nothing Then
    mX1 = mX1 + mOffset1.WidthLogical(mViewport)
    mY1 = mY1 + mOffset1.HeightLogical(mViewport)
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub getCoordinatesLogical2()
Const ProcName As String = "getCoordinatesLogical2"
Dim failpoint As String
On Error GoTo Err

If mPoint2 Is Nothing Then Exit Sub
mX2 = mPoint2.XLogical(mViewport)
mY2 = mPoint2.yLogical(mViewport)
If Not mOffset2 Is Nothing Then
    mX2 = mX2 + mOffset2.WidthLogical(mViewport)
    mY2 = mY2 + mOffset2.HeightLogical(mViewport)
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function getDomainInViewport( _
                ByRef Area As TRectangle) As TInterval
Dim yRange As TInterval

Const ProcName As String = "getDomainInViewport"
Dim failpoint As String
On Error GoTo Err

getDomainInViewport.isValid = False
If Vertical Then
    If IntContains(RectXIntersection(Area, mTruncatedViewportBoundary), mX1) Then
        getDomainInViewport.startValue = mX1
        getDomainInViewport.endValue = mX1
        getDomainInViewport.isValid = True
    End If
Else
    yRange = mapDomainInterval(RectXIntersection(Area, mTruncatedViewportBoundary))
    getDomainInViewport = inverseMapRangeInterval(IntIntersection(yRange, RectGetYInterval(Area)))
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getEndX() As Double
Const ProcName As String = "getEndX"
Dim failpoint As String
On Error GoTo Err

If ExtendAfter Then
    If mLeftToRight Then
        getEndX = mTruncatedViewportBoundary.Right
    Else
        getEndX = mTruncatedViewportBoundary.Left
    End If
Else
    If mLeftToRight And mX2 > mTruncatedViewportBoundary.Right Then
        getEndX = mTruncatedViewportBoundary.Right
    Else
        getEndX = mX2
    End If
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getHorizontalTInterval() As TInterval
Const ProcName As String = "getHorizontalTInterval"
Dim failpoint As String
On Error GoTo Err

With getHorizontalTInterval
    If mLeftToRight Then
        If ExtendBefore Then
            .startValue = MinusInfinityDouble
        Else
            .startValue = mX1
        End If
        If ExtendAfter Then
            .endValue = PlusInfinityDouble
        Else
            .endValue = mX2
        End If
    Else
        If ExtendAfter Then
            .startValue = MinusInfinityDouble
        Else
            .startValue = mX2
        End If
        If ExtendBefore Then
            .endValue = PlusInfinityDouble
        Else
            .endValue = mX1
        End If
    End If
    
    .isValid = True
End With

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getStartX() As Double
Const ProcName As String = "getStartX"
Dim failpoint As String
On Error GoTo Err

If ExtendBefore Then
    If mLeftToRight Then
        getStartX = mTruncatedViewportBoundary.Left
    Else
        getStartX = mTruncatedViewportBoundary.Right
    End If
Else
    If Not mLeftToRight And mX1 > mTruncatedViewportBoundary.Right Then
        getStartX = mTruncatedViewportBoundary.Right
    Else
        getStartX = mX1
    End If
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getTruncatedViewportBoundary() As TRectangle
Const ProcName As String = "getTruncatedViewportBoundary"
Dim failpoint As String
On Error GoTo Err

getTruncatedViewportBoundary = mViewport.Boundary
getTruncatedViewportBoundary.Right = Int(getTruncatedViewportBoundary.Right)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getVerticalTInterval() As TInterval
Const ProcName As String = "getVerticalTInterval"
Dim failpoint As String
On Error GoTo Err

With getVerticalTInterval
    If mBottomToTop Then
        If ExtendBefore Then
            .startValue = MinusInfinityDouble
        Else
            .startValue = mPoint1.yLogical(mViewport)
        End If
        If ExtendAfter Then
            .endValue = PlusInfinityDouble
        Else
            .endValue = mPoint2.yLogical(mViewport)
        End If
    Else
        If ExtendAfter Then
            .startValue = MinusInfinityDouble
        Else
            .startValue = mPoint2.yLogical(mViewport)
        End If
        If ExtendBefore Then
            .endValue = PlusInfinityDouble
        Else
            .endValue = mPoint1.yLogical(mViewport)
        End If
    End If
    .isValid = True
End With

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getXIntervalInViewport() As TInterval
Const ProcName As String = "getXIntervalInViewport"
Dim failpoint As String
On Error GoTo Err

getXIntervalInViewport = IntIntersection(RectGetXInterval(mTruncatedViewportBoundary), _
                                        getHorizontalTInterval)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getYIntervalInViewport() As TInterval
Const ProcName As String = "getYIntervalInViewport"
Dim failpoint As String
On Error GoTo Err

getYIntervalInViewport = IntIntersection(RectGetYInterval(mTruncatedViewportBoundary), _
                                        getVerticalTInterval)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function isHorizontal() As Boolean
isHorizontal = (mA = 0)
End Function

Private Function isVertical() As Boolean
isVertical = (mVertical <> VerticalNot)
End Function

Private Sub initiateRedraw( _
                ByVal undrawCurrentImage As Boolean, _
                ByVal recalcBoundaries As Boolean)
Const ProcName As String = "initiateRedraw"
Dim failpoint As String
On Error GoTo Err

If mHandle = 0 Then
    ' haven't yet been added to chart region
    Exit Sub
End If
If undrawCurrentImage Then undraw

If mBlank Then Exit Sub

If recalcBoundaries Then
    Set mPoint1 = convertPosition(mPoint1)
    Set mPoint2 = convertPosition(mPoint2)
    getCoordinatesLogical1
    getCoordinatesLogical2
    
    setDirection
    calcConstants
    calcEndArrowConstants
    calcStartArrowConstants
    calcCurrentBoundingRect
    calcBoundingRectInViewport
End If
mGOSP.ObjectChanged mHandle, mBoundingRect

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function InverseMapPoint( _
                ByVal Y As Double) As Double
Const ProcName As String = "InverseMapPoint"
Dim failpoint As String
On Error GoTo Err

If Vertical Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                            "Invalid for vertical line"
                 
InverseMapPoint = (Y - mB) / mA

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function inverseMapRangeInterval( _
                ByRef yRange As TInterval) As TInterval
Const ProcName As String = "inverseMapRangeInterval"
Dim failpoint As String
On Error GoTo Err

If Vertical Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                            "Invalid for vertical line"
                 
inverseMapRangeInterval.startValue = InverseMapPoint(yRange.startValue)
inverseMapRangeInterval.endValue = InverseMapPoint(yRange.endValue)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function isExtended() As Boolean
isExtended = Extended Or FixedX
End Function

Private Function IsPropertySet( _
                ByVal flag As LinePropertyFlags) As Boolean
Const ProcName As String = "IsPropertySet"
Dim failpoint As String
On Error GoTo Err

IsPropertySet = gIsFlagSet(mPropertyOverrideFlags, flag)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function mapDomainInterval( _
                ByRef xDomain As TInterval) As TInterval
Const ProcName As String = "mapDomainInterval"
Dim failpoint As String
On Error GoTo Err

If Vertical Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                            "Invalid for vertical line"
                 
mapDomainInterval.startValue = mapPoint(xDomain.startValue)
mapDomainInterval.endValue = mapPoint(xDomain.endValue)
mapDomainInterval.isValid = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function mapPoint( _
                ByVal X As Double) As Double
Const ProcName As String = "mapPoint"
Dim failpoint As String
On Error GoTo Err

If Vertical Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                            "Invalid for vertical line"
                 
If X = MinusInfinityDouble Then
    If mA > 0 Then
        mapPoint = MinusInfinityDouble
    ElseIf mA < 0 Then
        mapPoint = PlusInfinityDouble
    Else
        mapPoint = mB
    End If
ElseIf X = PlusInfinityDouble Then
    If mA > 0 Then
        mapPoint = PlusInfinityDouble
    ElseIf mA < 0 Then
        mapPoint = MinusInfinityDouble
    Else
        mapPoint = mB
    End If
Else
    mapPoint = mA * X + mB
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub setDrawingAttributes()
Const ProcName As String = "setDrawingAttributes"
Dim failpoint As String
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setDirection()
Const ProcName As String = "setDirection"
Dim failpoint As String
On Error GoTo Err

If mX1 = mX2 Then
    If mY1 <= mY2 Then
        mVertical = VerticalUp
    Else
        mVertical = VerticalDown
    End If
Else
    mVertical = VerticalNot
End If
If mX1 <= mX2 Then
    If mY1 <= mY2 Then
        mDirection = NE
        mLeftToRight = True
        mBottomToTop = True
    Else
        mDirection = SE
        mLeftToRight = True
        mBottomToTop = False
    End If
Else
    If mY1 <= mY2 Then
        mDirection = NW
        mLeftToRight = False
        mBottomToTop = True
    Else
        mDirection = SW
        mLeftToRight = False
        mBottomToTop = False
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setPropertyFlag( _
                ByVal flag As LinePropertyFlags)
Const ProcName As String = "setPropertyFlag"
Dim failpoint As String
On Error GoTo Err

mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub undraw()

Const ProcName As String = "undraw"
Dim failpoint As String
On Error GoTo Err

If Not mVisible Then Exit Sub
mGOSP.ObjectUndrawn mHandle, mBoundingRectInViewport

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub


