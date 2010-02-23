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

Private mBoundingRect As TRectangle ' the rectangle that surrounds the box (if
                                    ' any - including the thickness of the lines) or
                                    ' the Text if no box
Private mTextTRect As TRectangle    ' the TRectangle containing the Text
Private mBoxTRect As TRectangle     ' the TRectangle used to draw the box

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
Private mAngle As Double
Private mJustification As TextJustifyModes
Private mMultiLine As Boolean
Private mEllipsis As EllipsisModes
Private mExpandTabs As Boolean
Private mTabWidth As Long
Private mWordWrap As Boolean
Private mLeftMargin As DimensionX
Private mRightMargin As DimensionX
Private mHideIfBlank As Boolean

Private mPosition As Point
Private mOffset As Dimension

Private mKey As String

Private mSize As Dimension

Private mPropertyOverrideFlags As TextPropertyOverrideFlags

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mBoundingRect.isValid = False
mTextTRect.isValid = False
End Sub

Private Sub Class_Terminate()
'gLogger.Log LogLevelHighDetail, "Text terminated"
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

Const ProcName As String = "IGraphicObject_boundingRectangle"
Dim failpoint As String
On Error GoTo Err

viewportChanges = mViewport.CompareTo(prevGaugeX, prevGaugeY, prevRect)

If (viewportChanges And ViewportComparisonCodes.GaugeChangedX) Or _
    (viewportChanges And ViewportComparisonCodes.GaugeChangedY) Or _
    (viewportChanges And ViewportComparisonCodes.BoundsChangedX) Or _
    (viewportChanges And ViewportComparisonCodes.BoundsChangedY) Or _
    mSeriesUpdateNumber <> mSeries.SeriesUpdateNumber _
Then
    mSeriesUpdateNumber = mSeries.SeriesUpdateNumber
    calcBoundingRect
End If

prevGaugeX = mViewport.GaugeX
prevGaugeY = mViewport.GaugeY
prevRect = mViewport.Boundary

IGraphicObject_boundingRectangle = mBoundingRect

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get IGraphicObject_BoundingRectangleInViewport() As TRectangle
Const ProcName As String = "IGraphicObject_BoundingRectangleInViewport"
Dim failpoint As String
On Error GoTo Err

IGraphicObject_BoundingRectangleInViewport = IGraphicObject_boundingRectangle
RectSetXInterval IGraphicObject_BoundingRectangleInViewport, _
                IntIntersection(RectGetXInterval(IGraphicObject_BoundingRectangleInViewport), _
                                RectGetXInterval(mViewport.Boundary))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get IGraphicObject_BoundingRectanglesAt(ByVal PeriodNumber As Long) As TRectangle()
'
End Property

Private Property Get IGraphicObject_capabilities() As GraphicObjectCapabilities
Const ProcName As String = "IGraphicObject_capabilities"
Dim failpoint As String
On Error GoTo Err

IGraphicObject_capabilities = GraphicObjectCapabilities.BlockUndraw

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "IGraphicObject_draw"
Dim failpoint As String
On Error GoTo Err

'mCanvas.SetClippingRegion mViewport.ConvertTRectangleToRect(areas(0))
Draw
'mCanvas.ClearClippingRegion

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = isExtended
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
Const ProcName As String = "IGraphicObject_HitTest"
Dim failpoint As String
On Error GoTo Err

IGraphicObject_HitTest = RectContainsPoint(mBoundingRect, X, Y)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
IGraphicObject_noDraw = (mPosition Is Nothing) Or (HideIfBlank And mText = "")
End Property

Private Property Get IGraphicObject_PeriodNumber() As Long
If isExtended Or Not mBoundingRect.isValid Then
    IGraphicObject_PeriodNumber = &H80000000
Else
    IGraphicObject_PeriodNumber = Int(mPosition.XLogical)
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

Private Property Get IGraphicObject_VerticalRangeInXinterval(pXInterval As TInterval) As TInterval
If IntIntersection(RectGetXInterval(IGraphicObject_boundingRectangle), _
                                pXInterval).isValid _
Then
    IGraphicObject_VerticalRangeInXinterval = RectGetYInterval(mBoundingRect)
End If
End Property

Private Property Let IGraphicObject_visible(ByVal value As Boolean)
mVisible = value
End Property

'================================================================================
' mLocalStyle Event Handlers
'================================================================================

Private Sub mLocalStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
Const ProcName As String = "mLocalStyle_PropertyChanged"
Dim failpoint As String
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case "ALIGN"
    initiateRedraw True, True
Case "ANGLE"
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
Case "ELLIPSIS"
    initiateRedraw False, False
Case "EXPANDTABS"
    initiateRedraw True, True
Case "EXTENDED"
    initiateRedraw True, True
Case "FIXEDX"
    initiateRedraw True, True
Case "FIXEDY"
    initiateRedraw True, True
Case "FONT"
    initiateRedraw True, True
Case "HIDEIFBLANK"
    initiateRedraw False, False
Case "INCLUDEINAUTOSCALE"
    initiateRedraw False, False
Case "JUSTIFICATION"
    initiateRedraw False, False
Case "LEFTMARGIN"
    initiateRedraw True, True
Case "MULTILINE"
    initiateRedraw True, True
Case "OFFSET"
    initiateRedraw True, True
Case "PADDINGX"
    initiateRedraw True, True
Case "PADDINGY"
    initiateRedraw True, True
Case "RIGHTMARGIN"
    initiateRedraw True, True
Case "SIZE"
    initiateRedraw True, True
Case "TABWIDTH"
    initiateRedraw True, True
Case "WORDWRAP"
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

Public Property Get Align() As TextAlignModes
Const ProcName As String = "Align"
Dim failpoint As String
On Error GoTo Err

Align = mSeries.Align
If Not mLocalStyle Is Nothing Then Align = mLocalStyle.Align
If isPropertyOverrideFlagSet(TextIsSetAlign) Then Align = mAlign

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Align(ByVal value As TextAlignModes)
Const ProcName As String = "Align"
Dim failpoint As String
On Error GoTo Err

If mAlign = value And isPropertyOverrideFlagSet(TextIsSetAlign) Then Exit Property
mAlign = value
setPropertyOverrideFlag TextIsSetAlign
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Angle() As Double
Const ProcName As String = "Angle"
Dim failpoint As String
On Error GoTo Err

Angle = mSeries.Angle
If Not mLocalStyle Is Nothing Then Angle = mLocalStyle.Angle
If isPropertyOverrideFlagSet(TextIsSetAngle) Then Angle = mAngle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Angle(ByVal value As Double)
Const ProcName As String = "Angle"
Dim failpoint As String
On Error GoTo Err

If mAngle = value And isPropertyOverrideFlagSet(TextIsSetAngle) Then Exit Property
mAngle = value
setPropertyOverrideFlag TextIsSetAngle
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Box() As Boolean
Const ProcName As String = "Box"
Dim failpoint As String
On Error GoTo Err

Box = mSeries.Box
If Not mLocalStyle Is Nothing Then Box = mLocalStyle.Box
If isPropertyOverrideFlagSet(TextIsSetBox) Then Box = mBox

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Box(ByVal value As Boolean)
Const ProcName As String = "Box"
Dim failpoint As String
On Error GoTo Err

If mBox = value And isPropertyOverrideFlagSet(TextIsSetBox) Then Exit Property
mBox = value
setPropertyOverrideFlag TextIsSetBox
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxColor() As Long
Const ProcName As String = "BoxColor"
Dim failpoint As String
On Error GoTo Err

BoxColor = mSeries.BoxColor
If Not mLocalStyle Is Nothing Then BoxColor = mLocalStyle.BoxColor
If isPropertyOverrideFlagSet(TextIsSetBoxColor) Then BoxColor = mBoxColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let BoxColor(ByVal value As Long)
Const ProcName As String = "BoxColor"
Dim failpoint As String
On Error GoTo Err

If mBoxColor = value And isPropertyOverrideFlagSet(TextIsSetBoxColor) Then Exit Property
mBoxColor = value
setPropertyOverrideFlag TextIsSetBoxColor
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxFillColor() As Long
Const ProcName As String = "BoxFillColor"
Dim failpoint As String
On Error GoTo Err

BoxFillColor = mSeries.BoxFillColor
If Not mLocalStyle Is Nothing Then BoxFillColor = mLocalStyle.BoxFillColor
If isPropertyOverrideFlagSet(TextIsSetBoxFillColor) Then BoxFillColor = mBoxFillColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let BoxFillColor(ByVal value As Long)
Const ProcName As String = "BoxFillColor"
Dim failpoint As String
On Error GoTo Err

If mBoxFillColor = value And isPropertyOverrideFlagSet(TextIsSetBoxFillColor) Then Exit Property
mBoxFillColor = value
setPropertyOverrideFlag TextIsSetBoxFillColor
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxFillStyle() As FillStyles
Const ProcName As String = "BoxFillStyle"
Dim failpoint As String
On Error GoTo Err

BoxFillStyle = mSeries.BoxFillStyle
If Not mLocalStyle Is Nothing Then BoxFillStyle = mLocalStyle.BoxFillStyle
If isPropertyOverrideFlagSet(TextIsSetBoxFillStyle) Then BoxFillStyle = mBoxFillStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let BoxFillStyle(ByVal value As FillStyles)
Const ProcName As String = "BoxFillStyle"
Dim failpoint As String
On Error GoTo Err

If mBoxFillStyle = value And isPropertyOverrideFlagSet(TextIsSetBoxFillStyle) Then Exit Property
mBoxFillStyle = value
setPropertyOverrideFlag TextIsSetBoxFillStyle
initiateRedraw True, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxFillWithBackgroundColor() As Boolean
Const ProcName As String = "BoxFillWithBackgroundColor"
Dim failpoint As String
On Error GoTo Err

BoxFillWithBackgroundColor = mSeries.BoxFillWithBackgroundColor
If Not mLocalStyle Is Nothing Then BoxFillWithBackgroundColor = mLocalStyle.BoxFillWithBackgroundColor
If isPropertyOverrideFlagSet(TextIsSetBoxFillWithBackGroundColor) Then BoxFillWithBackgroundColor = mBoxFillWithBackgroundColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let BoxFillWithBackgroundColor(ByVal value As Boolean)
Const ProcName As String = "BoxFillWithBackgroundColor"
Dim failpoint As String
On Error GoTo Err

If mBoxFillWithBackgroundColor = value And isPropertyOverrideFlagSet(TextIsSetBoxFillWithBackGroundColor) Then Exit Property
mBoxFillWithBackgroundColor = value
setPropertyOverrideFlag TextIsSetBoxFillWithBackGroundColor
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxStyle() As LineStyles
Const ProcName As String = "BoxStyle"
Dim failpoint As String
On Error GoTo Err

BoxStyle = mSeries.BoxStyle
If Not mLocalStyle Is Nothing Then BoxStyle = mLocalStyle.BoxStyle
If isPropertyOverrideFlagSet(TextIsSetBoxStyle) Then BoxStyle = mBoxStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let BoxStyle(ByVal value As LineStyles)
Const ProcName As String = "BoxStyle"
Dim failpoint As String
On Error GoTo Err

If mBoxStyle = value And isPropertyOverrideFlagSet(TextIsSetBoxStyle) Then Exit Property
mBoxStyle = value
setPropertyOverrideFlag TextIsSetBoxStyle
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxThickness() As Long
Const ProcName As String = "BoxThickness"
Dim failpoint As String
On Error GoTo Err

BoxThickness = mSeries.BoxThickness
If Not mLocalStyle Is Nothing Then BoxThickness = mLocalStyle.BoxThickness
If isPropertyOverrideFlagSet(TextIsSetBoxThickness) Then BoxThickness = mBoxThickness

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let BoxThickness(ByVal value As Long)
Const ProcName As String = "BoxThickness"
Dim failpoint As String
On Error GoTo Err

If mBoxThickness = value And isPropertyOverrideFlagSet(TextIsSetBoxThickness) Then Exit Property
mBoxThickness = value
setPropertyOverrideFlag TextIsSetBoxThickness
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
If Not mLocalStyle Is Nothing Then Color = mLocalStyle.Color
If isPropertyOverrideFlagSet(TextIsSetColor) Then Color = mColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Color(ByVal value As Long)
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

If mColor = value And isPropertyOverrideFlagSet(TextIsSetColor) Then Exit Property
mColor = value
setPropertyOverrideFlag TextIsSetColor
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Ellipsis() As EllipsisModes
Const ProcName As String = "Ellipsis"
Dim failpoint As String
On Error GoTo Err

Ellipsis = mSeries.Ellipsis
If Not mLocalStyle Is Nothing Then Ellipsis = mLocalStyle.Ellipsis
If isPropertyOverrideFlagSet(TextIsSetEllipsis) Then Ellipsis = mEllipsis

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Ellipsis(ByVal value As EllipsisModes)
Const ProcName As String = "Ellipsis"
Dim failpoint As String
On Error GoTo Err

If mEllipsis = value And isPropertyOverrideFlagSet(TextIsSetEllipsis) Then Exit Property
mEllipsis = value
setPropertyOverrideFlag TextIsSetEllipsis
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ExpandTabs() As Boolean
Const ProcName As String = "ExpandTabs"
Dim failpoint As String
On Error GoTo Err

ExpandTabs = mSeries.ExpandTabs
If Not mLocalStyle Is Nothing Then ExpandTabs = mLocalStyle.ExpandTabs
If isPropertyOverrideFlagSet(TextIsSetExpandTabs) Then ExpandTabs = mExpandTabs

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ExpandTabs(ByVal value As Boolean)
Const ProcName As String = "ExpandTabs"
Dim failpoint As String
On Error GoTo Err

If mExpandTabs = value And isPropertyOverrideFlagSet(TextIsSetExpandTabs) Then Exit Property
mExpandTabs = value
setPropertyOverrideFlag TextIsSetExpandTabs
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
If Not mLocalStyle Is Nothing Then Extended = mLocalStyle.Extended
If isPropertyOverrideFlagSet(TextIsSetExtended) Then Extended = mExtended

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Extended(ByVal value As Boolean)
Const ProcName As String = "Extended"
Dim failpoint As String
On Error GoTo Err

If mExtended = value And isPropertyOverrideFlagSet(TextIsSetExtended) Then Exit Property
mExtended = value
setPropertyOverrideFlag TextIsSetExtended
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
If Not mLocalStyle Is Nothing Then FixedX = mLocalStyle.FixedX
If isPropertyOverrideFlagSet(TextIsSetFixedX) Then FixedX = mFixedX

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let FixedX(ByVal value As Boolean)
Const ProcName As String = "FixedX"
Dim failpoint As String
On Error GoTo Err

If mFixedX = value And isPropertyOverrideFlagSet(TextIsSetFixedX) Then Exit Property
mFixedX = value
setPropertyOverrideFlag TextIsSetFixedX
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
If Not mLocalStyle Is Nothing Then FixedY = mLocalStyle.FixedY
If isPropertyOverrideFlagSet(TextIsSetFixedY) Then FixedY = mFixedY

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let FixedY(ByVal value As Boolean)
Const ProcName As String = "FixedY"
Dim failpoint As String
On Error GoTo Err

If mFixedY = value And isPropertyOverrideFlagSet(TextIsSetFixedY) Then Exit Property
mFixedY = value
setPropertyOverrideFlag TextIsSetFixedY
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Font() As StdFont
Const ProcName As String = "Font"
Dim failpoint As String
On Error GoTo Err

If isPropertyOverrideFlagSet(TextIsSetFont) Then
    Set Font = mFont
ElseIf Not mLocalStyle Is Nothing Then
    Set Font = mLocalStyle.Font
Else
    Set Font = mSeries.Font
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Font(ByVal value As StdFont)
Const ProcName As String = "Font"
Dim failpoint As String
On Error GoTo Err

Set mFont = value
If mFont Is Nothing Then
    clearPropertyOverrideFlag TextIsSetFont
Else
    setPropertyOverrideFlag TextIsSetFont
End If
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Get HideIfBlank() As Boolean
Const ProcName As String = "HideIfBlank"
Dim failpoint As String
On Error GoTo Err

HideIfBlank = mSeries.HideIfBlank
If isPropertyOverrideFlagSet(TextIsSetHideIfBlank) Then HideIfBlank = mHideIfBlank

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let HideIfBlank(ByVal value As Boolean)
Const ProcName As String = "HideIfBlank"
Dim failpoint As String
On Error GoTo Err

If mHideIfBlank = value And isPropertyOverrideFlagSet(TextIsSetHideIfBlank) Then Exit Property
mHideIfBlank = value
setPropertyOverrideFlag TextIsSetHideIfBlank
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

IncludeInAutoscale = mSeries.IncludeInAutoscale
If isPropertyOverrideFlagSet(TextIsSetIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

If mIncludeInAutoscale = value And isPropertyOverrideFlagSet(TextIsSetIncludeInAutoscale) Then Exit Property
mIncludeInAutoscale = value
setPropertyOverrideFlag TextIsSetIncludeInAutoscale
initiateRedraw False, False

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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

Public Property Get Justification() As TextJustifyModes
Const ProcName As String = "Justification"
Dim failpoint As String
On Error GoTo Err

Justification = mSeries.Justification
If Not mLocalStyle Is Nothing Then Justification = mLocalStyle.Justification
If isPropertyOverrideFlagSet(TextIsSetJustification) Then Justification = mJustification

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Justification(ByVal value As TextJustifyModes)
Const ProcName As String = "Justification"
Dim failpoint As String
On Error GoTo Err

If mJustification = value And isPropertyOverrideFlagSet(TextIsSetJustification) Then Exit Property
mJustification = value
setPropertyOverrideFlag TextIsSetJustification
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Let Layer(ByVal value As LayerNumbers)
Const ProcName As String = "Layer"
Dim failpoint As String
On Error GoTo Err

If mLayer = value And isPropertyOverrideFlagSet(TextIsSetLayer) Then Exit Property
mLayer = value
setPropertyOverrideFlag TextIsSetLayer
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
If isPropertyOverrideFlagSet(TextIsSetLayer) Then Layer = mLayer

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LeftMargin() As DimensionX
Const ProcName As String = "LeftMargin"
Dim failpoint As String
On Error GoTo Err

Set LeftMargin = mSeries.LeftMargin
If isPropertyOverrideFlagSet(TextIsSetLeftMargin) Then Set LeftMargin = mLeftMargin

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let LeftMargin(ByVal value As DimensionX)
Const ProcName As String = "LeftMargin"
Dim failpoint As String
On Error GoTo Err

If value Is Nothing And (Not isPropertyOverrideFlagSet(TextIsSetLeftMargin)) Then
    Exit Property
ElseIf value Is Nothing Then
    Set mLeftMargin = Nothing
    clearPropertyOverrideFlag TextIsSetLeftMargin
ElseIf isPropertyOverrideFlagSet(TextIsSetLeftMargin) And value.Equals(mOffset) Then
    Exit Property
Else
    Set mLeftMargin = value
    setPropertyOverrideFlag TextIsSetLeftMargin
End If
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LocalStyle() As TextStyle
Set LocalStyle = mLocalStyle
End Property

Public Property Let LocalStyle( _
                ByVal value As TextStyle)
Const ProcName As String = "LocalStyle"
Dim failpoint As String
On Error GoTo Err

Set mLocalStyle = value
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get MultiLine() As Boolean
Const ProcName As String = "MultiLine"
Dim failpoint As String
On Error GoTo Err

MultiLine = mSeries.MultiLine
If Not mLocalStyle Is Nothing Then MultiLine = mLocalStyle.MultiLine
If isPropertyOverrideFlagSet(TextIsSetMultiLine) Then MultiLine = mMultiLine

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let MultiLine(ByVal value As Boolean)
Const ProcName As String = "MultiLine"
Dim failpoint As String
On Error GoTo Err

If mMultiLine = value And isPropertyOverrideFlagSet(TextIsSetMultiLine) Then Exit Property
mMultiLine = value
setPropertyOverrideFlag TextIsSetMultiLine
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let PaddingX(ByVal value As Double)
Const ProcName As String = "PaddingX"
Dim failpoint As String
On Error GoTo Err

If mPaddingX = value And isPropertyOverrideFlagSet(TextIsSetPaddingX) Then Exit Property
mPaddingX = value
setPropertyOverrideFlag TextIsSetPaddingX
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PaddingX() As Double
Const ProcName As String = "PaddingX"
Dim failpoint As String
On Error GoTo Err

PaddingX = mSeries.PaddingX
If Not mLocalStyle Is Nothing Then PaddingX = mLocalStyle.PaddingX
If isPropertyOverrideFlagSet(TextIsSetPaddingX) Then PaddingX = mPaddingX

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let PaddingY(ByVal value As Double)
Const ProcName As String = "PaddingY"
Dim failpoint As String
On Error GoTo Err

If mPaddingY = value And isPropertyOverrideFlagSet(TextIsSetPaddingY) Then Exit Property
mPaddingY = value
setPropertyOverrideFlag TextIsSetPaddingY
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PaddingY() As Double
Const ProcName As String = "PaddingY"
Dim failpoint As String
On Error GoTo Err

PaddingY = mSeries.PaddingY
If Not mLocalStyle Is Nothing Then PaddingY = mLocalStyle.PaddingY
If isPropertyOverrideFlagSet(TextIsSetPaddingY) Then PaddingY = mPaddingY

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Position() As Point
Set Position = mPosition
End Property

Public Property Let Position(ByVal value As Point)
Const ProcName As String = "Position"
Dim failpoint As String
On Error GoTo Err

If value Is Nothing Then
    Set mPosition = Nothing
    mBoundingRect.isValid = False
'    mBlank = True
    initiateRedraw True, True
    Exit Property
ElseIf Not mPosition Is Nothing Then
    If value.Equals(mPosition) Then
        If mPosition.CoordinateSystemX = CoordsLogical And mPosition.CoordinateSystemY = CoordsLogical Then Exit Property
    End If
End If

'If mText <> "" Then mBlank = False
Set mPosition = value
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Offset() As Dimension
Const ProcName As String = "Offset"
Dim failpoint As String
On Error GoTo Err

Set Offset = mSeries.Offset
If isPropertyOverrideFlagSet(TextIsSetOffset) Then Set Offset = mOffset

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Offset(ByVal value As Dimension)
Const ProcName As String = "Offset"
Dim failpoint As String
On Error GoTo Err

If value Is Nothing And (Not isPropertyOverrideFlagSet(TextIsSetOffset)) Then
    Exit Property
ElseIf value Is Nothing Then
    Set mOffset = Nothing
    clearPropertyOverrideFlag TextIsSetOffset
ElseIf isPropertyOverrideFlagSet(TextIsSetOffset) And value.Equals(mOffset) Then
    Exit Property
Else
    Set mOffset = value
    setPropertyOverrideFlag TextIsSetOffset
End If
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get RightMargin() As DimensionX
Const ProcName As String = "RightMargin"
Dim failpoint As String
On Error GoTo Err

Set RightMargin = mSeries.RightMargin
If isPropertyOverrideFlagSet(TextIsSetRightMargin) Then Set RightMargin = mRightMargin

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let RightMargin(ByVal value As DimensionX)
Const ProcName As String = "RightMargin"
Dim failpoint As String
On Error GoTo Err

If value Is Nothing And (Not isPropertyOverrideFlagSet(TextIsSetRightMargin)) Then
    Exit Property
ElseIf value Is Nothing Then
    Set mRightMargin = Nothing
    clearPropertyOverrideFlag TextIsSetRightMargin
ElseIf isPropertyOverrideFlagSet(TextIsSetRightMargin) And value.Equals(mOffset) Then
    Exit Property
Else
    Set mRightMargin = value
    setPropertyOverrideFlag TextIsSetRightMargin
End If
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Size(ByVal value As Dimension)
Const ProcName As String = "Size"
Dim failpoint As String
On Error GoTo Err

If value Is Nothing And (Not isPropertyOverrideFlagSet(TextIsSetSize)) Then
    Exit Property
ElseIf value Is Nothing Then
    Set mSize = Nothing
    clearPropertyOverrideFlag TextIsSetSize
ElseIf isPropertyOverrideFlagSet(TextIsSetSize) And value.Equals(mOffset) Then
    Exit Property
Else
    Set mSize = value
    setPropertyOverrideFlag TextIsSetSize
End If
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Size() As Dimension
Const ProcName As String = "Size"
Dim failpoint As String
On Error GoTo Err

Set Size = mSeries.Size
If isPropertyOverrideFlagSet(TextIsSetSize) Then Set Size = mSize

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Text(value As String)
Const ProcName As String = "Text"
Dim failpoint As String
On Error GoTo Err

If value = mText Then Exit Property
mText = value
'mBlank = (value = "" Or mPosition Is Nothing)
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TabWidth() As Long
Const ProcName As String = "TabWidth"
Dim failpoint As String
On Error GoTo Err

TabWidth = mSeries.TabWidth
If Not mLocalStyle Is Nothing Then TabWidth = mLocalStyle.TabWidth
If isPropertyOverrideFlagSet(TextIsSetTabWidth) Then TabWidth = mTabWidth

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let TabWidth(ByVal value As Long)
Const ProcName As String = "TabWidth"
Dim failpoint As String
On Error GoTo Err

If mTabWidth = value And isPropertyOverrideFlagSet(TextIsSetTabWidth) Then Exit Property
mTabWidth = value
setPropertyOverrideFlag TextIsSetTabWidth
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Text() As String
Text = mText
End Property

Public Property Get Visible() As Boolean
Visible = mVisible
End Property

Public Property Get WordWrap() As Boolean
Const ProcName As String = "WordWrap"
Dim failpoint As String
On Error GoTo Err

WordWrap = mSeries.WordWrap
If Not mLocalStyle Is Nothing Then WordWrap = mLocalStyle.WordWrap
If isPropertyOverrideFlagSet(TextIsSetWordWrap) Then WordWrap = mWordWrap

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let WordWrap(ByVal value As Boolean)
Const ProcName As String = "WordWrap"
Dim failpoint As String
On Error GoTo Err

If mWordWrap = value And isPropertyOverrideFlagSet(TextIsSetWordWrap) Then Exit Property
mWordWrap = value
setPropertyOverrideFlag TextIsSetWordWrap
initiateRedraw True, True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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

Friend Sub Draw()
Const ProcName As String = "Draw"
Dim failpoint As String
On Error GoTo Err

If Not canBeDrawn Then Exit Sub

If (Angle - 360# * Int(Angle / 360)) <> 0 Then
    Dim alignPoint As TPoint
    alignPoint = getAlignmentPoint(mTextTRect, mBoxTRect)
    mViewport.RotateAboutPoint Angle, alignPoint
End If

If Box Then drawBox
DrawText

If (Angle - 360# * Int(Angle / 360)) <> 0 Then
    mViewport.Reset
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Initialise( _
                ByVal pSeries As TextSeries, _
                ByVal pKey As String, _
                ByVal pGOSP As GraphObjServiceProvider, _
                ByVal pViewport As Viewport, _
                ByVal pLocalStyle As TextStyle)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mSeries = pSeries
mKey = pKey
Set mGOSP = pGOSP
Set mViewport = pViewport
Set mCanvas = mViewport.Canvas
Set mLocalStyle = pLocalStyle

calcBoundingRect

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub adjustRectSize( _
                ByRef pRect As RECT, _
                ByRef pSizeAdjustment As W32Point)
pRect.Right = pRect.Right + pSizeAdjustment.X
If pRect.Right < pRect.Left Then pRect.Right = pRect.Left

pRect.Bottom = pRect.Bottom + pSizeAdjustment.Y
If pRect.Bottom = pRect.Top Then pRect.Bottom = pRect.Top
End Sub

Private Sub calcBoundingRect()
Dim alignmentOffset As W32Point
Dim lTextDim As Dimension
Dim lTextRect As RECT
Dim lBoxRect As RECT
Dim lBoundingRect As RECT
Dim sizeAdjustment As W32Point

Const ProcName As String = "calcBoundingRect"
Dim failpoint As String
On Error GoTo Err

If mPosition Is Nothing Then
    mBoundingRect.isValid = False
    Exit Sub
End If

Set mPosition = convertPosition(mPosition)

mCanvas.SetTextAttributes Color, Font, True

Set lTextDim = mCanvas.GetTextDimensions(mText, _
                                    Size, _
                                    Justification, _
                                    MultiLine, _
                                    Ellipsis, _
                                    ExpandTabs, _
                                    TabWidth, _
                                    WordWrap, _
                                    LeftMargin, _
                                    RightMargin)
With mTextTRect
    .Left = mPosition.XLogical
    .Right = .Left + lTextDim.XLogical
    .Top = mPosition.yLogical
    .Bottom = .Top - lTextDim.yLogical
    .isValid = True
End With
    
If Not mOffset Is Nothing Then
    RectOffsetDim mTextTRect, mOffset
End If
    
lTextRect = mViewport.ConvertTRectangleToRect(mTextTRect)

If Box Then
    
    lBoxRect = lTextRect
    
    twwin32api.InflateRect lBoxRect, _
                            mViewport.ConvertDistanceToPixelsX(PaddingX / 10), _
                            mViewport.ConvertDistanceToPixelsY(PaddingY / 10)
                            
    If PaddingY = 0 Then
        ' add one pixel at the top for fonts that have no leading (such as Lucida Console)
        lBoxRect.Top = lBoxRect.Top - 1
    End If
    
    Dim lBoxThickness As Long
    lBoxThickness = BoxThickness
    If lBoxThickness = 0 Then
        lBoundingRect = lBoxRect
    ElseIf (lBoxThickness > 1 And BoxStyle = LineInsideSolid) Or lBoxThickness = 1 Then
        twwin32api.InflateRect lBoxRect, lBoxThickness, lBoxThickness
        lBoundingRect = lBoxRect
        twwin32api.InflateRect lBoundingRect, 1, 1
    Else
        twwin32api.InflateRect lBoxRect, _
                                lBoxThickness - Int(lBoxThickness / 2), _
                                lBoxThickness - Int(lBoxThickness / 2)
        lBoundingRect = lBoxRect
        twwin32api.InflateRect lBoundingRect, Int(lBoxThickness / 2) + 1, Int(lBoxThickness / 2) + 1
    End If
Else
    lBoxRect = lTextRect
    lBoundingRect = lTextRect
End If

'If Not Size Is Nothing Then
'    sizeAdjustment = getSizeAdjustment(lBoundingRect)
'
'    adjustRectSize lTextRect, sizeAdjustment
'    adjustRectSize lBoxRect, sizeAdjustment
'    adjustRectSize lBoundingRect, sizeAdjustment
'End If

alignmentOffset = getAlignmentOffset(lTextRect, lBoxRect)

twwin32api.OffsetRect lTextRect, alignmentOffset.X, alignmentOffset.Y
twwin32api.OffsetRect lBoxRect, alignmentOffset.X, alignmentOffset.Y
twwin32api.OffsetRect lBoundingRect, alignmentOffset.X, alignmentOffset.Y

lBoundingRect.Right = lBoundingRect.Right + 1
lBoundingRect.Bottom = lBoundingRect.Bottom + 1

If Angle <> 0 Then
    Dim alignPoint As W32Point
    alignPoint = getAlignmentPointW32(lTextRect, lBoxRect)
    W32RectAdjustForRotationAboutPoint lBoundingRect, _
                                    Angle, _
                                    alignPoint
End If

mTextTRect = mViewport.ConvertRectToTRectangle(lTextRect)
mBoxTRect = mViewport.ConvertRectToTRectangle(lBoxRect)
mBoundingRect = mViewport.ConvertRectToTRectangle(lBoundingRect)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Function canBeDrawn() As Boolean
Const ProcName As String = "canBeDrawn"
Dim failpoint As String
On Error GoTo Err

'If mBlank Then Exit Function
If mPosition Is Nothing Then Exit Function
If Not mVisible Then Exit Function
canBeDrawn = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As TextPropertyOverrideFlags)
Const ProcName As String = "clearPropertyOverrideFlag"
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

Set convertPosition = pPoint.ChangeCoordinateSystem(IIf(FixedX, CoordinateSystems.CoordsRelative, CoordinateSystems.CoordsLogical), _
                                                    IIf(FixedY, CoordinateSystems.CoordsRelative, CoordinateSystems.CoordsLogical))

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub drawBox()
Const ProcName As String = "drawBox"
Dim failpoint As String
On Error GoTo Err

If BoxFillWithBackgroundColor Then
    drawBoxWithBackgroundFillColor
Else
    drawBoxWithSpecifiedFillStyle BoxFillStyle
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub drawBoxWithBackgroundFillColor()
Const ProcName As String = "drawBoxWithBackgroundFillColor"
Dim failpoint As String
On Error GoTo Err

mCanvas.SaveClippingRegion
mCanvas.IntersectClippingRegion mViewport.ConvertTRectangleToRect(mBoxTRect)

mViewport.PaintBackground
drawBoxWithSpecifiedFillStyle FillTransparent

mCanvas.RestoreClippingRegion
Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub drawBoxWithSpecifiedFillStyle( _
                ByVal pFillStyle As FillStyles)
Const ProcName As String = "drawBoxWithSpecifiedFillStyle"
Dim failpoint As String
On Error GoTo Err

With mViewport
    .Canvas.SetPenAttributes BoxColor, _
                    BoxThickness, _
                    BoxStyle, _
                    DrawModes.DrawModeCopyPen
    .Canvas.SetBrushAttributes BoxFillColor, pFillStyle
    .Canvas.DrawRectangleFromTRectangle mBoxTRect
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub DrawText()
Const ProcName As String = "DrawText"
Dim failpoint As String
On Error GoTo Err

'mCanvas.SaveClippingRegion
'mCanvas.IntersectClippingRegion mViewport.ConvertTRectangleToRect(mTextTRect)

mCanvas.SetTextAttributes Color, Font, True
mCanvas.DrawText mText, _
                RectTopLeft(mTextTRect), _
                mTextTRect, _
                Justification, _
                MultiLine, _
                Ellipsis, _
                ExpandTabs, _
                TabWidth, _
                WordWrap, _
                LeftMargin, _
                RightMargin

'mCanvas.RestoreClippingRegion

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function getAlignmentOffset( _
                ByRef pTextBoundary As RECT, _
                ByRef pBoxBoundary As RECT) As W32Point

Const ProcName As String = "getAlignmentOffset"
Dim failpoint As String
On Error GoTo Err

getAlignmentOffset = W32PointSubtract(W32RectTopLeft(pTextBoundary), getAlignmentPointW32(pTextBoundary, pBoxBoundary))

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getAlignmentPoint( _
                ByRef pTextBoundary As TRectangle, _
                ByRef pBoxBoundary As TRectangle) As TPoint
Const ProcName As String = "getAlignmentPoint"
Dim failpoint As String
On Error GoTo Err

Select Case Align
Case AlignTopLeft
    getAlignmentPoint = RectTopLeft(pTextBoundary)
Case AlignCentreLeft
    getAlignmentPoint = RectCentreLeft(pTextBoundary)
Case AlignBottomLeft
    getAlignmentPoint = RectBottomLeft(pTextBoundary)
Case AlignTopCentre
    getAlignmentPoint = RectTopCentre(pTextBoundary)
Case AlignCentreCentre
    getAlignmentPoint = RectCentreCentre(pTextBoundary)
Case AlignBottomCentre
    getAlignmentPoint = RectBottomCentre(pTextBoundary)
Case AlignTopRight
    getAlignmentPoint = RectTopRight(pTextBoundary)
Case AlignCentreRight
    getAlignmentPoint = RectCentreRight(pTextBoundary)
Case AlignBottomRight
    getAlignmentPoint = RectBottomRight(pTextBoundary)
Case AlignBoxTopLeft
    getAlignmentPoint = RectTopLeft(pBoxBoundary)
Case AlignBoxCentreLeft
    getAlignmentPoint = RectCentreLeft(pBoxBoundary)
Case AlignBoxBottomLeft
    getAlignmentPoint = RectBottomLeft(pBoxBoundary)
Case AlignBoxTopCentre
    getAlignmentPoint = RectTopCentre(pBoxBoundary)
Case AlignBoxCentreCentre
    getAlignmentPoint = RectCentreCentre(pBoxBoundary)
Case AlignBoxBottomCentre
    getAlignmentPoint = RectBottomCentre(pBoxBoundary)
Case AlignBoxTopRight
    getAlignmentPoint = RectTopRight(pBoxBoundary)
Case AlignBoxCentreRight
    getAlignmentPoint = RectCentreRight(pBoxBoundary)
Case AlignBoxBottomRight
    getAlignmentPoint = RectBottomRight(pBoxBoundary)
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getAlignmentPointW32( _
                ByRef pTextBoundary As RECT, _
                ByRef pBoxBoundary As RECT) As W32Point
Const ProcName As String = "getAlignmentPointW32"
Dim failpoint As String
On Error GoTo Err

Select Case Align
Case AlignTopLeft
    getAlignmentPointW32 = W32RectTopLeft(pTextBoundary)
Case AlignCentreLeft
    getAlignmentPointW32 = W32RectCentreLeft(pTextBoundary)
Case AlignBottomLeft
    getAlignmentPointW32 = W32RectBottomLeft(pTextBoundary)
Case AlignTopCentre
    getAlignmentPointW32 = W32RectTopCentre(pTextBoundary)
Case AlignCentreCentre
    getAlignmentPointW32 = W32RectCentreCentre(pTextBoundary)
Case AlignBottomCentre
    getAlignmentPointW32 = W32RectBottomCentre(pTextBoundary)
Case AlignTopRight
    getAlignmentPointW32 = W32RectTopRight(pTextBoundary)
Case AlignCentreRight
    getAlignmentPointW32 = W32RectCentreRight(pTextBoundary)
Case AlignBottomRight
    getAlignmentPointW32 = W32RectBottomRight(pTextBoundary)
Case AlignBoxTopLeft
    getAlignmentPointW32 = W32RectTopLeft(pBoxBoundary)
Case AlignBoxCentreLeft
    getAlignmentPointW32 = W32RectCentreLeft(pBoxBoundary)
Case AlignBoxBottomLeft
    getAlignmentPointW32 = W32RectBottomLeft(pBoxBoundary)
Case AlignBoxTopCentre
    getAlignmentPointW32 = W32RectTopCentre(pBoxBoundary)
Case AlignBoxCentreCentre
    getAlignmentPointW32 = W32RectCentreCentre(pBoxBoundary)
Case AlignBoxBottomCentre
    getAlignmentPointW32 = W32RectBottomCentre(pBoxBoundary)
Case AlignBoxTopRight
    getAlignmentPointW32 = W32RectTopRight(pBoxBoundary)
Case AlignBoxCentreRight
    getAlignmentPointW32 = W32RectCentreRight(pBoxBoundary)
Case AlignBoxBottomRight
    getAlignmentPointW32 = W32RectBottomRight(pBoxBoundary)
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getSizeAdjustment( _
                ByRef lBoundingRect As RECT) As W32Point
If Me.Size.X <> 0 Then
    getSizeAdjustment.X = mViewport.ConvertLogicalToPixelsX(Me.Size.XLogical) - _
                            (lBoundingRect.Right - lBoundingRect.Left)
End If
If Me.Size.Y <> 0 Then
    getSizeAdjustment.Y = mViewport.ConvertLogicalToPixelsX(Me.Size.yLogical) - _
                            (lBoundingRect.Bottom - lBoundingRect.Top)
End If
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
If recalcBoundaries Then
    calcBoundingRect
End If
mGOSP.ObjectChanged mHandle, mBoundingRect

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function isExtended() As Boolean
isExtended = Extended Or FixedX
End Function

Private Function isPropertyOverrideFlagSet( _
                ByVal flag As TextPropertyOverrideFlags) As Boolean
Const ProcName As String = "isPropertyOverrideFlagSet"
Dim failpoint As String
On Error GoTo Err

isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub setPropertyOverrideFlag( _
                ByVal flag As TextPropertyOverrideFlags)
Const ProcName As String = "setPropertyOverrideFlag"
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

If canBeDrawn Then mGOSP.ObjectUndrawn mHandle, mBoundingRect

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

