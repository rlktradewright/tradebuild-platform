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
Private mLeftMargin As Dimension
Private mRightMargin As Dimension
Private mHideIfBlank As Boolean

Private mPosition As Point
Private mOffset As Size

Private mKey As String

Private mSize As Size

Private mPropertyOverrideFlags As TextPropertyFlags

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IGraphicObject_BoundingRectangleInViewport() As TRectangle
Const ProcName As String = "IGraphicObject_BoundingRectangleInViewport"

On Error GoTo Err

IGraphicObject_BoundingRectangleInViewport = IGraphicObject_boundingRectangle
RectSetXInterval IGraphicObject_BoundingRectangleInViewport, _
                IntIntersection(RectGetXInterval(IGraphicObject_BoundingRectangleInViewport), _
                                RectGetXInterval(mViewport.Boundary))

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IGraphicObject_BoundingRectanglesAt(ByVal PeriodNumber As Long) As TRectangle()
'
End Property

Private Property Get IGraphicObject_capabilities() As GraphicObjectCapabilities
Const ProcName As String = "IGraphicObject_capabilities"

On Error GoTo Err

IGraphicObject_capabilities = GraphicObjectCapabilities.BlockUndraw

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

On Error GoTo Err

'mCanvas.SetClippingRegion mViewport.ConvertTRectangleToRect(areas(0))
Draw
'mCanvas.ClearClippingRegion

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Private Property Let IGraphicObject_Handle(ByVal Value As Long)
mHandle = Value
End Property

Private Property Get IGraphicObject_Handle() As Long
IGraphicObject_Handle = mHandle
End Property

Private Function IGraphicObject_HitTest(ByVal X As Double, ByVal Y As Double) As Boolean
Const ProcName As String = "IGraphicObject_HitTest"

On Error GoTo Err

IGraphicObject_HitTest = RectContainsPoint(mBoundingRect, X, Y)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
IGraphicObject_noDraw = (mPosition Is Nothing) Or (HideIfBlank And mText = "")
End Property

Private Property Get IGraphicObject_PeriodNumber() As Long
If isExtended Or Not mBoundingRect.isValid Then
    IGraphicObject_PeriodNumber = &H80000000
Else
    IGraphicObject_PeriodNumber = Int(mPosition.XLogical(mViewport.Region))
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

Private Property Let IGraphicObject_visible(ByVal Value As Boolean)
mVisible = Value
End Property

'================================================================================
' mLocalStyle Event Handlers
'================================================================================

Private Sub mLocalStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mLocalStyle_PropertyChanged"

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
Case "LAYER"
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
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get Align() As TextAlignModes
Const ProcName As String = "Align"

On Error GoTo Err

Align = mSeries.Align
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyAlign) Then Align = mLocalStyle.Align
If IsPropertySet(TextPropertyAlign) Then Align = mAlign

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Align(ByVal Value As TextAlignModes)
Const ProcName As String = "Align"

On Error GoTo Err

If mAlign = Value And IsPropertySet(TextPropertyAlign) Then Exit Property
mAlign = Value
setPropertyFlag TextPropertyAlign
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Angle() As Double
Const ProcName As String = "Angle"

On Error GoTo Err

Angle = mSeries.Angle
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyAngle) Then Angle = mLocalStyle.Angle
If IsPropertySet(TextPropertyAngle) Then Angle = mAngle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Angle(ByVal Value As Double)
Const ProcName As String = "Angle"

On Error GoTo Err

If mAngle = Value And IsPropertySet(TextPropertyAngle) Then Exit Property
mAngle = Value
setPropertyFlag TextPropertyAngle
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Box() As Boolean
Const ProcName As String = "Box"

On Error GoTo Err

Box = mSeries.Box
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyBox) Then Box = mLocalStyle.Box
If IsPropertySet(TextPropertyBox) Then Box = mBox

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Box(ByVal Value As Boolean)
Const ProcName As String = "Box"

On Error GoTo Err

If mBox = Value And IsPropertySet(TextPropertyBox) Then Exit Property
mBox = Value
setPropertyFlag TextPropertyBox
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BoxColor() As Long
Const ProcName As String = "BoxColor"

On Error GoTo Err

BoxColor = mSeries.BoxColor
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyBoxColor) Then BoxColor = mLocalStyle.BoxColor
If IsPropertySet(TextPropertyBoxColor) Then BoxColor = mBoxColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BoxColor(ByVal Value As Long)
Const ProcName As String = "BoxColor"

On Error GoTo Err

If mBoxColor = Value And IsPropertySet(TextPropertyBoxColor) Then Exit Property
mBoxColor = Value
setPropertyFlag TextPropertyBoxColor
initiateRedraw False, False

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BoxFillColor() As Long
Const ProcName As String = "BoxFillColor"

On Error GoTo Err

BoxFillColor = mSeries.BoxFillColor
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyBoxFillColor) Then BoxFillColor = mLocalStyle.BoxFillColor
If IsPropertySet(TextPropertyBoxFillColor) Then BoxFillColor = mBoxFillColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BoxFillColor(ByVal Value As Long)
Const ProcName As String = "BoxFillColor"

On Error GoTo Err

If mBoxFillColor = Value And IsPropertySet(TextPropertyBoxFillColor) Then Exit Property
mBoxFillColor = Value
setPropertyFlag TextPropertyBoxFillColor
initiateRedraw False, False

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BoxFillStyle() As FillStyles
Const ProcName As String = "BoxFillStyle"

On Error GoTo Err

BoxFillStyle = mSeries.BoxFillStyle
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyBoxFillStyle) Then BoxFillStyle = mLocalStyle.BoxFillStyle
If IsPropertySet(TextPropertyBoxFillStyle) Then BoxFillStyle = mBoxFillStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BoxFillStyle(ByVal Value As FillStyles)
Const ProcName As String = "BoxFillStyle"

On Error GoTo Err

If mBoxFillStyle = Value And IsPropertySet(TextPropertyBoxFillStyle) Then Exit Property
mBoxFillStyle = Value
setPropertyFlag TextPropertyBoxFillStyle
initiateRedraw True, False

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BoxFillWithBackgroundColor() As Boolean
Const ProcName As String = "BoxFillWithBackgroundColor"

On Error GoTo Err

BoxFillWithBackgroundColor = mSeries.BoxFillWithBackgroundColor
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyBoxFillWithBackgroundColor) Then BoxFillWithBackgroundColor = mLocalStyle.BoxFillWithBackgroundColor
If IsPropertySet(TextPropertyBoxFillWithBackgroundColor) Then BoxFillWithBackgroundColor = mBoxFillWithBackgroundColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BoxFillWithBackgroundColor(ByVal Value As Boolean)
Const ProcName As String = "BoxFillWithBackgroundColor"

On Error GoTo Err

If mBoxFillWithBackgroundColor = Value And IsPropertySet(TextPropertyBoxFillWithBackgroundColor) Then Exit Property
mBoxFillWithBackgroundColor = Value
setPropertyFlag TextPropertyBoxFillWithBackgroundColor
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BoxStyle() As LineStyles
Const ProcName As String = "BoxStyle"

On Error GoTo Err

BoxStyle = mSeries.BoxStyle
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyBoxStyle) Then BoxStyle = mLocalStyle.BoxStyle
If IsPropertySet(TextPropertyBoxStyle) Then BoxStyle = mBoxStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BoxStyle(ByVal Value As LineStyles)
Const ProcName As String = "BoxStyle"

On Error GoTo Err

If mBoxStyle = Value And IsPropertySet(TextPropertyBoxStyle) Then Exit Property
mBoxStyle = Value
setPropertyFlag TextPropertyBoxStyle
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BoxThickness() As Long
Const ProcName As String = "BoxThickness"

On Error GoTo Err

BoxThickness = mSeries.BoxThickness
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyBoxThickness) Then BoxThickness = mLocalStyle.BoxThickness
If IsPropertySet(TextPropertyBoxThickness) Then BoxThickness = mBoxThickness

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BoxThickness(ByVal Value As Long)
Const ProcName As String = "BoxThickness"

On Error GoTo Err

If mBoxThickness = Value And IsPropertySet(TextPropertyBoxThickness) Then Exit Property
mBoxThickness = Value
setPropertyFlag TextPropertyBoxThickness
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Color() As Long
Const ProcName As String = "Color"

On Error GoTo Err

Color = mSeries.Color
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyColor) Then Color = mLocalStyle.Color
If IsPropertySet(TextPropertyColor) Then Color = mColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Color(ByVal Value As Long)
Const ProcName As String = "Color"

On Error GoTo Err

If mColor = Value And IsPropertySet(TextPropertyColor) Then Exit Property
mColor = Value
setPropertyFlag TextPropertyColor
initiateRedraw False, False

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Ellipsis() As EllipsisModes
Const ProcName As String = "Ellipsis"

On Error GoTo Err

Ellipsis = mSeries.Ellipsis
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyEllipsis) Then Ellipsis = mLocalStyle.Ellipsis
If IsPropertySet(TextPropertyEllipsis) Then Ellipsis = mEllipsis

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Ellipsis(ByVal Value As EllipsisModes)
Const ProcName As String = "Ellipsis"

On Error GoTo Err

If mEllipsis = Value And IsPropertySet(TextPropertyEllipsis) Then Exit Property
mEllipsis = Value
setPropertyFlag TextPropertyEllipsis
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ExpandTabs() As Boolean
Const ProcName As String = "ExpandTabs"

On Error GoTo Err

ExpandTabs = mSeries.ExpandTabs
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyExpandTabs) Then ExpandTabs = mLocalStyle.ExpandTabs
If IsPropertySet(TextPropertyExpandTabs) Then ExpandTabs = mExpandTabs

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ExpandTabs(ByVal Value As Boolean)
Const ProcName As String = "ExpandTabs"

On Error GoTo Err

If mExpandTabs = Value And IsPropertySet(TextPropertyExpandTabs) Then Exit Property
mExpandTabs = Value
setPropertyFlag TextPropertyExpandTabs
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Extended() As Boolean
Const ProcName As String = "Extended"

On Error GoTo Err

Extended = mSeries.Extended
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyExtended) Then Extended = mLocalStyle.Extended
If IsPropertySet(TextPropertyExtended) Then Extended = mExtended

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Extended(ByVal Value As Boolean)
Const ProcName As String = "Extended"

On Error GoTo Err

If mExtended = Value And IsPropertySet(TextPropertyExtended) Then Exit Property
mExtended = Value
setPropertyFlag TextPropertyExtended
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get FixedX() As Boolean
Const ProcName As String = "FixedX"

On Error GoTo Err

FixedX = mSeries.FixedX
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyFixedX) Then FixedX = mLocalStyle.FixedX
If IsPropertySet(TextPropertyFixedX) Then FixedX = mFixedX

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let FixedX(ByVal Value As Boolean)
Const ProcName As String = "FixedX"

On Error GoTo Err

If mFixedX = Value And IsPropertySet(TextPropertyFixedX) Then Exit Property
mFixedX = Value
setPropertyFlag TextPropertyFixedX
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get FixedY() As Boolean
Const ProcName As String = "FixedY"

On Error GoTo Err

FixedY = mSeries.FixedY
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyFixedY) Then FixedY = mLocalStyle.FixedY
If IsPropertySet(TextPropertyFixedY) Then FixedY = mFixedY

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let FixedY(ByVal Value As Boolean)
Const ProcName As String = "FixedY"

On Error GoTo Err

If mFixedY = Value And IsPropertySet(TextPropertyFixedY) Then Exit Property
mFixedY = Value
setPropertyFlag TextPropertyFixedY
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Font() As StdFont
Const ProcName As String = "Font"

On Error GoTo Err

Set Font = mSeries.Font
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyFont) Then Set Font = mLocalStyle.Font
If IsPropertySet(TextPropertyFont) Then Set Font = mFont

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Font(ByVal Value As StdFont)
Const ProcName As String = "Font"

On Error GoTo Err

Set mFont = Value
If mFont Is Nothing Then
    ClearProperty TextPropertyFont
Else
    setPropertyFlag TextPropertyFont
End If
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Public Property Get HideIfBlank() As Boolean
Const ProcName As String = "HideIfBlank"

On Error GoTo Err

HideIfBlank = mSeries.HideIfBlank
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyHideIfBlank) Then HideIfBlank = mLocalStyle.HideIfBlank
If IsPropertySet(TextPropertyHideIfBlank) Then HideIfBlank = mHideIfBlank

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HideIfBlank(ByVal Value As Boolean)
Const ProcName As String = "HideIfBlank"

On Error GoTo Err

If mHideIfBlank = Value And IsPropertySet(TextPropertyHideIfBlank) Then Exit Property
mHideIfBlank = Value
setPropertyFlag TextPropertyHideIfBlank
initiateRedraw False, False

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"

On Error GoTo Err

IncludeInAutoscale = mSeries.IncludeInAutoscale
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyIncludeInAutoscale) Then IncludeInAutoscale = mLocalStyle.IncludeInAutoscale
If IsPropertySet(TextPropertyIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"

On Error GoTo Err

If mIncludeInAutoscale = Value And IsPropertySet(TextPropertyIncludeInAutoscale) Then Exit Property
mIncludeInAutoscale = Value
setPropertyFlag TextPropertyIncludeInAutoscale
initiateRedraw False, False

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Public Property Get Justification() As TextJustifyModes
Const ProcName As String = "Justification"

On Error GoTo Err

Justification = mSeries.Justification
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyJustification) Then Justification = mLocalStyle.Justification
If IsPropertySet(TextPropertyJustification) Then Justification = mJustification

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Justification(ByVal Value As TextJustifyModes)
Const ProcName As String = "Justification"

On Error GoTo Err

If mJustification = Value And IsPropertySet(TextPropertyJustification) Then Exit Property
mJustification = Value
setPropertyFlag TextPropertyJustification
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"

On Error GoTo Err

If mLayer = Value And IsPropertySet(TextPropertyLayer) Then Exit Property
mLayer = Value
setPropertyFlag TextPropertyLayer
initiateRedraw False, False

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"

On Error GoTo Err

Layer = mSeries.Layer
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyLayer) Then Layer = mLocalStyle.Layer
If IsPropertySet(TextPropertyLayer) Then Layer = mLayer

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LeftMargin() As Dimension
Const ProcName As String = "LeftMargin"

On Error GoTo Err

Set LeftMargin = mSeries.LeftMargin
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyLeftMargin) Then Set LeftMargin = mLocalStyle.LeftMargin
If IsPropertySet(TextPropertyLeftMargin) Then Set LeftMargin = mLeftMargin

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let LeftMargin(ByVal Value As Dimension)
Const ProcName As String = "LeftMargin"

On Error GoTo Err

If Value Is Nothing And (Not IsPropertySet(TextPropertyLeftMargin)) Then
    Exit Property
ElseIf Value Is Nothing Then
    Set mLeftMargin = Nothing
    ClearProperty TextPropertyLeftMargin
ElseIf IsPropertySet(TextPropertyLeftMargin) And Value.Equals(mOffset) Then
    Exit Property
Else
    Set mLeftMargin = Value
    setPropertyFlag TextPropertyLeftMargin
End If
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LocalStyle() As TextStyle
Set LocalStyle = mLocalStyle
End Property

Public Property Let LocalStyle( _
                ByVal Value As TextStyle)
Const ProcName As String = "LocalStyle"

On Error GoTo Err

Set mLocalStyle = Value
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MultiLine() As Boolean
Const ProcName As String = "MultiLine"

On Error GoTo Err

MultiLine = mSeries.MultiLine
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyMultiLine) Then MultiLine = mLocalStyle.MultiLine
If IsPropertySet(TextPropertyMultiLine) Then MultiLine = mMultiLine

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let MultiLine(ByVal Value As Boolean)
Const ProcName As String = "MultiLine"

On Error GoTo Err

If mMultiLine = Value And IsPropertySet(TextPropertyMultiLine) Then Exit Property
mMultiLine = Value
setPropertyFlag TextPropertyMultiLine
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Offset() As Size
Const ProcName As String = "Offset"

On Error GoTo Err

Set Offset = mSeries.Offset
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyOffset) Then Set Offset = mLocalStyle.Offset
If IsPropertySet(TextPropertyOffset) Then Set Offset = mOffset

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Offset(ByVal Value As Size)
Const ProcName As String = "Offset"

On Error GoTo Err

If Value Is Nothing And (Not IsPropertySet(TextPropertyOffset)) Then
    Exit Property
ElseIf Value Is Nothing Then
    Set mOffset = Nothing
    ClearProperty TextPropertyOffset
ElseIf IsPropertySet(TextPropertyOffset) And Value.Equals(mOffset) Then
    Exit Property
Else
    Set mOffset = Value
    setPropertyFlag TextPropertyOffset
End If
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let PaddingX(ByVal Value As Double)
Const ProcName As String = "PaddingX"

On Error GoTo Err

If mPaddingX = Value And IsPropertySet(TextPropertyPaddingX) Then Exit Property
mPaddingX = Value
setPropertyFlag TextPropertyPaddingX
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PaddingX() As Double
Const ProcName As String = "PaddingX"

On Error GoTo Err

PaddingX = mSeries.PaddingX
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyPaddingX) Then PaddingX = mLocalStyle.PaddingX
If IsPropertySet(TextPropertyPaddingX) Then PaddingX = mPaddingX

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let PaddingY(ByVal Value As Double)
Const ProcName As String = "PaddingY"

On Error GoTo Err

If mPaddingY = Value And IsPropertySet(TextPropertyPaddingY) Then Exit Property
mPaddingY = Value
setPropertyFlag TextPropertyPaddingY
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PaddingY() As Double
Const ProcName As String = "PaddingY"

On Error GoTo Err

PaddingY = mSeries.PaddingY
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyPaddingY) Then PaddingY = mLocalStyle.PaddingY
If IsPropertySet(TextPropertyPaddingY) Then PaddingY = mPaddingY

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Position() As Point
Set Position = mPosition
End Property

Public Property Let Position(ByVal Value As Point)
Const ProcName As String = "Position"

On Error GoTo Err

If Value Is Nothing Then
    Set mPosition = Nothing
    mBoundingRect.isValid = False
'    mBlank = True
    initiateRedraw True, True
    Exit Property
ElseIf Not mPosition Is Nothing Then
    If Value.Equals(mPosition) Then
        If mPosition.CoordinateSystemX = CoordsLogical And mPosition.CoordinateSystemY = CoordsLogical Then Exit Property
    End If
End If

'If mText <> "" Then mBlank = False
Set mPosition = Value
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get RightMargin() As Dimension
Const ProcName As String = "RightMargin"

On Error GoTo Err

Set RightMargin = mSeries.RightMargin
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyRightMargin) Then Set RightMargin = mLocalStyle.RightMargin
If IsPropertySet(TextPropertyRightMargin) Then Set RightMargin = mRightMargin

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let RightMargin(ByVal Value As Dimension)
Const ProcName As String = "RightMargin"

On Error GoTo Err

If Value Is Nothing And (Not IsPropertySet(TextPropertyRightMargin)) Then
    Exit Property
ElseIf Value Is Nothing Then
    Set mRightMargin = Nothing
    ClearProperty TextPropertyRightMargin
ElseIf IsPropertySet(TextPropertyRightMargin) And Value.Equals(mOffset) Then
    Exit Property
Else
    Set mRightMargin = Value
    setPropertyFlag TextPropertyRightMargin
End If
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Size(ByVal Value As Size)
Const ProcName As String = "Size"

On Error GoTo Err

If Value Is Nothing And (Not IsPropertySet(TextPropertySize)) Then
    Exit Property
ElseIf Value Is Nothing Then
    Set mSize = Nothing
    ClearProperty TextPropertySize
ElseIf IsPropertySet(TextPropertySize) And Value.Equals(mOffset) Then
    Exit Property
Else
    Set mSize = Value
    setPropertyFlag TextPropertySize
End If
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Size() As Size
Const ProcName As String = "Size"

On Error GoTo Err

Set Size = mSeries.Size
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertySize) Then Set Size = mLocalStyle.Size
If IsPropertySet(TextPropertySize) Then Set Size = mSize

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Text(Value As String)
Const ProcName As String = "Text"

On Error GoTo Err

If Value = mText Then Exit Property
mText = Value
'mBlank = (value = "" Or mPosition Is Nothing)
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get TabWidth() As Long
Const ProcName As String = "TabWidth"

On Error GoTo Err

TabWidth = mSeries.TabWidth
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyTabWidth) Then TabWidth = mLocalStyle.TabWidth
If IsPropertySet(TextPropertyTabWidth) Then TabWidth = mTabWidth

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let TabWidth(ByVal Value As Long)
Const ProcName As String = "TabWidth"

On Error GoTo Err

If mTabWidth = Value And IsPropertySet(TextPropertyTabWidth) Then Exit Property
mTabWidth = Value
setPropertyFlag TextPropertyTabWidth
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Text() As String
Text = mText
End Property

Public Property Get Visible() As Boolean
Visible = mVisible
End Property

Public Property Get WordWrap() As Boolean
Const ProcName As String = "WordWrap"

On Error GoTo Err

WordWrap = mSeries.WordWrap
If Not mLocalStyle Is Nothing Then If mLocalStyle.IsPropertySet(TextPropertyWordWrap) Then WordWrap = mLocalStyle.WordWrap
If IsPropertySet(TextPropertyWordWrap) Then WordWrap = mWordWrap

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let WordWrap(ByVal Value As Boolean)
Const ProcName As String = "WordWrap"

On Error GoTo Err

If mWordWrap = Value And IsPropertySet(TextPropertyWordWrap) Then Exit Property
mWordWrap = Value
setPropertyFlag TextPropertyWordWrap
initiateRedraw True, True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Public Sub ClearOverrides()
Const ProcName As String = "ClearOverrides"

On Error GoTo Err

mPropertyOverrideFlags = 0
initiateRedraw True, True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Draw()
Const ProcName As String = "Draw"

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pSeries As TextSeries, _
                ByVal pKey As String, _
                ByVal pGOSP As GraphObjServiceProvider, _
                ByVal pViewport As Viewport, _
                ByVal pLocalStyle As TextStyle)
Const ProcName As String = "Initialise"

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub adjustRectSize( _
                ByRef pRect As GDI_RECT, _
                ByRef pSizeAdjustment As GDI_POINT)
pRect.Right = pRect.Right + pSizeAdjustment.X
If pRect.Right < pRect.Left Then pRect.Right = pRect.Left

pRect.Bottom = pRect.Bottom + pSizeAdjustment.Y
If pRect.Bottom = pRect.Top Then pRect.Bottom = pRect.Top
End Sub

Private Sub calcBoundingRect()
Dim alignmentOffset As GDI_POINT
Dim lTextSize As Size
Dim lTextRect As GDI_RECT
Dim lBoxRect As GDI_RECT
Dim lBoundingRect As GDI_RECT
Dim sizeAdjustment As GDI_POINT

Const ProcName As String = "calcBoundingRect"

On Error GoTo Err

If mPosition Is Nothing Then
    mBoundingRect.isValid = False
    Exit Sub
End If

Set mPosition = convertPosition(mPosition)

mCanvas.SetTextAttributes Color, Font, True

Set lTextSize = mCanvas.GetTextSize(mText, _
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
    .Left = mPosition.XLogical(mViewport.Region)
    .Right = .Left + lTextSize.WidthLogical(mViewport.Region)
    .Top = mPosition.yLogical(mViewport.Region)
    .Bottom = .Top - lTextSize.HeightLogical(mViewport.Region)
    .isValid = True
End With
    
If Not mOffset Is Nothing Then
    RectOffsetBySize mTextTRect, mOffset, mViewport
End If
    
lTextRect = mViewport.ConvertTRectangleToRect(mTextTRect)

If Box Then
    
    lBoxRect = lTextRect
    
    TWWin32API.InflateRect lBoxRect, _
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

'If Not Size Is Nothing Then
'    sizeAdjustment = getSizeAdjustment(lBoundingRect)
'
'    adjustRectSize lTextRect, sizeAdjustment
'    adjustRectSize lBoxRect, sizeAdjustment
'    adjustRectSize lBoundingRect, sizeAdjustment
'End If

alignmentOffset = getAlignmentOffset(lTextRect, lBoxRect)

TWWin32API.OffsetRect lTextRect, alignmentOffset.X, alignmentOffset.Y
TWWin32API.OffsetRect lBoxRect, alignmentOffset.X, alignmentOffset.Y
TWWin32API.OffsetRect lBoundingRect, alignmentOffset.X, alignmentOffset.Y

lBoundingRect.Right = lBoundingRect.Right + 1
lBoundingRect.Bottom = lBoundingRect.Bottom + 1

If Angle <> 0 Then
    Dim alignPoint As GDI_POINT
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function canBeDrawn() As Boolean
Const ProcName As String = "canBeDrawn"

On Error GoTo Err

'If mBlank Then Exit Function
If mPosition Is Nothing Then Exit Function
If Not mVisible Then Exit Function
canBeDrawn = True

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub ClearProperty( _
                ByVal flag As TextPropertyFlags)
Const ProcName As String = "ClearProperty"

On Error GoTo Err

mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function convertPosition(ByVal pPoint As Point) As Point
Const ProcName As String = "convertPosition"
On Error GoTo Err

Set convertPosition = pPoint.ChangeCoordinateSystem( _
                                            IIf(FixedX, CoordinateSystems.CoordsRelative, CoordinateSystems.CoordsLogical), _
                                            IIf(FixedY, CoordinateSystems.CoordsRelative, CoordinateSystems.CoordsLogical), _
                                            mViewport.Region)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub drawBox()
Const ProcName As String = "drawBox"

On Error GoTo Err

If BoxFillWithBackgroundColor Then
    drawBoxWithBackgroundFillColor
Else
    drawBoxWithSpecifiedFillStyle BoxFillStyle
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub drawBoxWithBackgroundFillColor()
Const ProcName As String = "drawBoxWithBackgroundFillColor"

On Error GoTo Err

mCanvas.SaveClippingRegion
mCanvas.IntersectClippingRegion mViewport.ConvertTRectangleToRect(mBoxTRect)

mViewport.PaintBackground
drawBoxWithSpecifiedFillStyle FillTransparent

mCanvas.RestoreClippingRegion
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub drawBoxWithSpecifiedFillStyle( _
                ByVal pFillStyle As FillStyles)
Const ProcName As String = "drawBoxWithSpecifiedFillStyle"

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub DrawText()
Const ProcName As String = "DrawText"

On Error GoTo Err

mCanvas.SetTextAttributes Color, Font, True
mCanvas.DrawText mText, _
                mTextTRect, _
                Justification, _
                MultiLine, _
                Ellipsis, _
                ExpandTabs, _
                TabWidth, _
                WordWrap, _
                LeftMargin, _
                RightMargin

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function getAlignmentOffset( _
                ByRef pTextBoundary As GDI_RECT, _
                ByRef pBoxBoundary As GDI_RECT) As GDI_POINT

Const ProcName As String = "getAlignmentOffset"

On Error GoTo Err

getAlignmentOffset = GDI_POINTSubtract(W32RectTopLeft(pTextBoundary), getAlignmentPointW32(pTextBoundary, pBoxBoundary))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getAlignmentPoint( _
                ByRef pTextBoundary As TRectangle, _
                ByRef pBoxBoundary As TRectangle) As TPoint
Const ProcName As String = "getAlignmentPoint"

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getAlignmentPointW32( _
                ByRef pTextBoundary As GDI_RECT, _
                ByRef pBoxBoundary As GDI_RECT) As GDI_POINT
Const ProcName As String = "getAlignmentPointW32"

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getSizeAdjustment( _
                ByRef lBoundingRect As GDI_RECT) As GDI_POINT
If Me.Size.Width <> 0 Then
    getSizeAdjustment.X = mViewport.ConvertLogicalToPixelsX(Me.Size.WidthLogical(mViewport.Region)) - _
                            (lBoundingRect.Right - lBoundingRect.Left)
End If
If Me.Size.Height <> 0 Then
    getSizeAdjustment.Y = mViewport.ConvertLogicalToPixelsX(Me.Size.HeightLogical(mViewport.Region)) - _
                            (lBoundingRect.Bottom - lBoundingRect.Top)
End If
End Function

Private Sub initiateRedraw( _
                ByVal undrawCurrentImage As Boolean, _
                ByVal recalcBoundaries As Boolean)
Const ProcName As String = "initiateRedraw"

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function isExtended() As Boolean
isExtended = Extended Or FixedX
End Function

Private Function IsPropertySet( _
                ByVal flag As TextPropertyFlags) As Boolean
Const ProcName As String = "IsPropertySet"

On Error GoTo Err

IsPropertySet = gIsFlagSet(mPropertyOverrideFlags, flag)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub setPropertyFlag( _
                ByVal flag As TextPropertyFlags)
Const ProcName As String = "setPropertyFlag"

On Error GoTo Err

mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub undraw()
Const ProcName As String = "undraw"

On Error GoTo Err

If canBeDrawn Then mGOSP.ObjectUndrawn mHandle, mBoundingRect

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

