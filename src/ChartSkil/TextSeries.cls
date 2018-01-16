VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements IGraphicObjectSeries

'================================================================================
' Events
'================================================================================

Event Click()

Event DblCLick()

Event PropertyChanged(ev As PropertyChangedEventData)

Event SelectionStateChanged()

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Constants
'================================================================================


Private Const ModuleName                As String = "TextSeries"

'================================================================================
' Member variables
'================================================================================

Private mSeriesID As Long
Private mSeriesUpdateNumber As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished As Boolean

Private mTexts As Collection

Private mViewport As ViewPort
Private mLayer As Long

Private mGOSP As GraphObjServiceProvider

Private WithEvents mStyle As TextStyle
Attribute mStyle.VB_VarHelpID = -1
Private mBaseStyle As TextStyle

Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

Private mName As String

' these fields override corresponding items in the Style
Private mColor As Long
Private mBox As Boolean
Private mBoxColor As Long
Private mBoxStyle As LineStyles
Private mBoxThickness As Long
Private mBoxFillColor As Long
Private mBoxFillWithBackgroundColor As Boolean
Private mBoxFillStyle As FillStyles
Private mAlign As TextAlignModes
Private mPaddingX As Double
Private mPaddingY As Double
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1
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

Private mOffset As Size
Private mSize As Size

Private mPropertyOverrideFlags As TextPropertyFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTexts = New Collection
Set mBaseStyle = gDefaultTextStyle
Set mStyle = New TextStyle
End Sub

Private Sub Class_Terminate()
gLogger.Log pLogLevel:=LogLevelHighDetail, pProcName:="Class_Terminate", pModName:=ModuleName, pMsg:="TextSeries terminated"
Debug.Print "TextSeries terminated"
End Sub

'================================================================================
' IGraphicObjectSeries Interface Members
'================================================================================

Private Function IGraphicObjectSeries_Add() As IGraphicObject
Set IGraphicObjectSeries_Add = Add
End Function

Private Sub IGraphicObjectSeries_Click()
RaiseEvent Click
End Sub

Private Sub IGraphicObjectSeries_DblCLick()
RaiseEvent DblCLick
End Sub

Private Property Get IGraphicObjectSeries_count() As Long
IGraphicObjectSeries_count = mTexts.Count
End Property

Private Sub IGraphicObjectSeries_finish()
mFinished = True
Set mGOSP = Nothing
Set mViewport = Nothing
End Sub

Private Sub IGraphicObjectSeries_Initialise( _
                ByVal pName As String, _
                ByVal pViewport As ViewPort, _
                ByVal pLayer As LayerNumbers, _
                ByVal pPeriods As Periods, _
                ByVal pSeriesID As Long, _
                ByVal pGOSP As GraphObjServiceProvider)
mName = pName
mSeriesID = pSeriesID
Set mGOSP = pGOSP
Layer = pLayer
Set mViewport = pViewport
End Sub

Private Property Get IGraphicObjectSeries_IsSelectable() As Boolean
IGraphicObjectSeries_IsSelectable = mIsSelectable
End Property

Private Function IGraphicObjectSeries_item(ByVal index As Long) As IGraphicObject
Const ProcName As String = "IGraphicObjectSeries_item"

On Error GoTo Err

Set IGraphicObjectSeries_item = mTexts.Item(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IGraphicObjectSeries_Name() As String
IGraphicObjectSeries_Name = mName
End Property

Private Sub IGraphicObjectSeries_Remove(ByVal index As Long)
Const ProcName As String = "IGraphicObjectSeries_Remove"

On Error GoTo Err

mTexts.Remove index

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IGraphicObjectSeries_RemoveObject(ByVal Value As IGraphicObject)
Dim lText As ChartSkil27.Text

Const ProcName As String = "IGraphicObjectSeries_RemoveObject"

On Error GoTo Err

On Error Resume Next
Set lText = Value

mTexts.Remove lText.Key

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let IGraphicObjectSeries_Selected(ByVal RHS As Boolean)
Const ProcName As String = "IGraphicObjectSeries_Selected"

On Error GoTo Err

If RHS <> mIsSelected Then
    mIsSelected = RHS
    firePropertyChanged "IsSelected"
    RaiseEvent SelectionStateChanged
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IGraphicObjectSeries_Selected() As Boolean
IGraphicObjectSeries_Selected = mIsSelected
End Property

Private Property Get IGraphicObjectSeries_SeriesID() As Long
IGraphicObjectSeries_SeriesID = mSeriesID
End Property

'================================================================================
' mFont Event Handlers
'================================================================================

Private Sub mFont_FontChanged(ByVal PropertyName As String)
firePropertyChanged "Font"
End Sub

'================================================================================
' mStyle Event Handlers
'================================================================================

Private Sub mStyle_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mStyle_PropertyChanged"

On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case "ALIGN"
    RepaintRegion True
Case "ANGLE"
    RepaintRegion True
Case "BOX"
    RepaintRegion True
Case "BOXCOLOR"
    RepaintRegion False
Case "BOXFILLCOLOR"
    RepaintRegion False
Case "BOXFILLSTYLE"
    RepaintRegion False
Case "BOXFILLWITHBACKGROUNDCOLOR"
    RepaintRegion False
Case "BOXSTYLE"
    RepaintRegion True
Case "BOXTHICKNESS"
    RepaintRegion True
Case "COLOR"
    RepaintRegion False
Case "ELLIPSIS"
    RepaintRegion False
Case "EXPANDTABS"
    RepaintRegion True
Case "EXTENDED"
    RepaintRegion True
Case "FIXEDX"
    RepaintRegion True
Case "FIXEDY"
    RepaintRegion True
Case "FONT"
    RepaintRegion True
Case "HIDEIFBLANK"
    RepaintRegion True
Case "INCLUDEINAUTOSCALE"
    RepaintRegion True
Case "JUSTIFICATION"
    RepaintRegion False
Case "LAYER"
    RepaintRegion True
Case "LEFTMARGIN"
    RepaintRegion True
Case "MULTILINE"
    RepaintRegion True
Case "OFFSET"
    RepaintRegion True
Case "PADDINGX"
    RepaintRegion True
Case "PADDINGY"
    RepaintRegion True
Case "RIGHTMARGIN"
    RepaintRegion True
Case "SIZE"
    RepaintRegion True
Case "TABWIDTH"
    RepaintRegion True
Case "WORDWRAP"
    RepaintRegion True
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Unhandled property change"
End Select

firePropertyChanged ev.PropertyName

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

Align = mBaseStyle.Align
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyAlign) Then Align = mStyle.Align
If IsPropertySet(TextPropertyAlign) Then Align = mAlign

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Align(ByVal Value As TextAlignModes)
Const ProcName As String = "Align"

On Error GoTo Err

setLongProperty mAlign, TextPropertyAlign, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Angle() As Double
Const ProcName As String = "Angle"

On Error GoTo Err

Angle = mBaseStyle.Angle
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyAngle) Then Angle = mStyle.Angle
If IsPropertySet(TextPropertyAngle) Then Angle = mAngle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Angle(ByVal Value As Double)
Const ProcName As String = "Angle"

On Error GoTo Err

setDoubleProperty mAngle, TextPropertyAngle, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Box() As Boolean
Const ProcName As String = "Box"

On Error GoTo Err

Box = mBaseStyle.Box
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyBox) Then Box = mStyle.Box
If IsPropertySet(TextPropertyBox) Then Box = mBox

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Box(ByVal Value As Boolean)
Const ProcName As String = "Box"

On Error GoTo Err

setBooleanProperty mBox, TextPropertyBox, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BoxColor() As Long
Const ProcName As String = "BoxColor"

On Error GoTo Err

BoxColor = mBaseStyle.BoxColor
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyBoxColor) Then BoxColor = mStyle.BoxColor
If IsPropertySet(TextPropertyBoxColor) Then BoxColor = mBoxColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BoxColor(ByVal Value As Long)
Const ProcName As String = "BoxColor"

On Error GoTo Err

setLongProperty mBoxColor, TextPropertyBoxColor, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BoxFillColor() As Long
Const ProcName As String = "BoxFillColor"

On Error GoTo Err

BoxFillColor = mBaseStyle.BoxFillColor
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyBoxFillColor) Then BoxFillColor = mStyle.BoxFillColor
If IsPropertySet(TextPropertyBoxFillColor) Then BoxFillColor = mBoxFillColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BoxFillColor(ByVal Value As Long)
Const ProcName As String = "BoxFillColor"

On Error GoTo Err

setLongProperty mBoxFillColor, TextPropertyBoxFillColor, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BoxFillStyle() As FillStyles
Const ProcName As String = "BoxFillStyle"

On Error GoTo Err

BoxFillStyle = mBaseStyle.BoxFillStyle
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyBoxFillStyle) Then BoxFillStyle = mStyle.BoxFillStyle
If IsPropertySet(TextPropertyBoxFillStyle) Then BoxFillStyle = mBoxFillStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BoxFillStyle(ByVal Value As FillStyles)
Const ProcName As String = "BoxFillStyle"

On Error GoTo Err

setLongProperty mBoxFillStyle, TextPropertyBoxFillStyle, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BoxFillWithBackgroundColor() As Boolean
Const ProcName As String = "BoxFillWithBackgroundColor"

On Error GoTo Err

BoxFillWithBackgroundColor = mBaseStyle.BoxFillWithBackgroundColor
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyBoxFillWithBackgroundColor) Then BoxFillWithBackgroundColor = mStyle.BoxFillWithBackgroundColor
If IsPropertySet(TextPropertyBoxFillWithBackgroundColor) Then BoxFillWithBackgroundColor = mBoxFillWithBackgroundColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BoxFillWithBackgroundColor(ByVal Value As Boolean)
Const ProcName As String = "BoxFillWithBackgroundColor"

On Error GoTo Err

setBooleanProperty mBoxFillWithBackgroundColor, TextPropertyBoxFillWithBackgroundColor, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BoxStyle() As LineStyles
Const ProcName As String = "BoxStyle"

On Error GoTo Err

BoxStyle = mBaseStyle.BoxStyle
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyBoxStyle) Then BoxStyle = mStyle.BoxStyle
If IsPropertySet(TextPropertyBoxStyle) Then BoxStyle = mBoxStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BoxStyle(ByVal Value As LineStyles)
Const ProcName As String = "BoxStyle"

On Error GoTo Err

setLongProperty mBoxStyle, TextPropertyBoxStyle, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BoxThickness() As Long
Const ProcName As String = "BoxThickness"

On Error GoTo Err

BoxThickness = mBaseStyle.BoxThickness
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyBoxThickness) Then BoxThickness = mStyle.BoxThickness
If IsPropertySet(TextPropertyBoxThickness) Then BoxThickness = mBoxThickness

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BoxThickness(ByVal Value As Long)
Const ProcName As String = "BoxThickness"

On Error GoTo Err

setLongProperty mBoxThickness, TextPropertyBoxThickness, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Color() As Long
Const ProcName As String = "Color"

On Error GoTo Err

Color = mBaseStyle.Color
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyColor) Then Color = mStyle.Color
If IsPropertySet(TextPropertyColor) Then Color = mColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Color(ByVal Value As Long)
Const ProcName As String = "Color"

On Error GoTo Err

setLongProperty mColor, TextPropertyColor, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Count() As Long
Const ProcName As String = "Count"

On Error GoTo Err

Count = mTexts.Count

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Ellipsis(ByVal Value As EllipsisModes)
Const ProcName As String = "Ellipsis"

On Error GoTo Err

setLongProperty mEllipsis, TextPropertyEllipsis, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Ellipsis() As EllipsisModes
Const ProcName As String = "Ellipsis"

On Error GoTo Err

Ellipsis = mBaseStyle.Ellipsis
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyEllipsis) Then Ellipsis = mStyle.Ellipsis
If IsPropertySet(TextPropertyEllipsis) Then Ellipsis = mEllipsis

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ExpandTabs(ByVal Value As Boolean)
Const ProcName As String = "ExpandTabs"

On Error GoTo Err

setBooleanProperty mExpandTabs, TextPropertyExpandTabs, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ExpandTabs() As Boolean
Const ProcName As String = "ExpandTabs"

On Error GoTo Err

ExpandTabs = mBaseStyle.ExpandTabs
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyExpandTabs) Then ExpandTabs = mStyle.ExpandTabs
If IsPropertySet(TextPropertyExpandTabs) Then ExpandTabs = mExpandTabs

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Extended() As Boolean
Const ProcName As String = "Extended"

On Error GoTo Err

Extended = mBaseStyle.Extended
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyExtended) Then Extended = mStyle.Extended
If IsPropertySet(TextPropertyExtended) Then Extended = mExtended

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Extended(ByVal Value As Boolean)
Const ProcName As String = "Extended"

On Error GoTo Err

setBooleanProperty mExtended, TextPropertyExtended, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FixedX() As Boolean
Const ProcName As String = "FixedX"

On Error GoTo Err

FixedX = mBaseStyle.FixedX
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyFixedX) Then FixedX = mStyle.FixedX
If IsPropertySet(TextPropertyFixedX) Then FixedX = mFixedX

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let FixedX(ByVal Value As Boolean)
Const ProcName As String = "FixedX"

On Error GoTo Err

setBooleanProperty mFixedX, TextPropertyFixedX, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FixedY() As Boolean
Const ProcName As String = "FixedY"

On Error GoTo Err

FixedY = mBaseStyle.FixedY
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyFixedY) Then FixedY = mStyle.FixedY
If IsPropertySet(TextPropertyFixedY) Then FixedY = mFixedY

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let FixedY(ByVal Value As Boolean)
Const ProcName As String = "FixedY"

On Error GoTo Err

setBooleanProperty mFixedY, TextPropertyFixedY, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Font() As StdFont
Const ProcName As String = "Font"

On Error GoTo Err

Set Font = mBaseStyle.Font
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyFont) Then Set Font = mStyle.Font
If IsPropertySet(TextPropertyFont) Then Set Font = mFont

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Font(ByVal Value As StdFont)
Const ProcName As String = "Font"

On Error GoTo Err

If mFont Is Value And IsPropertySet(TextPropertyFont) Then Exit Property
Set mFont = Value
If mFont Is Nothing Then
    ClearProperty TextPropertyFont
Else
    setPropertyFlag TextPropertyFont
End If
firePropertyChanged "Font"
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Id() As Long
Id = mSeriesID
End Property

Public Property Get HideIfBlank() As Boolean
Const ProcName As String = "HideIfBlank"

On Error GoTo Err

HideIfBlank = mBaseStyle.HideIfBlank
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyHideIfBlank) Then HideIfBlank = mStyle.HideIfBlank
If IsPropertySet(TextPropertyHideIfBlank) Then HideIfBlank = mHideIfBlank

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HideIfBlank(ByVal Value As Boolean)
Const ProcName As String = "HideIfBlank"

On Error GoTo Err

setBooleanProperty mHideIfBlank, TextPropertyHideIfBlank, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"

On Error GoTo Err

IncludeInAutoscale = mBaseStyle.IncludeInAutoscale
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyIncludeInAutoscale) Then IncludeInAutoscale = mStyle.IncludeInAutoscale
If IsPropertySet(TextPropertyIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"

On Error GoTo Err

setBooleanProperty mIncludeInAutoscale, TextPropertyIncludeInAutoscale, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let IsSelectable(ByVal Value As Boolean)
Const ProcName As String = "IsSelectable"

On Error GoTo Err

mIsSelectable = Value
firePropertyChanged "IsSelectable"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsSelectable() As Boolean
IsSelectable = mIsSelectable
End Property

Public Property Get IsSelected() As Boolean
IsSelected = mIsSelected
End Property

Public Property Let Justification(ByVal Value As TextJustifyModes)
Const ProcName As String = "Justification"

On Error GoTo Err

setLongProperty mJustification, TextPropertyJustification, Value
RepaintRegion False

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Justification() As TextJustifyModes
Justification = mBaseStyle.Justification
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyJustification) Then Justification = mStyle.Justification
If IsPropertySet(TextPropertyJustification) Then Justification = mJustification
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"

On Error GoTo Err

setLongProperty mLayer, TextPropertyLayer, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

Layer = mBaseStyle.Layer
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyLayer) Then Layer = mStyle.Layer
If IsPropertySet(TextPropertyLayer) Then Layer = mLayer

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let LeftMargin(ByVal Value As Dimension)
Const ProcName As String = "LeftMargin"

On Error GoTo Err

setDimensionProperty mLeftMargin, TextPropertyLeftMargin, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LeftMargin() As Dimension
Const ProcName As String = "LeftMargin"
On Error GoTo Err

Set LeftMargin = mBaseStyle.LeftMargin
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyLeftMargin) Then Set LeftMargin = mStyle.LeftMargin
If IsPropertySet(TextPropertyLeftMargin) Then Set LeftMargin = mLeftMargin

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let MultiLine(ByVal Value As Boolean)
Const ProcName As String = "MultiLine"

On Error GoTo Err

setBooleanProperty mMultiLine, TextPropertyMultiLine, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MultiLine() As Boolean
Const ProcName As String = "MultiLine"


On Error GoTo Err

MultiLine = mBaseStyle.MultiLine
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyMultiLine) Then MultiLine = mStyle.MultiLine
If IsPropertySet(TextPropertyMultiLine) Then MultiLine = mMultiLine

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Offset(ByVal Value As Size)
Const ProcName As String = "Offset"

On Error GoTo Err

setSizeProperty mOffset, TextPropertyOffset, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Offset() As Size
Const ProcName As String = "Offset"

On Error GoTo Err

Set Offset = mBaseStyle.Offset
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyOffset) Then Set Offset = mStyle.Offset
If IsPropertySet(TextPropertyOffset) Then Set Offset = mOffset

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PaddingX() As Double
Const ProcName As String = "PaddingX"

On Error GoTo Err

PaddingX = mBaseStyle.PaddingX
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyPaddingX) Then PaddingX = mStyle.PaddingX
If IsPropertySet(TextPropertyPaddingX) Then PaddingX = mPaddingX

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PaddingX(ByVal Value As Double)
Const ProcName As String = "PaddingX"

On Error GoTo Err

setDoubleProperty mPaddingX, TextPropertyPaddingX, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PaddingY() As Double
Const ProcName As String = "PaddingY"

On Error GoTo Err

PaddingY = mBaseStyle.PaddingY
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyPaddingY) Then PaddingY = mStyle.PaddingY
If IsPropertySet(TextPropertyPaddingY) Then PaddingY = mPaddingY

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PaddingY(ByVal Value As Double)
Const ProcName As String = "PaddingY"

On Error GoTo Err

setDoubleProperty mPaddingY, TextPropertyPaddingY, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RightMargin(ByVal Value As Dimension)
Const ProcName As String = "RightMargin"

On Error GoTo Err

setDimensionProperty mRightMargin, TextPropertyRightMargin, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RightMargin() As Dimension
Const ProcName As String = "RightMargin"
On Error GoTo Err

Set RightMargin = mBaseStyle.RightMargin
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyRightMargin) Then Set RightMargin = mStyle.RightMargin
If IsPropertySet(TextPropertyRightMargin) Then Set RightMargin = mRightMargin

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Size() As Size
Const ProcName As String = "Size"

On Error GoTo Err

Set Size = mBaseStyle.Size
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertySize) Then Set Size = mStyle.Size
If IsPropertySet(TextPropertySize) Then Set Size = mSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Size(ByVal Value As Size)
Const ProcName As String = "Size"

On Error GoTo Err

setSizeProperty mSize, TextPropertySize, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get SeriesUpdateNumber() As Long
SeriesUpdateNumber = mSeriesUpdateNumber
End Property

Public Property Get Style() As TextStyle
Set Style = mStyle
End Property

Public Property Let Style(ByVal Value As TextStyle)
Const ProcName As String = "Style"

On Error GoTo Err

Set mStyle = Value
RepaintRegion True

firePropertyChanged "Style"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TabWidth(ByVal Value As Long)
Const ProcName As String = "TabWidth"

On Error GoTo Err

setLongProperty mTabWidth, TextPropertyTabWidth, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TabWidth() As Long
Const ProcName As String = "TabWidth"

On Error GoTo Err

TabWidth = mBaseStyle.TabWidth
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyTabWidth) Then TabWidth = mStyle.TabWidth
If IsPropertySet(TextPropertyTabWidth) Then TabWidth = mTabWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let WordWrap(ByVal Value As Boolean)
Const ProcName As String = "WordWrap"

On Error GoTo Err

setBooleanProperty mWordWrap, TextPropertyWordWrap, Value
RepaintRegion True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get WordWrap() As Boolean
Const ProcName As String = "WordWrap"

On Error GoTo Err

WordWrap = mBaseStyle.WordWrap
If Not mStyle Is Nothing Then If mStyle.IsPropertySet(TextPropertyWordWrap) Then WordWrap = mStyle.WordWrap
If IsPropertySet(TextPropertyWordWrap) Then WordWrap = mWordWrap

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                Optional ByVal pKey As String, _
                Optional ByVal pLocalStyle As TextStyle) As Text
Const ProcName As String = "Add"
Dim lText As Text


On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Text series is finished"

If pKey = "" Then pKey = GenerateGUIDString

Set lText = New Text
lText.Initialise Me, pKey, mGOSP, mViewport, pLocalStyle

mTexts.Add lText, pKey
mGOSP.AddGraphicObject lText

Set Add = lText

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub ClearOverrides()
Const ProcName As String = "ClearOverrides"

On Error GoTo Err

mPropertyOverrideFlags = 0
RepaintRegion True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Finish()
mFinished = True
Set mViewport = Nothing
End Sub

Public Function Item(ByVal index As Variant) As Text
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"

On Error GoTo Err

Set Item = mTexts(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mTexts.[_NewEnum]
End Function

Public Sub RemoveObject(ByVal Value As ChartSkil27.Text)
On Error GoTo Err
mTexts.Remove Value.Key
mGOSP.RemoveObject Value.handle

Exit Sub

Err:
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub ClearProperty( _
                ByVal flag As TextPropertyFlags)
Const ProcName As String = "ClearProperty"

On Error GoTo Err

mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub firePropertyChanged( _
                ByVal Name As String)
Dim ev As PropertyChangedEventData

Const ProcName As String = "firePropertyChanged"

On Error GoTo Err

Set ev.Source = Me
ev.PropertyName = Name
RaiseEvent PropertyChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IsPropertySet( _
                ByVal flag As TextPropertyFlags) As Boolean
Const ProcName As String = "IsPropertySet"

On Error GoTo Err

IsPropertySet = gIsFlagSet(mPropertyOverrideFlags, flag)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub RepaintRegion( _
                ByVal recalcBoundaries As Boolean)
Const ProcName As String = "RepaintRegion"

On Error GoTo Err

If mTexts.Count <> 0 Then
    If recalcBoundaries Then
        mSeriesUpdateNumber = mSeriesUpdateNumber + 1
        mGOSP.RecalcBoundary mSeriesID
        mGOSP.PaintRegion True
    Else
        mGOSP.PaintRegion False
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setBooleanProperty( _
                ByRef pProp As Boolean, _
                ByVal pPropFlag As TextPropertyFlags, _
                ByVal Value As Boolean)
Const ProcName As String = "setBooleanProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
firePropertyChanged gTextPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setDimensionProperty( _
                ByRef pProp As Dimension, _
                ByVal pPropFlag As TextPropertyFlags, _
                ByVal Value As Dimension)
Const ProcName As String = "setDimensionProperty"
On Error GoTo Err

If pProp Is Value And IsPropertySet(pPropFlag) Then Exit Sub
Set pProp = Value
If pProp Is Nothing Then
    ClearProperty pPropFlag
Else
    setPropertyFlag pPropFlag
End If
firePropertyChanged gTextPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setDoubleProperty( _
                ByRef pProp As Double, _
                ByVal pPropFlag As TextPropertyFlags, _
                ByVal Value As Single)
Const ProcName As String = "setDoubleProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
firePropertyChanged gTextPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setLongProperty( _
                ByRef pProp As Long, _
                ByVal pPropFlag As TextPropertyFlags, _
                ByVal Value As Long)
Const ProcName As String = "setLongProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
firePropertyChanged gTextPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSingleProperty( _
                ByRef pProp As Single, _
                ByVal pPropFlag As TextPropertyFlags, _
                ByVal Value As Single)
Const ProcName As String = "setSingleProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
firePropertyChanged gTextPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSizeProperty( _
                ByRef pProp As Size, _
                ByVal pPropFlag As TextPropertyFlags, _
                ByVal Value As Size)
Const ProcName As String = "setSizeProperty"
On Error GoTo Err

If pProp Is Value And IsPropertySet(pPropFlag) Then Exit Sub
Set pProp = Value
If pProp Is Nothing Then
    ClearProperty pPropFlag
Else
    setPropertyFlag pPropFlag
End If
firePropertyChanged gTextPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setPropertyFlag( _
                ByVal flag As TextPropertyFlags)
Const ProcName As String = "setPropertyFlag"

On Error GoTo Err

mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



