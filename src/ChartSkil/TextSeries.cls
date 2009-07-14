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

Event SelectionStateChanged()

'================================================================================
' Enums
'================================================================================

Private Enum PropertyOverrideFlags
    ColorIsSet = 1
    BoxIsSet = 2
    BoxColorIsSet = 4
    BoxStyleIsSet = 8
    BoxThicknessIsSet = &H10&
    BoxFillColorIsSet = &H20&
    BoxFillStyleIsSet = &H40&
    AlignIsSet = &H80&
    PaddingXIsSet = &H100&
    PaddingYIsSet = &H200&
    FontIsSet = &H400&
    BoxFillWithBackGroundColorIsSet = &H800&
End Enum

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

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished As Boolean

Private mTexts As Collection

Private mCanvas As Canvas
Private mLayer As Long

Private mGOSP As GraphObjServiceProvider

Private WithEvents mStyle As TextStyle
Attribute mStyle.VB_VarHelpID = -1

Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

Private mName As String

Private mTaskContext As TaskContext

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

Private mPropertyOverrideFlags As DataPointPropertyOverrideFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTexts = New Collection
Set mStyle = New TextStyle

End Sub

Private Sub Class_Terminate()
Debug.Print "TextSeries terminated"
End Sub

'================================================================================
' IGraphicObjectSeries Interface Members
'================================================================================

Private Property Let IGraphicObjectSeries_Canvas(ByVal value As Canvas)
Set mCanvas = value
End Property

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
Set mCanvas = Nothing
End Sub

Private Sub IGraphicObjectSeries_Initialise( _
                ByVal Name As String, _
                ByVal Canvas As Canvas, _
                ByVal Layer As LayerNumbers, _
                ByVal pChart As Chart, _
                ByVal SeriesID As Long, _
                ByVal GOSP As GraphObjServiceProvider)
mName = Name
mSeriesID = SeriesID
Set mGOSP = GOSP
mLayer = Layer
Set mCanvas = Canvas
End Sub

Private Property Get IGraphicObjectSeries_IsSelectable() As Boolean
IGraphicObjectSeries_IsSelectable = mIsSelectable
End Property

Private Function IGraphicObjectSeries_item(ByVal index As Long) As IGraphicObject
Set IGraphicObjectSeries_item = mTexts.Item(index)
End Function

Private Property Get IGraphicObjectSeries_Name() As String
IGraphicObjectSeries_Name = mName
End Property

Public Sub IGraphicObjectSeries_Remove(ByVal index As Long)
mTexts.Remove index
End Sub

Private Sub IGraphicObjectSeries_RemoveObject(ByVal value As IGraphicObject)
Dim lText As ChartSkil26.Text

On Error Resume Next
Set lText = value

mTexts.Remove lText.Key
End Sub

Private Property Let IGraphicObjectSeries_Selected(ByVal RHS As Boolean)
If RHS <> mIsSelected Then
    mIsSelected = RHS
    RaiseEvent SelectionStateChanged
End If
End Property

Private Property Get IGraphicObjectSeries_Selected() As Boolean
IGraphicObjectSeries_Selected = mIsSelected
End Property

Private Property Get IGraphicObjectSeries_SeriesID() As Long
IGraphicObjectSeries_SeriesID = mSeriesID
End Property

'================================================================================
' mStyle Event Handlers
'================================================================================

Private Sub mStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
Dim failpoint As Long
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case UCase$("align")
    repaintRegion True
Case UCase$("box")
    repaintRegion True
Case UCase$("boxColor")
    repaintRegion False
Case UCase$("boxFillColor")
    repaintRegion False
Case UCase$("boxFillStyle")
    repaintRegion False
Case UCase$("boxStyle")
    repaintRegion True
Case UCase$("boxThickness")
    repaintRegion True
Case UCase$("color")
    repaintRegion False
Case UCase$("Font")
    repaintRegion True
Case UCase$("paddingX")
    repaintRegion True
Case UCase$("paddingY")
    repaintRegion True
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & "mStyle_PropertyChanged", _
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

Public Property Get Align() As TextAlignModes
Align = mStyle.Align
If isPropertyOverrideFlagSet(AlignIsSet) Then Align = mAlign
End Property

Public Property Let Align(ByVal value As TextAlignModes)
If mAlign = value And isPropertyOverrideFlagSet(AlignIsSet) Then Exit Property
mAlign = value
setPropertyOverrideFlag AlignIsSet
repaintRegion True
End Property

Public Property Get Box() As Boolean
Box = mStyle.Box
If isPropertyOverrideFlagSet(BoxIsSet) Then Box = mBox
End Property

Public Property Let Box(ByVal value As Boolean)
If mBox = value And isPropertyOverrideFlagSet(BoxIsSet) Then Exit Property
mBox = value
setPropertyOverrideFlag BoxIsSet
repaintRegion True
End Property

Public Property Get BoxColor() As OLE_COLOR
BoxColor = mStyle.BoxColor
If isPropertyOverrideFlagSet(BoxColorIsSet) Then BoxColor = mBoxColor
End Property

Public Property Let BoxColor(ByVal value As OLE_COLOR)
If mBoxColor = value And isPropertyOverrideFlagSet(BoxColorIsSet) Then Exit Property
mBoxColor = value
setPropertyOverrideFlag BoxColorIsSet
repaintRegion False
End Property

Public Property Get BoxFillColor() As OLE_COLOR
BoxFillColor = mStyle.BoxFillColor
If isPropertyOverrideFlagSet(BoxFillColorIsSet) Then BoxFillColor = mBoxFillColor
End Property

Public Property Let BoxFillColor(ByVal value As OLE_COLOR)
If mBoxFillColor = value And isPropertyOverrideFlagSet(BoxFillColorIsSet) Then Exit Property
mBoxFillColor = value
setPropertyOverrideFlag BoxFillColorIsSet
repaintRegion False
End Property

Public Property Get BoxFillStyle() As FillStyles
BoxFillStyle = mStyle.BoxFillStyle
If isPropertyOverrideFlagSet(BoxFillStyleIsSet) Then BoxFillStyle = mBoxFillStyle
End Property

Public Property Let BoxFillStyle(ByVal value As FillStyles)
If mBoxFillStyle = value And isPropertyOverrideFlagSet(BoxFillStyleIsSet) Then Exit Property
mBoxFillStyle = value
setPropertyOverrideFlag BoxFillStyleIsSet
repaintRegion False
End Property

Public Property Get BoxStyle() As LineStyles
BoxStyle = mStyle.BoxStyle
If isPropertyOverrideFlagSet(BoxStyleIsSet) Then BoxStyle = mBoxStyle
End Property

Public Property Let BoxStyle(ByVal value As LineStyles)
If mBoxStyle = value And isPropertyOverrideFlagSet(BoxStyleIsSet) Then Exit Property
mBoxStyle = value
setPropertyOverrideFlag BoxStyleIsSet
repaintRegion True
End Property

Public Property Get BoxThickness() As Long
BoxThickness = mStyle.BoxThickness
If isPropertyOverrideFlagSet(BoxThicknessIsSet) Then BoxThickness = mBoxThickness
End Property

Public Property Let BoxThickness(ByVal value As Long)
If mBoxThickness = value And isPropertyOverrideFlagSet(BoxThicknessIsSet) Then Exit Property
mBoxThickness = value
setPropertyOverrideFlag BoxThicknessIsSet
repaintRegion True
End Property

Public Property Get BoxFillWithBackgroundColor() As Boolean
BoxFillWithBackgroundColor = mStyle.BoxFillWithBackgroundColor
If isPropertyOverrideFlagSet(BoxFillWithBackGroundColorIsSet) Then BoxFillWithBackgroundColor = mBoxFillWithBackgroundColor
End Property

Public Property Let BoxFillWithBackgroundColor(ByVal value As Boolean)
If mBoxFillWithBackgroundColor = value And isPropertyOverrideFlagSet(BoxFillWithBackGroundColorIsSet) Then Exit Property
mBoxFillWithBackgroundColor = value
setPropertyOverrideFlag BoxFillWithBackGroundColorIsSet
repaintRegion False
End Property

Public Property Get Color() As OLE_COLOR
Color = mStyle.Color
If isPropertyOverrideFlagSet(ColorIsSet) Then Color = mColor
End Property

Public Property Let Color(ByVal value As OLE_COLOR)
If mColor = value And isPropertyOverrideFlagSet(ColorIsSet) Then Exit Property
mColor = value
setPropertyOverrideFlag ColorIsSet
repaintRegion False
End Property

Public Property Get Count() As Long
Count = mTexts.Count
End Property

Public Property Get Extended() As Boolean
Extended = mExtended
End Property

Public Property Let Extended(ByVal value As Boolean)
If mExtended = value Then Exit Property
mExtended = value
repaintRegion True
End Property

Public Property Get FixedX() As Boolean
FixedX = mFixedX
End Property

Public Property Let FixedX(ByVal value As Boolean)
If mFixedX = value Then Exit Property
mFixedX = value
repaintRegion True
End Property

Public Property Get FixedY() As Boolean
FixedY = mFixedY
End Property

Public Property Let FixedY(ByVal value As Boolean)
If mFixedY = value Then Exit Property
mFixedY = value
repaintRegion True
End Property

Public Property Get Font() As StdFont
Set Font = mStyle.Font
If isPropertyOverrideFlagSet(FontIsSet) Then Set Font = mFont
End Property

Public Property Let Font(ByVal value As StdFont)
If mFont Is value And isPropertyOverrideFlagSet(FontIsSet) Then Exit Property
Set mFont = value
If mFont Is Nothing Then
    clearPropertyOverrideFlag FontIsSet
Else
    setPropertyOverrideFlag FontIsSet
End If
repaintRegion True
End Property

Friend Property Get Id() As Long
Id = mSeriesID
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value Then Exit Property
mIncludeInAutoscale = value
repaintRegion True
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

Public Property Let Layer(ByVal value As LayerNumbers)
mLayer = value
repaintRegion False
End Property

Public Property Get Layer() As LayerNumbers
Layer = mLayer
End Property

Public Property Get PaddingX() As Double
PaddingX = mStyle.PaddingX
If isPropertyOverrideFlagSet(PaddingXIsSet) Then PaddingX = mPaddingX
End Property

Public Property Let PaddingX(ByVal value As Double)
If mPaddingX = value And isPropertyOverrideFlagSet(PaddingXIsSet) Then Exit Property
mPaddingX = value
setPropertyOverrideFlag PaddingXIsSet
repaintRegion True
End Property

Public Property Get PaddingY() As Double
PaddingY = mStyle.PaddingY
If isPropertyOverrideFlagSet(PaddingYIsSet) Then PaddingY = mPaddingY
End Property

Public Property Let PaddingY(ByVal value As Double)
If mPaddingY = value And isPropertyOverrideFlagSet(PaddingYIsSet) Then Exit Property
mPaddingY = value
setPropertyOverrideFlag PaddingYIsSet
repaintRegion True
End Property

Public Property Get Style() As TextStyle
Set Style = mStyle
End Property

Public Property Let Style(ByVal value As TextStyle)
Dim failpoint As Long
On Error GoTo Err

If value Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "Style", _
            "Style is null reference"
End If

Set mStyle = value
repaintRegion True

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Style" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                Optional ByVal pKey As String, _
                Optional ByVal pLocalStyle As TextStyle) As Text
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil26.TextSeries::Add", _
                            "Text series is finished"

If pKey = "" Then pKey = GenerateGUIDString

Set Add = New Text
Add.Initialise Me, pKey, mGOSP, mCanvas, pLocalStyle

mTexts.Add Add, pKey
mGOSP.AddGraphicObject Add, mSeriesID
End Function

Public Sub ClearOverrides()
mPropertyOverrideFlags = 0
repaintRegion True
End Sub

Friend Sub Finish()
mFinished = True
Set mCanvas = Nothing
End Sub

Public Function Item(ByVal index As Variant) As Text
Attribute Item.VB_UserMemId = 0
Set Item = mTexts(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mTexts.[_NewEnum]
End Function

Public Sub RemoveObject(ByVal value As ChartSkil26.Text)
On Error GoTo Err
mTexts.Remove value.Key
mGOSP.RemoveObject value.Handle

Exit Sub

Err:
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As PropertyOverrideFlags)
mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)
End Sub

Private Function isPropertyOverrideFlagSet( _
                ByVal flag As PropertyOverrideFlags) As Boolean
isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)
End Function

Private Sub repaintRegion( _
                ByVal recalcBoundaries As Boolean)
If mTexts.Count <> 0 Then
    If recalcBoundaries Then
        mGOSP.RecalcBoundary mSeriesID
        mGOSP.PaintRegion True
    Else
        mGOSP.PaintRegion False
    End If
End If
End Sub

Private Sub setPropertyOverrideFlag( _
                ByVal flag As PropertyOverrideFlags)
mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)
End Sub



