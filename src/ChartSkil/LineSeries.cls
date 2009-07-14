VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LineSeries"
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
    ThicknessIsSet = 2
    LineStyleIsSet = 4
    ExtendBeforeIsSet = 8
    ExtendAfterIsSet = &H10&
    ArrowStartStyleIsSet = &H20&
    ArrowStartLengthIsSet = &H40&
    ArrowStartWidthIsSet = &H80&
    ArrowStartColorIsSet = &H100&
    ArrowStartFillColorIsSet = &H200&
    ArrowStartFillStyleIsSet = &H400&
    ArrowEndStyleIsSet = &H800&
    ArrowEndLengthIsSet = &H1000&
    ArrowEndWidthIsSet = &H2000&
    ArrowEndColorIsSet = &H4000&
    ArrowEndFillColorIsSet = &H8000&
    ArrowEndFillStyleIsSet = &H10000
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Constants
'================================================================================


Private Const ModuleName                As String = "LineSeries"

'================================================================================
' Member variables
'================================================================================

Private mName As String
Private mSeriesID As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished As Boolean

Private mLines As Collection

'Private mSurface As PictureBox
Private mCanvas As Canvas
Private mLayer As Long

Private mGOSP As GraphObjServiceProvider

Private WithEvents mStyle As LineStyle
Attribute mStyle.VB_VarHelpID = -1

Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

' these fields override corresponding items in the Style
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

Private mPropertyOverrideFlags As PropertyOverrideFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mLines = New Collection
Set mStyle = New LineStyle

End Sub

Private Sub Class_Terminate()
Debug.Print "LineSeries terminated"
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
IGraphicObjectSeries_count = mLines.Count
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
Set IGraphicObjectSeries_item = mLines.Item(index)
End Function

Private Property Get IGraphicObjectSeries_Name() As String
IGraphicObjectSeries_Name = mName
End Property

Public Sub IGraphicObjectSeries_Remove(ByVal index As Long)
mLines.Remove index
End Sub

Private Sub IGraphicObjectSeries_RemoveObject(ByVal value As IGraphicObject)
Dim lLine As ChartSkil26.Line

On Error Resume Next
Set lLine = value

mLines.Remove lLine.Key
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
Case UCase$("ArrowEndColor")
    repaintRegion False
Case UCase$("ArrowEndFillColor")
    repaintRegion False
Case UCase$("ArrowEndFillStyle")
    repaintRegion False
Case UCase$("ArrowEndLength")
    repaintRegion True
Case UCase$("ArrowEndStyle")
    repaintRegion True
Case UCase$("ArrowEndWidth")
    repaintRegion True
Case UCase$("ArrowStartColor")
    repaintRegion False
Case UCase$("ArrowStartFillColor")
    repaintRegion False
Case UCase$("ArrowStartfillStyle")
    repaintRegion False
Case UCase$("ArrowStartLength")
    repaintRegion True
Case UCase$("ArrowStartStyle")
    repaintRegion True
Case UCase$("ArrowStartWidth")
    repaintRegion True
Case UCase$("Color")
    repaintRegion False
Case UCase$("ExtendAfter")
    repaintRegion True
Case UCase$("ExtendBefore")
    repaintRegion True
Case UCase$("LineStyle")
    repaintRegion False
Case UCase$("Thickness")
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
Err.Raise errNumber, errSource, errDescription

End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get ArrowEndColor() As OLE_COLOR
ArrowEndColor = mStyle.ArrowEndColor
If isPropertyOverrideFlagSet(ArrowEndColorIsSet) Then ArrowEndColor = mArrowEndColor
End Property

Public Property Let ArrowEndColor(ByVal value As OLE_COLOR)
If mArrowEndColor = value And isPropertyOverrideFlagSet(ArrowEndColorIsSet) Then Exit Property
mArrowEndColor = value
setPropertyOverrideFlag ArrowEndColorIsSet
repaintRegion False
End Property

Public Property Get ArrowEndFillColor() As OLE_COLOR
ArrowEndFillColor = mStyle.ArrowEndFillColor
If isPropertyOverrideFlagSet(ArrowEndFillColorIsSet) Then ArrowEndFillColor = mArrowEndFillColor
End Property

Public Property Let ArrowEndFillColor(ByVal value As OLE_COLOR)
If mArrowEndFillColor = value And isPropertyOverrideFlagSet(ArrowEndFillColorIsSet) Then Exit Property
mArrowEndFillColor = value
setPropertyOverrideFlag ArrowEndFillColorIsSet
repaintRegion False
End Property

Public Property Get ArrowEndFillStyle() As FillStyles
ArrowEndFillStyle = mStyle.ArrowEndFillStyle
If isPropertyOverrideFlagSet(ArrowEndFillStyleIsSet) Then ArrowEndFillStyle = mArrowEndFillStyle
End Property

Public Property Let ArrowEndFillStyle(ByVal value As FillStyles)
If mArrowEndFillStyle = value And isPropertyOverrideFlagSet(ArrowEndFillStyleIsSet) Then Exit Property
mArrowEndFillStyle = value
setPropertyOverrideFlag ArrowEndFillStyleIsSet
repaintRegion False
End Property

Public Property Get ArrowEndLength() As Long
ArrowEndLength = mStyle.ArrowEndLength
If isPropertyOverrideFlagSet(ArrowEndLengthIsSet) Then ArrowEndLength = mArrowEndLength
End Property

Public Property Let ArrowEndLength(ByVal value As Long)
If mArrowEndLength = value And isPropertyOverrideFlagSet(ArrowEndLengthIsSet) Then Exit Property
mArrowEndLength = value
setPropertyOverrideFlag ArrowEndLengthIsSet
repaintRegion True
End Property

Public Property Get ArrowEndStyle() As ArrowStyles
ArrowEndStyle = mStyle.ArrowEndStyle
If isPropertyOverrideFlagSet(ArrowEndStyleIsSet) Then ArrowEndStyle = mArrowEndStyle
End Property

Public Property Let ArrowEndStyle(ByVal value As ArrowStyles)
If mArrowEndStyle = value And isPropertyOverrideFlagSet(ArrowEndStyleIsSet) Then Exit Property
mArrowEndStyle = value
setPropertyOverrideFlag ArrowEndStyleIsSet
repaintRegion True
End Property

Public Property Get ArrowEndWidth() As Long
ArrowEndWidth = mStyle.ArrowEndWidth
If isPropertyOverrideFlagSet(ArrowEndWidthIsSet) Then ArrowEndWidth = mArrowEndWidth
End Property

Public Property Let ArrowEndWidth(ByVal value As Long)
If mArrowEndWidth = value And isPropertyOverrideFlagSet(ArrowEndWidthIsSet) Then Exit Property
mArrowEndWidth = value
setPropertyOverrideFlag ArrowEndWidthIsSet
repaintRegion True
End Property

Public Property Get ArrowStartColor() As OLE_COLOR
ArrowStartColor = mStyle.ArrowStartColor
If isPropertyOverrideFlagSet(ArrowStartColorIsSet) Then ArrowStartColor = mArrowStartColor
End Property

Public Property Let ArrowStartColor(ByVal value As OLE_COLOR)
If mArrowStartColor = value And isPropertyOverrideFlagSet(ArrowStartColorIsSet) Then Exit Property
mArrowStartColor = value
setPropertyOverrideFlag ArrowStartColorIsSet
repaintRegion False
End Property

Public Property Get ArrowStartFillColor() As OLE_COLOR
ArrowStartFillColor = mStyle.ArrowStartFillColor
If isPropertyOverrideFlagSet(ArrowStartFillColorIsSet) Then ArrowStartFillColor = mArrowStartFillColor
End Property

Public Property Let ArrowStartFillColor(ByVal value As OLE_COLOR)
If mArrowStartFillColor = value And isPropertyOverrideFlagSet(ArrowStartFillColorIsSet) Then Exit Property
mArrowStartFillColor = value
setPropertyOverrideFlag ArrowStartFillColorIsSet
repaintRegion False
End Property

Public Property Get ArrowStartFillStyle() As FillStyles
ArrowStartFillStyle = mStyle.ArrowStartFillStyle
If isPropertyOverrideFlagSet(ArrowStartFillStyleIsSet) Then ArrowStartFillStyle = mArrowStartfillStyle
End Property

Public Property Let ArrowStartFillStyle(ByVal value As FillStyles)
If mArrowStartfillStyle = value And isPropertyOverrideFlagSet(ArrowStartFillStyleIsSet) Then Exit Property
mArrowStartfillStyle = value
setPropertyOverrideFlag ArrowStartFillStyleIsSet
repaintRegion False
End Property

Public Property Get ArrowStartLength() As Long
ArrowStartLength = mStyle.ArrowStartLength
If isPropertyOverrideFlagSet(ArrowStartLengthIsSet) Then ArrowStartLength = mArrowStartLength
End Property

Public Property Let ArrowStartLength(ByVal value As Long)
If mArrowStartLength = value And isPropertyOverrideFlagSet(ArrowStartLengthIsSet) Then Exit Property
mArrowStartLength = value
setPropertyOverrideFlag ArrowStartLengthIsSet
repaintRegion True
End Property

Public Property Get ArrowStartStyle() As ArrowStyles
ArrowStartStyle = mStyle.ArrowStartStyle
If isPropertyOverrideFlagSet(ArrowStartStyleIsSet) Then ArrowStartStyle = mArrowStartStyle
End Property

Public Property Let ArrowStartStyle(ByVal value As ArrowStyles)
If mArrowStartStyle = value And isPropertyOverrideFlagSet(ArrowStartStyleIsSet) Then Exit Property
mArrowStartStyle = value
setPropertyOverrideFlag ArrowStartStyleIsSet
repaintRegion True
End Property

Public Property Get ArrowStartWidth() As Long
ArrowStartWidth = mStyle.ArrowStartWidth
If isPropertyOverrideFlagSet(ArrowStartWidthIsSet) Then ArrowStartWidth = mArrowStartWidth
End Property

Public Property Let ArrowStartWidth(ByVal value As Long)
If mArrowStartWidth = value And isPropertyOverrideFlagSet(ArrowStartWidthIsSet) Then Exit Property
mArrowStartWidth = value
setPropertyOverrideFlag ArrowStartWidthIsSet
repaintRegion True
End Property

Friend Property Let Canvas(ByVal value As Canvas)
Set mCanvas = value
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
Count = mLines.Count
End Property

Public Property Get Extended() As Boolean
Extended = mExtended
End Property

Public Property Let Extended(ByVal value As Boolean)
If mExtended = value Then Exit Property
mExtended = value
repaintRegion True
End Property

Public Property Get ExtendAfter() As Boolean
ExtendAfter = mStyle.ExtendAfter
If isPropertyOverrideFlagSet(ExtendAfterIsSet) Then ExtendAfter = mExtendAfter
End Property

Public Property Let ExtendAfter(ByVal value As Boolean)
If mExtendAfter = value And isPropertyOverrideFlagSet(ExtendAfterIsSet) Then Exit Property
mExtendAfter = value
setPropertyOverrideFlag ExtendAfterIsSet
repaintRegion True
End Property

Public Property Get ExtendBefore() As Boolean
ExtendBefore = mStyle.ExtendBefore
If isPropertyOverrideFlagSet(ExtendBeforeIsSet) Then ExtendBefore = mExtendBefore
End Property

Public Property Let ExtendBefore(ByVal value As Boolean)
If mExtendBefore = value And isPropertyOverrideFlagSet(ExtendBeforeIsSet) Then Exit Property
mExtendBefore = value
setPropertyOverrideFlag ExtendBeforeIsSet
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

Public Property Get LineStyle() As LineStyles
LineStyle = mStyle.LineStyle
If isPropertyOverrideFlagSet(LineStyleIsSet) Then LineStyle = mLineStyle
End Property

Public Property Let LineStyle(ByVal value As LineStyles)
If mLineStyle = value And isPropertyOverrideFlagSet(LineStyleIsSet) Then Exit Property
mLineStyle = value
setPropertyOverrideFlag LineStyleIsSet
repaintRegion False
End Property

Public Property Get Style() As LineStyle
Set Style = mStyle
End Property

Public Property Let Style(ByVal value As LineStyle)
Dim somethingChanged As Boolean
Dim recalcBoundaries As Boolean
Dim recalcRegion As Boolean

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

Public Property Get Thickness() As Long
Thickness = mStyle.Thickness
If isPropertyOverrideFlagSet(ThicknessIsSet) Then Thickness = mThickness
End Property

Public Property Let Thickness(ByVal value As Long)
If mThickness = value And isPropertyOverrideFlagSet(ThicknessIsSet) Then Exit Property
mThickness = value
setPropertyOverrideFlag ThicknessIsSet
repaintRegion True
End Property

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                Optional ByVal Key As String, _
                Optional ByVal LocalStyle As LineStyle) As ChartSkil26.Line
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil26.LineSeries::Add", _
                            "Line series is finished"
Set Add = New ChartSkil26.Line
If Key = "" Then
    Key = GenerateGUIDString
End If
Add.Initialise Me, Key, mGOSP, mCanvas, LocalStyle

mLines.Add Add, Key
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

Public Function Item(ByVal index As Variant) As ChartSkil26.Line
Attribute Item.VB_UserMemId = 0
Set Item = mLines(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
   Set NewEnum = mLines.[_NewEnum]
End Function

Public Sub RemoveObject(ByVal value As ChartSkil26.Line)
On Error GoTo Err
mLines.Remove value.Key
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
If mLines.Count <> 0 Then
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



