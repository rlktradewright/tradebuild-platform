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

Event PropertyChanged(ev As PropertyChangedEvent)

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


Private Const ModuleName                As String = "LineSeries"

'================================================================================
' Member variables
'================================================================================

Private mName As String
Private mSeriesID As Long
Private mSeriesUpdateNumber As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished As Boolean

Private mLines As Collection

'Private mSurface As PictureBox
Private mViewport As Viewport
Private mLayer As Long

Private mGOSP As GraphObjServiceProvider

Private WithEvents mStyle As LineStyle
Attribute mStyle.VB_VarHelpID = -1

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
Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

Private mPropertyOverrideFlags As LinePropertyOverrideFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mLines = New Collection
Set mStyle = New LineStyle

End Sub

Private Sub Class_Terminate()
gLogger.Log LogLevelHighDetail, "LineSeries terminated"
Debug.Print "LineSeries terminated"
End Sub

'================================================================================
' IGraphicObjectSeries Interface Members
'================================================================================

Private Property Let IGraphicObjectSeries_Canvas(ByVal value As Canvas)
Set mViewport = value
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
Set mViewport = Nothing
End Sub

Private Sub IGraphicObjectSeries_Initialise( _
                ByVal pName As String, _
                ByVal pviewport As Viewport, _
                ByVal pLayer As LayerNumbers, _
                ByVal pChart As Chart, _
                ByVal pSeriesID As Long, _
                ByVal pGOSP As GraphObjServiceProvider)
mName = pName
mSeriesID = pSeriesID
Set mGOSP = pGOSP
mLayer = pLayer
Set mViewport = pviewport
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
    firePropertyChanged "IsSelected"
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
Case "ARROWENDCOLOR"
    repaintRegion False
Case "ARROWENDFILLCOLOR"
    repaintRegion False
Case "ARROWENDFILLSTYLE"
    repaintRegion False
Case "ARROWENDLENGTH"
    repaintRegion True
Case "ARROWENDSTYLE"
    repaintRegion True
Case "ARROWENDWIDTH"
    repaintRegion True
Case "ARROWSTARTCOLOR"
    repaintRegion False
Case "ARROWSTARTFILLCOLOR"
    repaintRegion False
Case "ARROWSTARTFILLSTYLE"
    repaintRegion False
Case "ARROWSTARTLENGTH"
    repaintRegion True
Case "ARROWSTARTSTYLE"
    repaintRegion True
Case "ARROWSTARTWIDTH"
    repaintRegion True
Case "COLOR"
    repaintRegion False
Case "EXTENDAFTER"
    repaintRegion True
Case "EXTENDBEFORE"
    repaintRegion True
Case "LINESTYLE"
    repaintRegion False
Case "THICKNESS"
    repaintRegion True
Case "EXTENDED"
    repaintRegion True
Case "FIXEDX"
    repaintRegion True
Case "FIXEDY"
    repaintRegion True
Case "INCLUDEINAUTOSCALE"
    repaintRegion True
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & "mStyle_PropertyChanged", _
            "Unhandled property change"
End Select

firePropertyChanged ev.PropertyName

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

Public Property Get ArrowEndColor() As Long
ArrowEndColor = mStyle.ArrowEndColor
If isPropertyOverrideFlagSet(LineIsSetArrowEndColor) Then ArrowEndColor = mArrowEndColor
End Property

Public Property Let ArrowEndColor(ByVal value As Long)
If mArrowEndColor = value And isPropertyOverrideFlagSet(LineIsSetArrowEndColor) Then Exit Property
mArrowEndColor = value
setPropertyOverrideFlag LineIsSetArrowEndColor
repaintRegion False
firePropertyChanged "ArrowEndColor"
End Property

Public Property Get ArrowEndFillColor() As Long
ArrowEndFillColor = mStyle.ArrowEndFillColor
If isPropertyOverrideFlagSet(LineIsSetArrowEndFillColor) Then ArrowEndFillColor = mArrowEndFillColor
End Property

Public Property Let ArrowEndFillColor(ByVal value As Long)
If mArrowEndFillColor = value And isPropertyOverrideFlagSet(LineIsSetArrowEndFillColor) Then Exit Property
mArrowEndFillColor = value
setPropertyOverrideFlag LineIsSetArrowEndFillColor
repaintRegion False
firePropertyChanged "ArrowEndFillColor"
End Property

Public Property Get ArrowEndFillStyle() As FillStyles
ArrowEndFillStyle = mStyle.ArrowEndFillStyle
If isPropertyOverrideFlagSet(LineIsSetArrowEndFillStyle) Then ArrowEndFillStyle = mArrowEndFillStyle
End Property

Public Property Let ArrowEndFillStyle(ByVal value As FillStyles)
If mArrowEndFillStyle = value And isPropertyOverrideFlagSet(LineIsSetArrowEndFillStyle) Then Exit Property
mArrowEndFillStyle = value
setPropertyOverrideFlag LineIsSetArrowEndFillStyle
repaintRegion False
firePropertyChanged "ArrowEndFillStyle"
End Property

Public Property Get ArrowEndLength() As Long
ArrowEndLength = mStyle.ArrowEndLength
If isPropertyOverrideFlagSet(LineIsSetArrowEndLength) Then ArrowEndLength = mArrowEndLength
End Property

Public Property Let ArrowEndLength(ByVal value As Long)
If mArrowEndLength = value And isPropertyOverrideFlagSet(LineIsSetArrowEndLength) Then Exit Property
mArrowEndLength = value
setPropertyOverrideFlag LineIsSetArrowEndLength
repaintRegion True
firePropertyChanged "ArrowEndLength"
End Property

Public Property Get ArrowEndStyle() As ArrowStyles
ArrowEndStyle = mStyle.ArrowEndStyle
If isPropertyOverrideFlagSet(LineIsSetArrowEndStyle) Then ArrowEndStyle = mArrowEndStyle
End Property

Public Property Let ArrowEndStyle(ByVal value As ArrowStyles)
If mArrowEndStyle = value And isPropertyOverrideFlagSet(LineIsSetArrowEndStyle) Then Exit Property
mArrowEndStyle = value
setPropertyOverrideFlag LineIsSetArrowEndStyle
repaintRegion True
firePropertyChanged "ArrowEndStyle"
End Property

Public Property Get ArrowEndWidth() As Long
ArrowEndWidth = mStyle.ArrowEndWidth
If isPropertyOverrideFlagSet(LineIsSetArrowEndWidth) Then ArrowEndWidth = mArrowEndWidth
End Property

Public Property Let ArrowEndWidth(ByVal value As Long)
If mArrowEndWidth = value And isPropertyOverrideFlagSet(LineIsSetArrowEndWidth) Then Exit Property
mArrowEndWidth = value
setPropertyOverrideFlag LineIsSetArrowEndWidth
repaintRegion True
firePropertyChanged "ArrowEndWidth"
End Property

Public Property Get ArrowStartColor() As Long
ArrowStartColor = mStyle.ArrowStartColor
If isPropertyOverrideFlagSet(LineIsSetArrowStartColor) Then ArrowStartColor = mArrowStartColor
End Property

Public Property Let ArrowStartColor(ByVal value As Long)
If mArrowStartColor = value And isPropertyOverrideFlagSet(LineIsSetArrowStartColor) Then Exit Property
mArrowStartColor = value
setPropertyOverrideFlag LineIsSetArrowStartColor
repaintRegion False
firePropertyChanged "ArrowStartColor"
End Property

Public Property Get ArrowStartFillColor() As Long
ArrowStartFillColor = mStyle.ArrowStartFillColor
If isPropertyOverrideFlagSet(LineIsSetArrowStartFillColor) Then ArrowStartFillColor = mArrowStartFillColor
End Property

Public Property Let ArrowStartFillColor(ByVal value As Long)
If mArrowStartFillColor = value And isPropertyOverrideFlagSet(LineIsSetArrowStartFillColor) Then Exit Property
mArrowStartFillColor = value
setPropertyOverrideFlag LineIsSetArrowStartFillColor
repaintRegion False
firePropertyChanged "ArrowStartFillColor"
End Property

Public Property Get ArrowStartFillStyle() As FillStyles
ArrowStartFillStyle = mStyle.ArrowStartFillStyle
If isPropertyOverrideFlagSet(LineIsSetArrowStartFillStyle) Then ArrowStartFillStyle = mArrowStartfillStyle
End Property

Public Property Let ArrowStartFillStyle(ByVal value As FillStyles)
If mArrowStartfillStyle = value And isPropertyOverrideFlagSet(LineIsSetArrowStartFillStyle) Then Exit Property
mArrowStartfillStyle = value
setPropertyOverrideFlag LineIsSetArrowStartFillStyle
repaintRegion False
firePropertyChanged "ArrowStartFillStyle"
End Property

Public Property Get ArrowStartLength() As Long
ArrowStartLength = mStyle.ArrowStartLength
If isPropertyOverrideFlagSet(LineIsSetArrowStartLength) Then ArrowStartLength = mArrowStartLength
End Property

Public Property Let ArrowStartLength(ByVal value As Long)
If mArrowStartLength = value And isPropertyOverrideFlagSet(LineIsSetArrowStartLength) Then Exit Property
mArrowStartLength = value
setPropertyOverrideFlag LineIsSetArrowStartLength
repaintRegion True
firePropertyChanged "ArrowStartLength"
End Property

Public Property Get ArrowStartStyle() As ArrowStyles
ArrowStartStyle = mStyle.ArrowStartStyle
If isPropertyOverrideFlagSet(LineIsSetArrowStartStyle) Then ArrowStartStyle = mArrowStartStyle
End Property

Public Property Let ArrowStartStyle(ByVal value As ArrowStyles)
If mArrowStartStyle = value And isPropertyOverrideFlagSet(LineIsSetArrowStartStyle) Then Exit Property
mArrowStartStyle = value
setPropertyOverrideFlag LineIsSetArrowStartStyle
repaintRegion True
firePropertyChanged "ArrowStartStyle"
End Property

Public Property Get ArrowStartWidth() As Long
ArrowStartWidth = mStyle.ArrowStartWidth
If isPropertyOverrideFlagSet(LineIsSetArrowStartWidth) Then ArrowStartWidth = mArrowStartWidth
End Property

Public Property Let ArrowStartWidth(ByVal value As Long)
If mArrowStartWidth = value And isPropertyOverrideFlagSet(LineIsSetArrowStartWidth) Then Exit Property
mArrowStartWidth = value
setPropertyOverrideFlag LineIsSetArrowStartWidth
repaintRegion True
firePropertyChanged "ArrowStartWidth"
End Property

Friend Property Let Canvas(ByVal value As Canvas)
Set mViewport = value
End Property

Public Property Get Color() As Long
Color = mStyle.Color
If isPropertyOverrideFlagSet(LineIsSetColor) Then Color = mColor
End Property

Public Property Let Color(ByVal value As Long)
If mColor = value And isPropertyOverrideFlagSet(LineIsSetColor) Then Exit Property
mColor = value
setPropertyOverrideFlag LineIsSetColor
repaintRegion False
firePropertyChanged "Color"
End Property

Public Property Get Count() As Long
Count = mLines.Count
End Property

Public Property Get ExtendAfter() As Boolean
ExtendAfter = mStyle.ExtendAfter
If isPropertyOverrideFlagSet(LineIsSetExtendAfter) Then ExtendAfter = mExtendAfter
End Property

Public Property Let ExtendAfter(ByVal value As Boolean)
If mExtendAfter = value And isPropertyOverrideFlagSet(LineIsSetExtendAfter) Then Exit Property
mExtendAfter = value
setPropertyOverrideFlag LineIsSetExtendAfter
repaintRegion True
firePropertyChanged "ExtendAfter"
End Property

Public Property Get ExtendBefore() As Boolean
ExtendBefore = mStyle.ExtendBefore
If isPropertyOverrideFlagSet(LineIsSetExtendBefore) Then ExtendBefore = mExtendBefore
End Property

Public Property Let ExtendBefore(ByVal value As Boolean)
If mExtendBefore = value And isPropertyOverrideFlagSet(LineIsSetExtendBefore) Then Exit Property
mExtendBefore = value
setPropertyOverrideFlag LineIsSetExtendBefore
repaintRegion True
firePropertyChanged "ExtendBefore"
End Property

Public Property Get Extended() As Boolean
Extended = mStyle.Extended
If isPropertyOverrideFlagSet(LineIsSetExtended) Then Extended = mExtended
End Property

Public Property Let Extended(ByVal value As Boolean)
If mExtended = value And isPropertyOverrideFlagSet(LineIsSetExtended) Then Exit Property
mExtended = value
setPropertyOverrideFlag LineIsSetExtended
repaintRegion True
firePropertyChanged "Extended"
End Property

Public Property Get FixedX() As Boolean
FixedX = mStyle.FixedX
If isPropertyOverrideFlagSet(LineIsSetFixedX) Then FixedX = mFixedX
End Property

Public Property Let FixedX(ByVal value As Boolean)
If mFixedX = value And isPropertyOverrideFlagSet(LineIsSetFixedX) Then Exit Property
mFixedX = value
setPropertyOverrideFlag LineIsSetFixedX
repaintRegion True
firePropertyChanged "FixedX"
End Property

Public Property Get FixedY() As Boolean
FixedY = mStyle.FixedY
If isPropertyOverrideFlagSet(LineIsSetFixedY) Then FixedY = mFixedY
End Property

Public Property Let FixedY(ByVal value As Boolean)
If mFixedY = value And isPropertyOverrideFlagSet(LineIsSetFixedY) Then Exit Property
mFixedY = value
setPropertyOverrideFlag LineIsSetFixedY
repaintRegion True
firePropertyChanged "FixedY"
End Property

Friend Property Get Id() As Long
Id = mSeriesID
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mStyle.IncludeInAutoscale
If isPropertyOverrideFlagSet(LineIsSetIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value And isPropertyOverrideFlagSet(LineIsSetIncludeInAutoscale) Then Exit Property
mIncludeInAutoscale = value
setPropertyOverrideFlag LineIsSetIncludeInAutoscale
repaintRegion True
firePropertyChanged "IncludeInAutoscale"
End Property

Public Property Let IsSelectable(ByVal value As Boolean)
mIsSelectable = value
firePropertyChanged "IsSelectable"
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
firePropertyChanged "Layer"
End Property

Public Property Get Layer() As LayerNumbers
Layer = mLayer
End Property

Public Property Get LineStyle() As LineStyles
LineStyle = mStyle.LineStyle
If isPropertyOverrideFlagSet(LineIsSetLineStyle) Then LineStyle = mLineStyle
End Property

Public Property Let LineStyle(ByVal value As LineStyles)
If mLineStyle = value And isPropertyOverrideFlagSet(LineIsSetLineStyle) Then Exit Property
mLineStyle = value
setPropertyOverrideFlag LineIsSetLineStyle
repaintRegion False
firePropertyChanged "LineStyle"
End Property

Friend Property Get SeriesUpdateNumber() As Long
SeriesUpdateNumber = mSeriesUpdateNumber
End Property

Public Property Get Style() As LineStyle
Set Style = mStyle
End Property

Public Property Let Style(ByVal value As LineStyle)

Dim failpoint As Long
On Error GoTo Err

If value Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "Style", _
            "Style is null reference"
End If

Set mStyle = value
repaintRegion True

firePropertyChanged "Style"
Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Style" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get Thickness() As Long
Thickness = mStyle.Thickness
If isPropertyOverrideFlagSet(LineIsSetThickness) Then Thickness = mThickness
End Property

Public Property Let Thickness(ByVal value As Long)
If mThickness = value And isPropertyOverrideFlagSet(LineIsSetThickness) Then Exit Property
mThickness = value
setPropertyOverrideFlag LineIsSetThickness
repaintRegion True
firePropertyChanged "Thickness"
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
Add.Initialise Me, Key, mGOSP, mViewport, LocalStyle

mLines.Add Add, Key
mGOSP.AddGraphicObject Add
End Function

Public Sub ClearOverrides()
mPropertyOverrideFlags = 0
repaintRegion True
End Sub

Friend Sub Finish()
mFinished = True
Set mViewport = Nothing
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
mGOSP.RemoveObject value.handle

Exit Sub

Err:
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As LinePropertyOverrideFlags)
mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)
End Sub

Private Sub firePropertyChanged( _
                ByVal Name As String)
Dim ev As PropertyChangedEvent

Set ev.Source = Me
ev.PropertyName = Name
RaiseEvent PropertyChanged(ev)
End Sub

Private Function isPropertyOverrideFlagSet( _
                ByVal flag As LinePropertyOverrideFlags) As Boolean
isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)
End Function

Private Sub repaintRegion( _
                ByVal recalcBoundaries As Boolean)
If mLines.Count <> 0 Then
    If recalcBoundaries Then
        mSeriesUpdateNumber = mSeriesUpdateNumber + 1
        mGOSP.RecalcBoundary mSeriesID
        mGOSP.PaintRegion True
    Else
        mGOSP.PaintRegion False
    End If
End If
End Sub

Private Sub setPropertyOverrideFlag( _
                ByVal flag As LinePropertyOverrideFlags)
mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)
End Sub



