VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataPointSeries"
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

'================================================================================
' Types
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Member variables
'================================================================================


Private Const ModuleName                As String = "DataPointSeries"

Private mChart              As Chart
Private mSeriesID           As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished           As Boolean

Private mDataPoints         As Collection

Private mCanvas             As Canvas
Private mLayer              As Long

Private mGOSP               As GraphObjServiceProvider

Private WithEvents mStyle   As DataPointStyle
Attribute mStyle.VB_VarHelpID = -1

Private mIncludeInAutoscale As Boolean

Private mName               As String

Private mPrevDataPoint      As DataPoint

' these fields override corresponding items in the Style
Private mLineThickness      As Long
Private mColor              As Long
Private mUpColor            As Long
Private mDownColor          As Long
Private mLineStyle          As LineStyles
Private mPointStyle         As PointStyles
Private mDisplayMode        As DataPointDisplayModes
Private mHistBarWidth       As Single

Private mPropertyOverrideFlags As DataPointPropertyOverrideFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mDataPoints = New Collection
Set mStyle = New DataPointStyle

End Sub

Private Sub Class_Terminate()
Debug.Print "DataPointSeries terminated"
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
IGraphicObjectSeries_count = mDataPoints.Count
End Property

Private Sub IGraphicObjectSeries_finish()
mFinished = True
Set mChart = Nothing
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
Set mChart = pChart
End Sub

Private Property Get IGraphicObjectSeries_IsSelectable() As Boolean
IGraphicObjectSeries_IsSelectable = mIsSelectable
End Property

Private Function IGraphicObjectSeries_item(ByVal index As Long) As IGraphicObject
Set IGraphicObjectSeries_item = mDataPoints.Item(index)
End Function

Private Property Get IGraphicObjectSeries_Name() As String
IGraphicObjectSeries_Name = mName
End Property

Public Sub IGraphicObjectSeries_Remove(ByVal index As Long)
mDataPoints.Remove index
End Sub

Private Sub IGraphicObjectSeries_RemoveObject(ByVal value As IGraphicObject)
Dim dp As ChartSkil26.DataPoint

On Error Resume Next
Set dp = value

mDataPoints.Remove dp.Key
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
Dim s As DataPointStyle
Dim failpoint As Long
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case UCase$("Color")
    repaintRegion False
Case UCase$("displayMode")
    repaintRegion True
Case UCase$("downColor")
    repaintRegion False
Case UCase$("histBarWidth")
    repaintRegion True
Case UCase$("LineStyle")
    repaintRegion False
Case UCase$("lineThickness")
    repaintRegion True
Case UCase$("pointStyle")
    repaintRegion False
Case UCase$("upColor")
    repaintRegion False
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

Public Property Get Color() As OLE_COLOR
Color = mStyle.Color
If isPropertyOverrideFlagSet(DataPointIsSetColor) Then Color = mColor
End Property

Public Property Let Color(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointSeries" & ":" & "color", _
            "Value is not a valid color"
If mColor = value And isPropertyOverrideFlagSet(DataPointIsSetColor) Then Exit Property
mColor = value
setPropertyOverrideFlag DataPointIsSetColor
repaintRegion False
End Property

Public Property Get Count() As Long
Count = mDataPoints.Count
End Property

Public Property Get DisplayMode() As DataPointDisplayModes
DisplayMode = mStyle.DisplayMode
If isPropertyOverrideFlagSet(DataPointIsSetDisplayMode) Then DisplayMode = mDisplayMode
End Property

Public Property Let DisplayMode(ByVal value As DataPointDisplayModes)
If mDisplayMode = value And isPropertyOverrideFlagSet(DataPointIsSetDisplayMode) Then Exit Property
mDisplayMode = value
setPropertyOverrideFlag DataPointIsSetDisplayMode
repaintRegion True
End Property

Public Property Get DownColor() As OLE_COLOR
DownColor = mStyle.DownColor
If isPropertyOverrideFlagSet(DataPointIsSetDownColor) Then DownColor = mDownColor
End Property

Public Property Let DownColor(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointSeries" & ":" & "downColor", _
            "Value is not a valid color or -1"
If mDownColor = value And isPropertyOverrideFlagSet(DataPointIsSetDownColor) Then Exit Property
mDownColor = value
setPropertyOverrideFlag DataPointIsSetDownColor
repaintRegion False
End Property

Public Property Get HistBarWidth() As Single
HistBarWidth = mStyle.HistBarWidth
If isPropertyOverrideFlagSet(DataPointIsSetHistBarWidth) Then HistBarWidth = mHistBarWidth
End Property

Public Property Let HistBarWidth(ByVal value As Single)
If value <= 0 Or value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointSeries" & ":" & "histBarWidth", _
            "histBarWidth must be greater than zero but not greater than 1"
If mHistBarWidth = value And isPropertyOverrideFlagSet(DataPointIsSetHistBarWidth) Then Exit Property
mHistBarWidth = value
setPropertyOverrideFlag DataPointIsSetHistBarWidth
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
If isPropertyOverrideFlagSet(DataPointIsSetLineStyle) Then LineStyle = mLineStyle
End Property

Public Property Let LineStyle(ByVal value As LineStyles)
If mLineStyle = value And isPropertyOverrideFlagSet(DataPointIsSetLineStyle) Then Exit Property
mLineStyle = value
setPropertyOverrideFlag DataPointIsSetLineStyle
repaintRegion False
End Property

Public Property Get LineThickness() As Long
LineThickness = mStyle.LineThickness
If isPropertyOverrideFlagSet(DataPointIsSetLineThickness) Then LineThickness = mLineThickness
End Property

Public Property Let LineThickness(ByVal value As Long)
If mLineThickness = value And isPropertyOverrideFlagSet(DataPointIsSetLineThickness) Then Exit Property
mLineThickness = value
setPropertyOverrideFlag DataPointIsSetLineThickness
repaintRegion True
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let Name(ByVal value As String)
mName = value
End Property

Public Property Get PointStyle() As PointStyles
PointStyle = mStyle.PointStyle
If isPropertyOverrideFlagSet(DataPointIsSetPointStyle) Then PointStyle = mPointStyle
End Property

Public Property Let PointStyle(ByVal value As PointStyles)
If mPointStyle = value And isPropertyOverrideFlagSet(DataPointIsSetPointStyle) Then Exit Property
mPointStyle = value
setPropertyOverrideFlag DataPointIsSetPointStyle
repaintRegion False
End Property

Public Property Get Style() As DataPointStyle
Set Style = mStyle
End Property

Public Property Let Style( _
                ByVal value As DataPointStyle)
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

Public Property Get UpColor() As OLE_COLOR
UpColor = mStyle.UpColor
If isPropertyOverrideFlagSet(DataPointIsSetUpColor) Then UpColor = mUpColor
End Property

Public Property Let UpColor(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointSeries" & ":" & "upColor", _
            "Value is not a valid color"
If mUpColor = value And isPropertyOverrideFlagSet(DataPointIsSetUpColor) Then Exit Property
mUpColor = value
setPropertyOverrideFlag DataPointIsSetUpColor
repaintRegion False
End Property

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                ByVal Timestamp As Date, _
                Optional ByVal Key As String, _
                Optional ByVal LocalStyle As DataPointStyle) As DataPoint
Static duplicateCount As Long       ' counts the number of successive adds with
                                    ' the same Timestamp
Static prevTimestamp As Date

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil26.DatPointSeries::Add", _
                            "DataPoint series is finished"

If Not mPrevDataPoint Is Nothing Then
    If Timestamp = prevTimestamp Then
        duplicateCount = duplicateCount + 1
    Else
        duplicateCount = 0
    End If
End If

Set Add = New DataPoint
If CDbl(Timestamp) <> 0 Then Add.X = mChart.GetXFromTimestamp(Timestamp, False, duplicateCount)
If Key = "" Then
    Key = GenerateGUIDString
End If
Add.Initialise Me, Timestamp, Key, mGOSP, mCanvas, LocalStyle
Add.IncludeInAutoscale = mIncludeInAutoscale

Add.PrevDataPoint = mPrevDataPoint
mDataPoints.Add Add, Key
Set mPrevDataPoint = Add
prevTimestamp = Timestamp
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

Public Function Item(ByVal index As Variant) As DataPoint
Attribute Item.VB_UserMemId = 0
Set Item = mDataPoints(index)
End Function

Public Sub RemoveObject(ByVal value As ChartSkil26.DataPoint)
On Error GoTo Err
mDataPoints.Remove value.Key
mGOSP.RemoveObject value.Handle

Exit Sub

Err:
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mDataPoints.[_NewEnum]
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As DataPointPropertyOverrideFlags)
mPropertyOverrideFlags = gClearFlag(mPropertyOverrideFlags, flag)
End Sub

Private Function isPropertyOverrideFlagSet( _
                ByVal flag As DataPointPropertyOverrideFlags) As Boolean
isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)
End Function

Private Sub repaintRegion( _
                ByVal recalcBoundaries As Boolean)
If mDataPoints.Count <> 0 Then
    If recalcBoundaries Then
        mGOSP.RecalcBoundary mSeriesID
        mGOSP.PaintRegion True
    Else
        mGOSP.PaintRegion False
    End If
End If
End Sub

Private Sub setPropertyOverrideFlag( _
                ByVal flag As DataPointPropertyOverrideFlags)
mPropertyOverrideFlags = gSetFlag(mPropertyOverrideFlags, flag)
End Sub



