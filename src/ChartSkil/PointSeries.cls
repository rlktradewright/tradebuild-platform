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

'================================================================================
' Member variables
'================================================================================


Private Const ModuleName                As String = "DataPointSeries"

Private mChartRef           As WeakReference
Private mSeriesID           As Long
Private mSeriesUpdateNumber As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished           As Boolean

Private mDataPoints         As Collection

Private mCanvas             As Canvas
Private mLayer              As Long

Private mGOSP               As GraphObjServiceProvider

Private WithEvents mStyle   As DataPointStyle
Attribute mStyle.VB_VarHelpID = -1

Private mName               As String

Private mPrevDataPointRef   As WeakReference

' these fields override corresponding items in the Style
Private mLineThickness      As Long
Private mColor              As Long
Private mUpColor            As Long
Private mDownColor          As Long
Private mLineStyle          As LineStyles
Private mPointStyle         As PointStyles
Private mDisplayMode        As DataPointDisplayModes
Private mHistWidth          As Single
Private mIncludeInAutoscale As Boolean

Private mPropertyOverrideFlags As DataPointPropertyOverrideFlags

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mDataPoints = New Collection
Set mStyle = New DataPointStyle

End Sub

Private Sub Class_Terminate()
gLogger.Log LogLevelHighDetail, "DataPointSeries terminated"
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
Set mChartRef = Nothing
Set mPrevDataPointRef = Nothing
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
Set mChartRef = CreateWeakReference(pChart)
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
Dim dp As DataPoint
Set dp = mDataPoints.Item(index)
dp.Finish
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
Case "COLOR"
    repaintRegion False
Case "DISPLAYMODE"
    repaintRegion True
Case "DOWNCOLOR"
    repaintRegion False
Case "HistogramBarWidth"
    repaintRegion True
Case "INCLUDEINAUTOSCALE"
    repaintRegion True
Case "LINESTYLE"
    repaintRegion False
Case "LINETHICKNESS"
    repaintRegion True
Case "POINTSTYLE"
    repaintRegion False
Case "UPCOLOR"
    repaintRegion False
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

End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get Color() As Long
Color = mStyle.Color
If isPropertyOverrideFlagSet(DataPointIsSetColor) Then Color = mColor
End Property

Public Property Let Color(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointSeries" & ":" & "color", _
            "Value is not a valid color"
If mColor = value And isPropertyOverrideFlagSet(DataPointIsSetColor) Then Exit Property
mColor = value
setPropertyOverrideFlag DataPointIsSetColor
repaintRegion False
firePropertyChanged "Color"
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
firePropertyChanged "DisplayMode"
End Property

Public Property Get DownColor() As Long
DownColor = mStyle.DownColor
If isPropertyOverrideFlagSet(DataPointIsSetDownColor) Then DownColor = mDownColor
End Property

Public Property Let DownColor(ByVal value As Long)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointSeries" & ":" & "downColor", _
            "Value is not a valid color or -1"
If mDownColor = value And isPropertyOverrideFlagSet(DataPointIsSetDownColor) Then Exit Property
mDownColor = value
setPropertyOverrideFlag DataPointIsSetDownColor
repaintRegion False
firePropertyChanged "DownColor"
End Property

Public Property Get HistogramBarWidth() As Single
HistogramBarWidth = mStyle.HistogramBarWidth
If isPropertyOverrideFlagSet(DataPointIsSetHistWidth) Then HistogramBarWidth = mHistWidth
End Property

Public Property Let HistogramBarWidth(ByVal value As Single)
If value <= 0 Or value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointSeries" & ":" & "HistogramBarWidth", _
            "HistogramBarWidth must be greater than zero but not greater than 1"
If mHistWidth = value And isPropertyOverrideFlagSet(DataPointIsSetHistWidth) Then Exit Property
mHistWidth = value
setPropertyOverrideFlag DataPointIsSetHistWidth
repaintRegion True
firePropertyChanged "HistogramBarWidth"
End Property

Friend Property Get Id() As Long
Id = mSeriesID
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mStyle.IncludeInAutoscale
If isPropertyOverrideFlagSet(DataPointIsSetIncludeInAutoscale) Then IncludeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value And isPropertyOverrideFlagSet(DataPointIsSetIncludeInAutoscale) Then Exit Property
mIncludeInAutoscale = value
setPropertyOverrideFlag DataPointIsSetIncludeInAutoscale
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
If isPropertyOverrideFlagSet(DataPointIsSetLineStyle) Then LineStyle = mLineStyle
End Property

Public Property Let LineStyle(ByVal value As LineStyles)
If mLineStyle = value And isPropertyOverrideFlagSet(DataPointIsSetLineStyle) Then Exit Property
mLineStyle = value
setPropertyOverrideFlag DataPointIsSetLineStyle
repaintRegion False
firePropertyChanged "LineStyle"
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
firePropertyChanged "LineThickness"
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let Name(ByVal value As String)
mName = value
firePropertyChanged "Name"
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
firePropertyChanged "PointStyle"
End Property

Friend Property Get SeriesUpdateNumber() As Long
SeriesUpdateNumber = mSeriesUpdateNumber
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

firePropertyChanged "Style"
Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Style" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get UpColor() As Long
UpColor = mStyle.UpColor
If isPropertyOverrideFlagSet(DataPointIsSetUpColor) Then UpColor = mUpColor
End Property

Public Property Let UpColor(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointSeries" & ":" & "upColor", _
            "Value is not a valid color"
If mUpColor = value And isPropertyOverrideFlagSet(DataPointIsSetUpColor) Then Exit Property
mUpColor = value
setPropertyOverrideFlag DataPointIsSetUpColor
repaintRegion False
firePropertyChanged "UpColor"
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

If Not mPrevDataPointRef Is Nothing Then
    If Timestamp = prevTimestamp Then
        duplicateCount = duplicateCount + 1
    Else
        duplicateCount = 0
    End If
End If

Set Add = New DataPoint
If CDbl(Timestamp) <> 0 Then Add.X = chartObj.GetXFromTimestamp(Timestamp, False, duplicateCount)
If Key = "" Then
    Key = GenerateGUIDString
End If
Add.Initialise Me, Timestamp, Key, mGOSP, mCanvas, LocalStyle

If Not mPrevDataPointRef Is Nothing Then Add.PrevDataPoint = prevDataPointObj
mDataPoints.Add Add, Key
Set mPrevDataPointRef = CreateWeakReference(Add)
prevTimestamp = Timestamp
mGOSP.AddGraphicObject Add
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
mGOSP.RemoveObject value.handle

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

Private Function chartObj() As Chart
Set chartObj = mChartRef.Target
End Function

Private Sub clearPropertyOverrideFlag( _
                ByVal flag As DataPointPropertyOverrideFlags)
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
                ByVal flag As DataPointPropertyOverrideFlags) As Boolean
isPropertyOverrideFlagSet = gIsFlagSet(mPropertyOverrideFlags, flag)
End Function

Private Function prevDataPointObj() As DataPoint
Set prevDataPointObj = mPrevDataPointRef.Target
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



