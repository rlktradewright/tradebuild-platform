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

Event StyleChanged()

Event SelectionStateChanged()

'================================================================================
' Types
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private Const ProjectName                As String = "ChartSkil26"
Private Const ModuleName                As String = "DataPointSeries"

Private mController         As ChartController
Private mSeriesID           As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mFinished           As Boolean

Private mDataPoints         As Collection

Private mCanvas             As Canvas
Private mLayer              As Long

Private mGOSP               As GraphObjServiceProvider

Private mStyle              As DataPointStyle
Private mInhibitChangeEvent As Boolean

Private mName               As String

Private mPrevDataPoint      As DataPoint

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mDataPoints = New Collection
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
End Sub

Private Sub IGraphicObjectSeries_Initialise( _
                ByVal name As String, _
                ByVal Canvas As Canvas, _
                ByVal Layer As LayerNumbers, _
                ByVal controller As ChartController, _
                ByVal seriesId As Long, _
                ByVal GOSP As GraphObjServiceProvider)
mName = name
mSeriesID = seriesId
Set mGOSP = GOSP
mLayer = Layer
Set mCanvas = Canvas
Set mController = controller
If mStyle Is Nothing Then Set mStyle = controller.DefaultDataPointStyle.clone
End Sub

Private Property Get IGraphicObjectSeries_IsSelectable() As Boolean
IGraphicObjectSeries_IsSelectable = mIsSelectable
End Property

Private Function IGraphicObjectSeries_item(ByVal index As Long) As IGraphicObject
Set IGraphicObjectSeries_item = mDataPoints.Item(index)
End Function

Private Property Get IGraphicObjectSeries_name() As String
IGraphicObjectSeries_name = mName
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
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get Color() As Long
Color = mStyle.Color
End Property

Public Property Let Color(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointSeries" & ":" & "color", _
            "Value is not a valid color"
If mStyle.Color = value Then Exit Property
mStyle.Color = value
If mDataPoints.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get Count() As Long
Count = mDataPoints.Count
End Property

Public Property Get displayMode() As DataPointDisplayModes
displayMode = mStyle.displayMode
End Property

Public Property Let displayMode(ByVal value As DataPointDisplayModes)
If mStyle.displayMode = value Then Exit Property
mStyle.displayMode = value
If mDataPoints.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion True
End If
fireStyleChanged
End Property

Public Property Get downColor() As Long
downColor = mStyle.downColor
End Property

Public Property Let downColor(ByVal value As Long)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointSeries" & ":" & "downColor", _
            "Value is not a valid color or -1"
If mStyle.downColor = value Then Exit Property
mStyle.downColor = value
If mDataPoints.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get histBarWidth() As Single
histBarWidth = mStyle.histBarWidth
End Property

Public Property Let histBarWidth(ByVal value As Single)
If value <= 0 Or value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointSeries" & ":" & "histBarWidth", _
            "histBarWidth must be greater than zero but not greater than 1"
If mStyle.histBarWidth = value Then Exit Property
mStyle.histBarWidth = value
If mDataPoints.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Friend Property Get includeInAutoscale() As Boolean
includeInAutoscale = mStyle.includeInAutoscale
End Property

Public Property Let includeInAutoscale(ByVal value As Boolean)
If mStyle.includeInAutoscale = value Then Exit Property
mStyle.includeInAutoscale = value
If mDataPoints.Count <> 0 Then
    mGOSP.PaintRegion True
End If
fireStyleChanged
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

Public Property Get Layer() As Long
Layer = mLayer
End Property

Public Property Get linestyle() As LineStyles
linestyle = mStyle.linestyle
End Property

Public Property Let linestyle(ByVal value As LineStyles)
If mStyle.linestyle = value Then Exit Property
mStyle.linestyle = value
If mDataPoints.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get lineThickness() As Long
lineThickness = mStyle.lineThickness
End Property

Public Property Let lineThickness(ByVal value As Long)
If mStyle.lineThickness = value Then Exit Property
mStyle.lineThickness = value
If mDataPoints.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion True
End If
fireStyleChanged
End Property

Public Property Get name() As String
name = mName
End Property

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get pointStyle() As PointStyles
pointStyle = mStyle.pointStyle
End Property

Public Property Let pointStyle(ByVal value As PointStyles)
If mStyle.pointStyle = value Then Exit Property
mStyle.pointStyle = value
If mDataPoints.Count <> 0 Then
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

Public Property Get style() As DataPointStyle
Set style = mStyle.clone
End Property

Public Property Let style( _
                ByVal value As DataPointStyle)
Dim somethingChanged As Boolean
Dim recalcBoundaries As Boolean
Dim recalcRegion As Boolean

Dim failpoint As Long
On Error GoTo Err

If value Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "style", _
            "Style is null reference"
End If

mInhibitChangeEvent = True

If value.displayMode <> mStyle.displayMode Then
    somethingChanged = True
    mStyle.displayMode = value.displayMode
    recalcBoundaries = True
End If
If value.downColor <> mStyle.downColor Then
    somethingChanged = True
    mStyle.downColor = value.downColor
End If
If value.histBarWidth <> mStyle.histBarWidth Then
    somethingChanged = True
    mStyle.histBarWidth = value.histBarWidth
    recalcBoundaries = True
End If
If value.includeInAutoscale <> mStyle.includeInAutoscale Then
    somethingChanged = True
    recalcRegion = True
    mStyle.includeInAutoscale = value.includeInAutoscale
End If
If value.linestyle <> mStyle.linestyle Then
    somethingChanged = True
    mStyle.linestyle = value.linestyle
End If
If value.lineThickness <> mStyle.lineThickness Then
    somethingChanged = True
    mStyle.lineThickness = value.lineThickness
    recalcBoundaries = True
End If
If value.upColor <> mStyle.upColor Then
    somethingChanged = True
    mStyle.upColor = value.upColor
End If

mInhibitChangeEvent = False

If somethingChanged Then
    If recalcBoundaries Then If mDataPoints.Count <> 0 Then mGOSP.RecalcBoundary mSeriesID
    If mDataPoints.Count <> 0 Then mGOSP.PaintRegion recalcRegion
    fireStyleChanged
End If

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "style" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get upColor() As Long
upColor = mStyle.upColor
End Property

Public Property Let upColor(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "DataPointSeries" & ":" & "upColor", _
            "Value is not a valid color"
If mStyle.upColor = value Then Exit Property
mStyle.upColor = value
If mDataPoints.Count <> 0 Then
    mGOSP.RecalcBoundary mSeriesID
    mGOSP.PaintRegion False
End If
fireStyleChanged
End Property

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                ByVal timestamp As Date, _
                Optional ByVal Key As String, _
                Optional ByVal LocalStyle As DataPointStyle, _
                Optional ByVal IsSelectable As Boolean) As DataPoint
Static duplicateCount As Long       ' counts the number of successive adds with
                                    ' the same timestamp
Static prevTimestamp As Date

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil26.DatPointSeries::add", _
                            "DataPoint series is finished"

If Not mPrevDataPoint Is Nothing Then
    If timestamp = prevTimestamp Then
        duplicateCount = duplicateCount + 1
    Else
        duplicateCount = 0
    End If
End If

Set Add = New DataPoint
If CDbl(timestamp) <> 0 Then Add.x = gCalculateX(timestamp, mController, False, duplicateCount)
If Key = "" Then
    Key = GenerateGUIDString
End If
Add.initialise mSeriesID, timestamp, Key, mGOSP, mLayer, mCanvas, mStyle, LocalStyle, IsSelectable

Add.prevDataPoint = mPrevDataPoint
mDataPoints.Add Add, Key
Set mPrevDataPoint = Add
prevTimestamp = timestamp
mGOSP.AddGraphicObject Add, mSeriesID
End Function

Friend Sub Finish()
mFinished = True
Set mCanvas = Nothing
End Sub

Friend Sub initialise( _
                ByVal name As String, _
                ByVal style As DataPointStyle)
mName = name
Set mStyle = style.clone    ' style cannot be shared between DataPointSeries
End Sub

Public Function Item(ByVal index As Variant) As DataPoint
Attribute Item.VB_UserMemId = 0
Set Item = mDataPoints(index)
End Function

Friend Sub RemoveAll()
StartTask Me, PriorityLow
End Sub

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

Private Sub fireStyleChanged()
If Not mInhibitChangeEvent Then RaiseEvent StyleChanged
End Sub


