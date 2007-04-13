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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements ChartSkilTypes.IGraphicObjectSeries

'================================================================================
' Events
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

Private mController         As ChartController
Private mSeriesId           As Long

Private mFinished           As Boolean

Private mDataPoints         As Collection

Private mCanvas             As canvas
Private mLayer              As Long

Private mEventProxy         As ChartRegionEventProxy

Private mStyle              As DataPointStyle

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
mSeriesId = GenerateID
End Sub

'================================================================================
' IGraphicObjectSeries Interface Members
'================================================================================

Private Property Get IGraphicObjectSeries_count() As Long
IGraphicObjectSeries_count = mDataPoints.count
End Property

Private Sub IGraphicObjectSeries_finish()
mFinished = True
End Sub

Private Function IGraphicObjectSeries_item(ByVal index As Long) As ChartSkilTypes.IGraphicObject
Set IGraphicObjectSeries_item = mDataPoints.item(index)
End Function

Private Function IGraphicObjectSeries_NewEnum() As Object
Set IGraphicObjectSeries_NewEnum = NewEnum
End Function

Private Sub IGraphicObjectSeries_remove(ByVal index As Long)
mDataPoints.remove index
End Sub

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
            "ChartSkil25" & "." & "DataPointSeries" & ":" & "color", _
            "Value is not a valid color"
If mStyle.Color = value Then Exit Property
mStyle.Color = value
If mDataPoints.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get count() As Long
count = mDataPoints.count
End Property

Public Property Get displayMode() As DataPointDisplayModes
displayMode = mStyle.displayMode
End Property

Public Property Let displayMode(ByVal value As DataPointDisplayModes)
If mStyle.displayMode = value Then Exit Property
mStyle.displayMode = value
If mDataPoints.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion True
End If
End Property

Public Property Get downColor() As Long
downColor = mStyle.downColor
End Property

Public Property Let downColor(ByVal value As Long)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "DataPointSeries" & ":" & "downColor", _
            "Value is not a valid color or -1"
If mStyle.downColor = value Then Exit Property
mStyle.downColor = value
If mDataPoints.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get histBarWidth() As Single
histBarWidth = mStyle.histBarWidth
End Property

Public Property Let histBarWidth(ByVal value As Single)
If value <= 0 Or value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "DataPointSeries" & ":" & "histBarWidth", _
            "histBarWidth must be greater than zero but not greater than 1"
If mStyle.histBarWidth = value Then Exit Property
mStyle.histBarWidth = value
If mDataPoints.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Friend Property Get includeInAutoscale() As Boolean
includeInAutoscale = mStyle.includeInAutoscale
End Property

Public Property Let includeInAutoscale(ByVal value As Boolean)
If mStyle.includeInAutoscale = value Then Exit Property
mStyle.includeInAutoscale = value
If mDataPoints.count <> 0 Then
    mEventProxy.paintRegion True
End If
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Get linestyle() As LineStyles
linestyle = mStyle.linestyle
End Property

Public Property Let linestyle(ByVal value As LineStyles)
If mStyle.linestyle = value Then Exit Property
mStyle.linestyle = value
If mDataPoints.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get lineThickness() As Long
lineThickness = mStyle.lineThickness
End Property

Public Property Let lineThickness(ByVal value As Long)
If mStyle.lineThickness = value Then Exit Property
mStyle.lineThickness = value
If mDataPoints.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion True
End If
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
If mDataPoints.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get style() As DataPointStyle
Set style = mStyle.clone
End Property

Public Property Let style( _
                ByVal value As DataPointStyle)
Dim somethingChanged As Boolean
Dim recalcBoundaries As Boolean
Dim recalcRegion As Boolean

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

If somethingChanged Then
    If recalcBoundaries Then If mDataPoints.count <> 0 Then mEventProxy.recalcBoundary mSeriesId
    If mDataPoints.count <> 0 Then mEventProxy.paintRegion recalcRegion
End If
End Property

Public Property Get upColor() As Long
upColor = mStyle.upColor
End Property

Public Property Let upColor(ByVal value As Long)
If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil25" & "." & "DataPointSeries" & ":" & "upColor", _
            "Value is not a valid color"
If mStyle.upColor = value Then Exit Property
mStyle.upColor = value
If mDataPoints.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                ByVal timestamp As Date, _
                Optional ByVal key As String, _
                Optional ByVal localStyle As DataPointStyle) As DataPoint
Static duplicateCount As Long       ' counts the number of successive adds with
                                    ' the same timestamp
Static prevTimestamp As Date

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil25.DatPointSeries::add", _
                            "DataPoint series is finished"

If Not mPrevDataPoint Is Nothing Then
    If timestamp = prevTimestamp Then
        duplicateCount = duplicateCount + 1
    Else
        duplicateCount = 0
    End If
End If

Set add = New DataPoint
add.x = gCalculateX(timestamp, mController, False, duplicateCount)
If key = "" Then
    key = GenerateGUIDString
End If
add.initialise key, mEventProxy, mLayer, mCanvas, mStyle, localStyle

add.prevDataPoint = mPrevDataPoint
mDataPoints.add add, key
Set mPrevDataPoint = add
prevTimestamp = timestamp
mEventProxy.addGraphicObject add, mSeriesId
End Function

Friend Sub finish()
mFinished = True
Set mCanvas = Nothing
End Sub

Friend Sub initialise( _
                ByVal name As String, _
                ByVal controller As ChartController, _
                ByVal eventProxy As ChartRegionEventProxy, _
                ByVal layer As Long, _
                ByVal canvas As canvas, _
                ByVal style As DataPointStyle)
mName = name
Set mController = controller
Set mEventProxy = eventProxy
mLayer = layer
Set mCanvas = canvas
Set mStyle = style.clone    ' style cannot be shared between DataPointSeries
End Sub

Public Function item(ByVal index As Variant) As DataPoint
Attribute item.VB_UserMemId = 0
Set item = mDataPoints(index)
End Function

Friend Sub removeAll()
Set mDataPoints = New Collection
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mDataPoints.[_NewEnum]
End Function

'================================================================================
' Helper Functions
'================================================================================



