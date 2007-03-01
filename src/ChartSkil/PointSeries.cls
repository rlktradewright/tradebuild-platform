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
' Member variables and constants
'================================================================================

Private mController As ChartController

Private mSeriesId As Long

Private mFinished As Boolean

Private Const DefaultHistBarWidth As Single = 0.5
Private mDataPoints As Collection

'Private mSurface As PictureBox
Private mCanvas As canvas
Private mLayer As Long

Private mEventProxy As ChartRegionEventProxy

Private mName As String
Private mLineThickness As Long
Private mLineColor As Long
Private mLineStyle As LineStyles
Private mDisplayMode As DataPointDisplayModes
Private mIncludeInAutoscale As Boolean

Private mHistBarWidth As Single

Private mPrevDataPoint As DataPoint

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mDataPoints = New Collection
mLineStyle = LineStyles.LineSolid
mHistBarWidth = DefaultHistBarWidth
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

Friend Property Let canvas(ByVal value As canvas)
Set mCanvas = value
End Property

Friend Property Let controller( _
                ByVal value As ChartController)
Set mController = value
End Property

Public Property Get count() As Long
count = mDataPoints.count
End Property

Public Property Get displayMode() As DataPointDisplayModes
displayMode = mDisplayMode
End Property

Public Property Let displayMode(ByVal value As DataPointDisplayModes)
Dim dp As DataPoint
For Each dp In mDataPoints
    dp.displayMode = value
Next
mDisplayMode = value
End Property

Friend Property Let eventProxy(ByVal value As ChartRegionEventProxy)
Set mEventProxy = value
End Property

Public Property Get histBarWidth() As Single
histBarWidth = mHistBarWidth
End Property

Public Property Let histBarWidth(ByVal value As Single)
mHistBarWidth = value
End Property

Friend Property Get includeInAutoscale() As Boolean
includeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let includeInAutoscale(ByVal newValue As Boolean)
mIncludeInAutoscale = newValue
End Property

Friend Property Let layer(ByVal value As Long)
mLayer = value
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Get lineColor() As Long
lineColor = mLineColor
End Property

Public Property Let lineColor(ByVal value As Long)
Dim dp As DataPoint
For Each dp In mDataPoints
    dp.lineColor = value
Next
mLineColor = value
End Property

Public Property Get lineStyle() As LineStyles
lineStyle = mLineStyle
End Property

Public Property Let lineStyle(ByVal value As LineStyles)
Dim dp As DataPoint
For Each dp In mDataPoints
    dp.lineStyle = value
Next
mLineStyle = value
End Property

Public Property Get lineThickness() As Long
lineThickness = mLineThickness
End Property

Public Property Let lineThickness(ByVal value As Long)
Dim dp As DataPoint
For Each dp In mDataPoints
    dp.lineThickness = value
Next
mLineThickness = value
End Property

Public Property Get name() As String
name = mName
End Property

Public Property Let name(ByVal value As String)
mName = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                ByVal timestamp As Date, _
                Optional ByVal key As String) As DataPoint
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil25.DatPointSeries::add", _
                            "DataPoint series is finished"
Set add = New DataPoint
add.x = gCalculateX(timestamp, mController)
If key = "" Then
    add.key = GenerateGUIDString
Else
    add.key = key
End If
add.canvas = mCanvas
add.histBarWidth = mHistBarWidth
add.displayMode = mDisplayMode
add.eventProxy = mEventProxy
add.includeInAutoscale = mIncludeInAutoscale
add.layer = mLayer
add.lineColor = mLineColor
add.lineStyle = mLineStyle
add.lineThickness = mLineThickness
add.prevDataPoint = mPrevDataPoint
mDataPoints.add add, add.key
Set mPrevDataPoint = add
mEventProxy.addGraphicObject add, mSeriesId
End Function

Friend Sub finish()
mFinished = True
Set mCanvas = Nothing
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



