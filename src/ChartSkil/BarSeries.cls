VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarSeries"
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

Private mBars As Collection

Private mCanvas As canvas
Private mLayer As Long

Private mEventProxy As ChartRegionEventProxy

Private mStyle As BarStyle

Private mName As String

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mBars = New Collection
mSeriesId = GenerateID
End Sub

'================================================================================
' IGraphicObjectSeries Interface Members
'================================================================================

Private Property Get IGraphicObjectSeries_count() As Long
IGraphicObjectSeries_count = mBars.count
End Property

Private Sub IGraphicObjectSeries_finish()
mFinished = True
End Sub

Private Function IGraphicObjectSeries_item(ByVal index As Long) As ChartSkilTypes.IGraphicObject
Set IGraphicObjectSeries_item = mBars.item(index)
End Function

Private Function IGraphicObjectSeries_NewEnum() As Object
Set IGraphicObjectSeries_NewEnum = NewEnum
End Function

Private Sub IGraphicObjectSeries_remove(ByVal index As Long)
mBars.remove index
End Sub

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let barColor(ByVal value As Long)
If mStyle.barColor = value Then Exit Property
mStyle.barColor = value
If mBars.count <> 0 Then mEventProxy.paintRegion False
End Property

Public Property Get barThickness() As Long
barThickness = mStyle.barThickness
End Property

Public Property Let barThickness(ByVal value As Long)
If mStyle.barThickness = value Then Exit Property
mStyle.barThickness = value
If mBars.count <> 0 Then mEventProxy.paintRegion False
End Property

Public Property Get barWidth() As Single
barWidth = mStyle.barWidth
End Property

Public Property Let barWidth(ByVal value As Single)
If mStyle.barWidth = value Then Exit Property
mStyle.barWidth = value
If mBars.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get count() As Long
count = mBars.count
End Property

Public Property Get displayMode() As BarDisplayModes
displayMode = mStyle.displayMode
End Property

Public Property Let displayMode(ByVal value As BarDisplayModes)
If mStyle.displayMode = value Then Exit Property
mStyle.displayMode = value
If mBars.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get downColor() As Long
downColor = mStyle.downColor
End Property

Public Property Let downColor(ByVal value As Long)
If mStyle.downColor = value Then Exit Property
mStyle.downColor = value
If mBars.count <> 0 Then mEventProxy.paintRegion False
End Property

Friend Property Get includeInAutoscale() As Boolean
includeInAutoscale = mStyle.includeInAutoscale
End Property

Public Property Let includeInAutoscale(ByVal value As Boolean)
If mStyle.includeInAutoscale = value Then Exit Property
mStyle.includeInAutoscale = value
If mBars.count <> 0 Then
    mEventProxy.paintRegion True
End If
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Get name() As String
name = mName
End Property

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get outlineThickness() As Long
outlineThickness = mStyle.outlineThickness
End Property

Public Property Let outlineThickness(ByVal value As Long)
If mStyle.outlineThickness = value Then Exit Property
mStyle.outlineThickness = value
If mBars.count <> 0 Then mEventProxy.paintRegion False
End Property

Public Property Get solidUpBody() As Boolean
solidUpBody = mStyle.solidUpBody
End Property

Public Property Let solidUpBody(ByVal value As Boolean)
If mStyle.solidUpBody = value Then Exit Property
mStyle.solidUpBody = value
If mBars.count <> 0 Then mEventProxy.paintRegion False
End Property

Public Property Get style() As BarStyle
Set style = mStyle.clone
End Property

Public Property Let style( _
                ByVal value As BarStyle)
Dim somethingChanged As Boolean

If value.barColor <> mStyle.barColor Then
    somethingChanged = True
    mStyle.barColor = value.barColor
End If
If value.barThickness <> mStyle.barThickness Then
    somethingChanged = True
    mStyle.barThickness = value.barThickness
End If
If value.barWidth <> mStyle.barWidth Then
    somethingChanged = True
    mStyle.barWidth = value.barWidth
    If mBars.count <> 0 Then mEventProxy.recalcBoundary mSeriesId
End If
If value.displayMode <> mStyle.displayMode Then
    somethingChanged = True
    mStyle.displayMode = value.displayMode
End If
If value.downColor <> mStyle.downColor Then
    somethingChanged = True
    mStyle.downColor = value.downColor
End If
If value.outlineThickness <> mStyle.outlineThickness Then
    somethingChanged = True
    mStyle.outlineThickness = value.outlineThickness
End If
If value.solidUpBody <> mStyle.solidUpBody Then
    somethingChanged = True
    mStyle.solidUpBody = value.solidUpBody
End If
If value.tailThickness <> mStyle.tailThickness Then
    somethingChanged = True
    mStyle.tailThickness = value.tailThickness
End If
If value.upColor <> mStyle.upColor Then
    somethingChanged = True
    mStyle.upColor = value.upColor
End If

If somethingChanged Then
    If mBars.count <> 0 Then mEventProxy.paintRegion False
End If
End Property

Public Property Get tailThickness() As Long
tailThickness = mStyle.tailThickness
End Property

Public Property Let tailThickness(ByVal value As Long)
If mStyle.tailThickness = value Then Exit Property
mStyle.tailThickness = value
If mBars.count <> 0 Then mEventProxy.paintRegion False
End Property

Public Property Get upColor() As Long
upColor = mStyle.upColor
End Property

Public Property Let upColor(ByVal value As Long)
If mStyle.upColor = value Then Exit Property
mStyle.upColor = value
If mBars.count <> 0 Then mEventProxy.paintRegion False
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                ByVal timestamp As Date, _
                Optional ByVal key As String, _
                Optional ByVal localStyle As BarStyle) As Bar

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil25.BarSeries::add", _
                            "Bar series is finished"

Set add = New Bar

add.x = gCalculateX(timestamp, mController)
If key = "" Then
    key = GenerateGUIDString
End If
add.initialise key, mEventProxy, mLayer, mCanvas, mStyle, localStyle

mBars.add add, key
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
                ByVal style As BarStyle)
mName = name
Set mController = controller
Set mEventProxy = eventProxy
mLayer = layer
Set mCanvas = canvas
Set mStyle = style.clone    ' style cannot be shared between BarSeries
End Sub

Public Function item(ByVal index As Variant) As Bar
Set item = mBars(index)
End Function

Friend Sub removeAll()
Dim i As Long
Set mBars = New Collection
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mBars.[_NewEnum]
End Function

'================================================================================
' Helper Functions
'================================================================================


