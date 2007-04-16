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

Private mName As String
Private mSeriesId As Long

Private mFinished As Boolean

Private mLines As Collection

'Private mSurface As PictureBox
Private mCanvas As canvas
Private mLayer As Long

Private mEventProxy As ChartRegionEventProxy

Private mStyle As linestyle

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mLines = New Collection
mSeriesId = GenerateID
End Sub

'================================================================================
' IGraphicObjectSeries Interface Members
'================================================================================

Private Property Get IGraphicObjectSeries_count() As Long
IGraphicObjectSeries_count = mLines.count
End Property

Private Sub IGraphicObjectSeries_finish()
mFinished = True
End Sub

Private Function IGraphicObjectSeries_item(ByVal index As Long) As ChartSkilTypes.IGraphicObject
Set IGraphicObjectSeries_item = mLines.item(index)
End Function

Private Function IGraphicObjectSeries_NewEnum() As Object
Set IGraphicObjectSeries_NewEnum = NewEnum
End Function

Private Sub IGraphicObjectSeries_remove(ByVal index As Long)
mLines.remove index
End Sub

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get arrowEndColor() As Long
arrowEndColor = mStyle.arrowEndColor
End Property

Public Property Let arrowEndColor(ByVal value As Long)
If mStyle.arrowEndColor = value Then Exit Property
mStyle.arrowEndColor = value
If mLines.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get arrowEndFillColor() As Long
arrowEndFillColor = mStyle.arrowEndFillColor
End Property

Public Property Let arrowEndFillColor(ByVal value As Long)
If mStyle.arrowEndFillColor = value Then Exit Property
mStyle.arrowEndFillColor = value
If mLines.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get arrowEndFillStyle() As FillStyles
arrowEndFillStyle = mStyle.arrowEndFillStyle
End Property

Public Property Let arrowEndFillStyle(ByVal value As FillStyles)
If mStyle.arrowEndFillStyle = value Then Exit Property
mStyle.arrowEndFillStyle = value
If mLines.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get arrowEndLength() As Long
arrowEndLength = mStyle.arrowEndLength
End Property

Public Property Let arrowEndLength(ByVal value As Long)
If mStyle.arrowEndLength = value Then Exit Property
mStyle.arrowEndLength = value
If mLines.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get arrowEndStyle() As ArrowStyles
arrowEndStyle = mStyle.arrowEndStyle
End Property

Public Property Let arrowEndStyle(ByVal value As ArrowStyles)
If mStyle.arrowEndStyle = value Then Exit Property
mStyle.arrowEndStyle = value
If mLines.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get arrowEndWidth() As Long
arrowEndWidth = mStyle.arrowEndWidth
End Property

Public Property Let arrowEndWidth(ByVal value As Long)
If mStyle.arrowEndWidth = value Then Exit Property
mStyle.arrowEndWidth = value
If mLines.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get arrowStartColor() As Long
arrowStartColor = mStyle.arrowStartColor
End Property

Public Property Let arrowStartColor(ByVal value As Long)
If mStyle.arrowStartColor = value Then Exit Property
mStyle.arrowStartColor = value
If mLines.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get arrowStartFillColor() As Long
arrowStartFillColor = mStyle.arrowStartFillColor
End Property

Public Property Let arrowStartFillColor(ByVal value As Long)
If mStyle.arrowStartFillColor = value Then Exit Property
mStyle.arrowStartFillColor = value
If mLines.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get arrowStartfillstyle() As FillStyles
arrowStartfillstyle = arrowStartfillstyle
End Property

Public Property Let arrowStartfillstyle(ByVal value As FillStyles)
If mStyle.arrowStartfillstyle = value Then Exit Property
mStyle.arrowStartfillstyle = value
If mLines.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get arrowStartLength() As Long
arrowStartLength = mStyle.arrowStartLength
End Property

Public Property Let arrowStartLength(ByVal value As Long)
If mStyle.arrowStartLength = value Then Exit Property
mStyle.arrowStartLength = value
If mLines.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get arrowStartStyle() As ArrowStyles
arrowStartStyle = mStyle.arrowStartStyle
End Property

Public Property Let arrowStartStyle(ByVal value As ArrowStyles)
If mStyle.arrowStartStyle = value Then Exit Property
mStyle.arrowStartStyle = value
If mLines.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get arrowStartWidth() As Long
arrowStartWidth = mStyle.arrowStartWidth
End Property

Public Property Let arrowStartWidth(ByVal value As Long)
If mStyle.arrowStartWidth = value Then Exit Property
mStyle.arrowStartWidth = value
If mLines.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Friend Property Let canvas(ByVal value As canvas)
Set mCanvas = value
End Property

Public Property Get Color() As Long
Color = mStyle.Color
End Property

Public Property Let Color(ByVal value As Long)
If mStyle.Color = value Then Exit Property
mStyle.Color = value
If mLines.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get count() As Long
count = mLines.count
End Property

Friend Property Let eventProxy(ByVal value As ChartRegionEventProxy)
Set mEventProxy = value
End Property

Friend Property Get extended() As Boolean
extended = mStyle.extended
End Property

Public Property Let extended(ByVal value As Boolean)
If mStyle.extended = value Then Exit Property
mStyle.extended = value
If mLines.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get extendAfter() As Boolean
extendAfter = mStyle.extendAfter
End Property

Public Property Let extendAfter(ByVal value As Boolean)
If mStyle.extendAfter = value Then Exit Property
mStyle.extendAfter = value
If mLines.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get extendBefore() As Boolean
extendBefore = mStyle.extendBefore
End Property

Public Property Let extendBefore(ByVal value As Boolean)
If mStyle.extendBefore = value Then Exit Property
mStyle.extendBefore = value
If mLines.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Friend Property Get fixedX() As Boolean
fixedX = mStyle.fixedX
End Property

Public Property Let fixedX(ByVal value As Boolean)
If mStyle.fixedX = value Then Exit Property
mStyle.fixedX = value
If mLines.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Friend Property Get fixedY() As Boolean
fixedY = mStyle.fixedY
End Property

Public Property Let fixedY(ByVal value As Boolean)
If mStyle.fixedY = value Then Exit Property
mStyle.fixedY = value
If mLines.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get includeInAutoscale() As Boolean
includeInAutoscale = mStyle.includeInAutoscale
End Property

Public Property Let includeInAutoscale(ByVal value As Boolean)
If mStyle.includeInAutoscale = value Then Exit Property
mStyle.includeInAutoscale = value
If mLines.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion True
End If
End Property

'Public Property Get keepInView() As Boolean
'keepInView = mStyle.KeepInView
'End Property

'Public Property Let keepInView(ByVal value As Boolean)
'If mStyle.KeepInView = value Then Exit Property
'undraw
'mStyle.KeepInView = value
'update
'End Property

Public Property Let layer(ByVal value As Long)
mLayer = value
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
If mLines.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get style() As linestyle
Set style = mStyle.clone
End Property

Public Property Let style(ByVal value As linestyle)
Dim somethingChanged As Boolean
Dim recalcBoundaries As Boolean
Dim recalcRegion As Boolean

If value.arrowEndColor <> mStyle.arrowEndColor Then
    mStyle.arrowEndColor = value.arrowEndColor
    somethingChanged = True
End If
If value.arrowEndFillColor <> mStyle.arrowEndFillColor Then
    mStyle.arrowEndFillColor = value.arrowEndFillColor
    somethingChanged = True
End If
If value.arrowEndFillStyle <> mStyle.arrowEndFillStyle Then
    mStyle.arrowEndFillStyle = value.arrowEndFillStyle
    somethingChanged = True
End If
If value.arrowEndLength <> mStyle.arrowEndLength Then
    mStyle.arrowEndLength = value.arrowEndLength
    somethingChanged = True
    recalcBoundaries = True
End If
If value.arrowEndStyle <> mStyle.arrowEndStyle Then
    mStyle.arrowEndStyle = value.arrowEndStyle
    somethingChanged = True
End If
If value.arrowEndWidth <> mStyle.arrowEndWidth Then
    mStyle.arrowEndWidth = value.arrowEndWidth
    somethingChanged = True
    recalcBoundaries = True
End If
If value.arrowStartColor <> mStyle.arrowStartColor Then
    mStyle.arrowStartColor = value.arrowStartColor
    somethingChanged = True
End If
If value.arrowStartFillColor <> mStyle.arrowStartFillColor Then
    mStyle.arrowStartFillColor = value.arrowStartFillColor
    somethingChanged = True
End If
If value.arrowStartfillstyle <> mStyle.arrowStartfillstyle Then
    mStyle.arrowStartfillstyle = value.arrowStartfillstyle
    somethingChanged = True
End If
If value.arrowStartLength <> mStyle.arrowStartLength Then
    mStyle.arrowStartLength = value.arrowStartLength
    somethingChanged = True
    recalcBoundaries = True
End If
If value.arrowStartStyle <> mStyle.arrowStartStyle Then
    mStyle.arrowStartStyle = value.arrowStartStyle
    somethingChanged = True
End If
If value.arrowStartWidth <> mStyle.arrowStartWidth Then
    mStyle.arrowStartWidth = value.arrowStartWidth
    somethingChanged = True
    recalcBoundaries = True
End If
If value.Color <> mStyle.Color Then
    mStyle.Color = value.Color
    somethingChanged = True
End If
If value.extendAfter <> mStyle.extendAfter Then
    mStyle.extendAfter = value.extendAfter
    somethingChanged = True
    recalcBoundaries = True
End If
If value.extendBefore <> mStyle.extendBefore Then
    mStyle.extendBefore = value.extendBefore
    somethingChanged = True
    recalcBoundaries = True
End If
If value.extended <> mStyle.extended Then
    mStyle.extended = value.extended
    somethingChanged = True
    recalcBoundaries = True
End If
If value.fixedX <> mStyle.fixedX Then
    mStyle.fixedX = value.fixedX
    somethingChanged = True
    recalcBoundaries = True
End If
If value.fixedY <> mStyle.fixedY Then
    mStyle.fixedY = value.fixedY
    somethingChanged = True
    recalcBoundaries = True
End If
If value.includeInAutoscale <> mStyle.includeInAutoscale Then
    mStyle.includeInAutoscale = value.includeInAutoscale
    somethingChanged = True
    recalcBoundaries = True
    recalcRegion = True
End If
If value.linestyle <> mStyle.linestyle Then
    mStyle.linestyle = value.linestyle
    somethingChanged = True
End If
If value.thickness <> mStyle.thickness Then
    mStyle.thickness = value.thickness
    somethingChanged = True
    recalcBoundaries = True
End If

If somethingChanged Then
    If recalcBoundaries Then If mLines.count <> 0 Then mEventProxy.recalcBoundary mSeriesId
    If mLines.count <> 0 Then mEventProxy.paintRegion recalcRegion
End If
End Property

Public Property Get thickness() As Long
thickness = mStyle.thickness
End Property

Public Property Let thickness(ByVal value As Long)
If mStyle.thickness = value Then Exit Property
mStyle.thickness = value
If mLines.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                Optional ByVal key As String, _
                Optional ByVal localStyle As linestyle) As ChartSkil26.Line
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil26.LineSeries::add", _
                            "Line series is finished"
Set add = New ChartSkil26.Line
If key = "" Then
    key = GenerateGUIDString
End If
add.initialise key, mEventProxy, mLayer, mCanvas, mStyle, localStyle

mLines.add add, key
mEventProxy.addGraphicObject add, mSeriesId
End Function

Friend Sub finish()
mFinished = True
Set mCanvas = Nothing
End Sub

Friend Sub initialise( _
                ByVal name As String, _
                ByVal eventProxy As ChartRegionEventProxy, _
                ByVal layer As Long, _
                ByVal canvas As canvas, _
                ByVal style As linestyle)
mName = name
Set mEventProxy = eventProxy
mLayer = layer
Set mCanvas = canvas
Set mStyle = style.clone    ' style cannot be shared between BarSeries
End Sub

Public Function item(ByVal index As Variant) As ChartSkil26.Line
Attribute item.VB_UserMemId = 0
Set item = mLines(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
   Set NewEnum = mLines.[_NewEnum]
End Function

Friend Sub removeAll()
Set mLines = New Collection
End Sub

'================================================================================
' Helper Functions
'================================================================================


