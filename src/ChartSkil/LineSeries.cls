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

Private mFinished As Boolean

Private mLines As Collection

'Private mSurface As PictureBox
Private mCanvas As canvas
Private mLayer As Long

Private mEventProxy As ChartRegionEventProxy

Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

Private mColor As Long
Private mThickness As Long
Private mStyle As LineStyles
Private mExtendBefore As Boolean
Private mExtendAfter As Boolean
Private mArrowStartStyle As ArrowStyles
Private mArrowStartLength As Long
Private mArrowStartWidth As Long
Private mArrowStartColor As Long
Private mArrowStartFillColor As Long
Private mArrowStartFillStyle As FillStyles
Private mArrowEndStyle As ArrowStyles
Private mArrowEndLength As Long
Private mArrowEndWidth As Long
Private mArrowEndColor As Long
Private mArrowEndFillColor As Long
Private mArrowEndFillStyle As FillStyles


'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mLines = New Collection
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
arrowEndColor = mArrowEndColor
End Property

Public Property Let arrowEndColor(ByVal newValue As Long)
mArrowEndColor = newValue
End Property

Public Property Get arrowEndFillColor() As Long
arrowEndFillColor = mArrowEndFillColor
End Property

Public Property Let arrowEndFillColor(ByVal newValue As Long)
mArrowEndFillColor = newValue
End Property

Public Property Get arrowEndFillStyle() As FillStyles
arrowEndFillStyle = mArrowEndFillStyle
End Property

Public Property Let arrowEndFillStyle(ByVal newValue As FillStyles)
mArrowEndFillStyle = newValue
End Property

Public Property Get arrowEndLength() As Long
arrowEndLength = mArrowEndLength
End Property

Public Property Let arrowEndLength(ByVal newValue As Long)
mArrowEndLength = newValue
End Property

Public Property Get arrowEndStyle() As ArrowStyles
arrowEndStyle = mArrowEndStyle
End Property

Public Property Let arrowEndStyle(ByVal newValue As ArrowStyles)
mArrowEndStyle = newValue
End Property

Public Property Get arrowEndWidth() As Long
arrowEndWidth = mArrowEndWidth
End Property

Public Property Let arrowEndWidth(ByVal newValue As Long)
mArrowEndWidth = newValue
End Property

Public Property Get arrowStartColor() As Long
arrowStartColor = mArrowStartColor
End Property

Public Property Let arrowStartColor(ByVal newValue As Long)
mArrowStartColor = newValue
End Property

Public Property Get arrowStartFillColor() As Long
arrowStartFillColor = mArrowStartFillColor
End Property

Public Property Let arrowStartFillColor(ByVal newValue As Long)
mArrowStartFillColor = newValue
End Property

Public Property Get arrowStartFillStyle() As FillStyles
arrowStartFillStyle = arrowStartFillStyle
End Property

Public Property Let arrowStartFillStyle(ByVal newValue As FillStyles)
mArrowStartFillStyle = newValue
End Property

Public Property Get arrowStartLength() As Long
arrowStartLength = mArrowStartLength
End Property

Public Property Let arrowStartLength(ByVal newValue As Long)
mArrowStartLength = newValue
End Property

Public Property Get arrowStartStyle() As ArrowStyles
arrowStartStyle = mArrowStartStyle
End Property

Public Property Let arrowStartStyle(ByVal newValue As ArrowStyles)
mArrowStartStyle = newValue
End Property

Public Property Get arrowStartWidth() As Long
arrowStartWidth = mArrowStartWidth
End Property

Public Property Let arrowStartWidth(ByVal newValue As Long)
mArrowStartWidth = newValue
End Property

Friend Property Let canvas(ByVal value As canvas)
Set mCanvas = value
End Property

Public Property Get Color() As Long
Color = mColor
End Property

Public Property Let Color(ByVal newValue As Long)
If mColor = newValue Then Exit Property
mColor = newValue
End Property

Public Property Get count() As Long
count = mLines.count
End Property

Friend Property Let eventProxy(ByVal value As ChartRegionEventProxy)
Set mEventProxy = value
End Property

Friend Property Get extended() As Boolean
extended = mExtended
End Property

Public Property Let extended(ByVal newValue As Boolean)
If mExtended = newValue Then Exit Property
mExtended = newValue
End Property

Public Property Get extendAfter() As Boolean
extendAfter = mExtendAfter
End Property

Public Property Let extendAfter(ByVal newValue As Boolean)
If mExtendAfter = newValue Then Exit Property
mExtendAfter = newValue
End Property

Public Property Get extendBefore() As Boolean
extendBefore = mExtendBefore
End Property

Public Property Let extendBefore(ByVal newValue As Boolean)
If mExtendBefore = newValue Then Exit Property
mExtendBefore = newValue
End Property

Friend Property Get fixedX() As Boolean
fixedX = mFixedX
End Property

Public Property Let fixedX(ByVal newValue As Boolean)
If mFixedX = newValue Then Exit Property
mFixedX = newValue
End Property

Friend Property Get fixedY() As Boolean
fixedY = mFixedY
End Property

Public Property Let fixedY(ByVal newValue As Boolean)
If mFixedY = newValue Then Exit Property
mFixedY = newValue
End Property

Public Property Get includeInAutoscale() As Boolean
includeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let includeInAutoscale(ByVal newValue As Boolean)
If mIncludeInAutoscale = newValue Then Exit Property
mIncludeInAutoscale = newValue
End Property

'Public Property Get keepInView() As Boolean
'keepInView = mKeepInView
'End Property

'Public Property Let keepInView(ByVal newValue As Boolean)
'If mKeepInView = newValue Then Exit Property
'undraw
'mKeepInView = newValue
'update
'End Property

Public Property Let layer(ByVal value As Long)
mLayer = value
End Property

Public Property Get layer() As Long
layer = mLayer
End Property
Public Property Get style() As LineStyles
style = mStyle
End Property

Public Property Let style(ByVal newValue As LineStyles)
If mStyle = newValue Then Exit Property
mStyle = newValue
End Property

Public Property Get thickness() As Long
thickness = mThickness
End Property

Public Property Let thickness(ByVal newValue As Long)
If mThickness = newValue Then Exit Property
mThickness = newValue
End Property

'================================================================================
' Methods
'================================================================================

Friend Function add(Optional ByVal key As String) As ChartSkil.Line
If mFinished Then err.Raise CommonErrorCodes.ErrIllegalStateException, _
                            "ChartSkil.LineSeries::add", _
                            "Line series is finished"
Set add = New ChartSkil.Line
With add
    .eventProxy = mEventProxy
    .canvas = mCanvas
    .arrowEndColor = mArrowEndColor
    .arrowEndFillColor = mArrowEndFillColor
    .arrowEndFillStyle = mArrowEndFillStyle
    .arrowEndLength = mArrowEndLength
    .arrowEndStyle = mArrowEndStyle
    .arrowEndWidth = mArrowEndWidth
    .arrowStartColor = mArrowStartColor
    .arrowStartFillColor = mArrowStartFillColor
    .arrowStartFillStyle = mArrowStartFillStyle
    .arrowStartLength = mArrowStartLength
    .arrowStartStyle = mArrowStartStyle
    .arrowStartWidth = mArrowStartWidth
    .Color = mColor
    .extendAfter = mExtendAfter
    .extendBefore = mExtendBefore
    .extended = mExtended
    .fixedX = mFixedX
    .fixedY = mFixedY
    .style = mStyle
    .thickness = mThickness
    .includeInAutoscale = mIncludeInAutoscale
    .layer = mLayer
End With
If key <> "" Then
    mLines.add add, key
Else
    mLines.add add
End If
mEventProxy.addGraphicObject add
End Function

Public Function addLine() As ChartSkil.Line
Set addLine = add
End Function

Friend Sub finish()
mFinished = True
Set mCanvas = Nothing
End Sub

Public Function item(ByVal index As Long) As ChartSkil.Line
Attribute item.VB_UserMemId = 0
Set item = mLines(index)
End Function

Friend Function ItemByKey(ByVal key As String) As ChartSkil.Line
Set ItemByKey = mLines(key)
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


