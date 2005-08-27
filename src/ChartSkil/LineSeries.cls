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

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private mLines As Collection

'Private mSurface As PictureBox
Private mCanvas As canvas
Private mLayer As Long

Private mEventProxy As ChartRegionEventProxy

Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

Private mColour As Long
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
' Interface Members
'================================================================================

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

Public Property Get color() As Long
color = mColour
End Property

Public Property Let color(ByVal newValue As Long)
If mColour = newValue Then Exit Property
mColour = newValue
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

Public Function addLine() As ChartSkil.Line
Set addLine = New ChartSkil.Line
With addLine
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
    .color = mColour
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
mLines.Add addLine
mEventProxy.addGraphicObject addLine
End Function

Friend Sub finish()
Set mCanvas = Nothing
End Sub

Public Function Item(ByVal index As Long) As ChartSkil.Line
Attribute Item.VB_UserMemId = 0
Set Item = mLines(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
   Set NewEnum = mLines.[_NewEnum]
End Function

Friend Sub removeAll()
Dim i As Long
For i = mLines.count To 1 Step -1
    mLines.Remove (i)
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================


