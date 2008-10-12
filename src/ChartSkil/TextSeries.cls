VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements ChartSkilTypes.IGraphicObjectSeries
Implements Task

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private mSeriesId As Long

Private mFinished As Boolean

Private mTexts As Collection

Private mCanvas As canvas
Private mLayer As Long

Private mEventProxy As ChartRegionEventProxy

Private mStyle As TextStyle

Private mName As String

Private mTaskContext As TaskContext

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTexts = New Collection
mSeriesId = GenerateID
End Sub

'================================================================================
' IGraphicObjectSeries Interface Members
'================================================================================

Private Property Get IGraphicObjectSeries_count() As Long
IGraphicObjectSeries_count = mTexts.count
End Property

Private Sub IGraphicObjectSeries_finish()
mFinished = True
End Sub

Private Function IGraphicObjectSeries_item(ByVal index As Long) As ChartSkilTypes.IGraphicObject
Set IGraphicObjectSeries_item = mTexts.item(index)
End Function

Private Function IGraphicObjectSeries_NewEnum() As Object
Set IGraphicObjectSeries_NewEnum = NewEnum
End Function

Private Sub IGraphicObjectSeries_remove(ByVal index As Long)
mTexts.remove index
End Sub

'================================================================================
' Task Interface Members
'================================================================================

Private Sub Task_cancel()

End Sub

Private Sub Task_run()
Dim j As Long

Do While mTexts.count <> 0
    mTexts.remove mTexts.count
    j = j + 1
    If j = 100 Then Exit Sub
Loop

mTaskContext.finish 0, False
End Sub

Private Property Let Task_TaskContext(ByVal RHS As TWUtilities30.TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.name
End Property

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get align() As TextAlignModes
align = mStyle.align
End Property

Public Property Let align(ByVal value As TextAlignModes)
If mStyle.align = value Then Exit Property
mStyle.align = value
If mTexts.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Friend Property Get box() As Boolean
box = mStyle.box
End Property

Public Property Let box(ByVal value As Boolean)
If mStyle.box = value Then Exit Property
mStyle.box = value
If mTexts.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get boxColor() As Long
boxColor = mStyle.boxColor
End Property

Public Property Let boxColor(ByVal value As Long)
If mStyle.boxColor = value Then Exit Property
mStyle.boxColor = value
If mTexts.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Friend Property Get boxFillColor() As Long
boxFillColor = mStyle.boxFillColor
End Property

Public Property Let boxFillColor(ByVal value As Long)
If mStyle.boxFillColor = value Then Exit Property
mStyle.boxFillColor = value
If mTexts.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get boxFillStyle() As FillStyles
boxFillStyle = mStyle.boxFillStyle
End Property

Public Property Let boxFillStyle(ByVal value As FillStyles)
If mStyle.boxFillStyle = value Then Exit Property
mStyle.boxFillStyle = value
If mTexts.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get boxStyle() As LineStyles
boxStyle = mStyle.boxStyle
End Property

Public Property Let boxStyle(ByVal value As LineStyles)
If mStyle.boxStyle = value Then Exit Property
mStyle.boxStyle = value
If mTexts.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get boxThickness() As Long
boxThickness = mStyle.boxThickness
End Property

Public Property Let boxThickness(ByVal value As Long)
If mStyle.boxThickness = value Then Exit Property
mStyle.boxThickness = value
If mTexts.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get Color() As Long
Color = mStyle.Color
End Property

Public Property Let Color(ByVal value As Long)
If mStyle.Color = value Then Exit Property
mStyle.Color = value
If mTexts.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get count() As Long
count = mTexts.count
End Property

Friend Property Get extended() As Boolean
extended = mStyle.extended
End Property

Public Property Let extended(ByVal value As Boolean)
If mStyle.extended = value Then Exit Property
mStyle.extended = value
If mTexts.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Friend Property Get fixedX() As Boolean
fixedX = mStyle.fixedX
End Property

Public Property Let fixedX(ByVal value As Boolean)
If mStyle.fixedX = value Then Exit Property
mStyle.fixedX = value
If mTexts.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Friend Property Get fixedY() As Boolean
fixedY = mStyle.fixedY
End Property

Public Property Let fixedY(ByVal value As Boolean)
If mStyle.fixedY = value Then Exit Property
mStyle.fixedY = value
If mTexts.count <> 0 Then
    mEventProxy.paintRegion False
End If
End Property

Public Property Get font() As StdFont
Set font = mStyle.font
End Property

Public Property Let font(ByVal value As StdFont)
'If mStyle.font is value Then Exit Property
mStyle.font = value
If mTexts.count <> 0 Then
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
If mTexts.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion True
End If
End Property

Friend Property Let layer(ByVal value As Long)
mLayer = value
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Get paddingX() As Double
paddingX = mStyle.paddingX
End Property

Public Property Let paddingX(ByVal value As Double)
If mStyle.paddingX = value Then Exit Property
mStyle.paddingX = value
If mTexts.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get paddingY() As Double
paddingY = mStyle.paddingY
End Property

Public Property Let paddingY(ByVal value As Double)
If mStyle.paddingY = value Then Exit Property
mStyle.paddingY = value
If mTexts.count <> 0 Then
    mEventProxy.recalcBoundary mSeriesId
    mEventProxy.paintRegion False
End If
End Property

Public Property Get style() As TextStyle
Set style = mStyle.clone
End Property

Public Property Let style(ByVal value As TextStyle)
Dim somethingChanged As Boolean
Dim recalcBoundaries As Boolean
Dim recalcRegion As Boolean

If value.align <> mStyle.align Then
    mStyle.align = value.align
    somethingChanged = True
    recalcBoundaries = True
End If
If value.box <> mStyle.box Then
    mStyle.box = value.box
    somethingChanged = True
    recalcBoundaries = True
End If
If value.boxColor <> mStyle.boxColor Then
    mStyle.boxColor = value.boxColor
    somethingChanged = True
End If
If value.boxFillColor <> mStyle.boxFillColor Then
    mStyle.boxFillColor = value.boxFillColor
    somethingChanged = True
End If
If value.boxFillStyle <> mStyle.boxFillStyle Then
    mStyle.boxFillStyle = value.boxFillStyle
    somethingChanged = True
End If
If value.boxStyle <> mStyle.boxStyle Then
    mStyle.boxStyle = value.boxStyle
    somethingChanged = True
    recalcBoundaries = True
End If
If value.boxThickness <> mStyle.boxThickness Then
    mStyle.boxThickness = value.boxThickness
    somethingChanged = True
    recalcBoundaries = True
End If
If value.Color <> mStyle.Color Then
    mStyle.Color = value.Color
    somethingChanged = True
End If
If value.extended <> mStyle.extended Then
    mStyle.extended = value.extended
    somethingChanged = True
End If
If value.fixedX <> mStyle.fixedX Then
    mStyle.fixedX = value.fixedX
    somethingChanged = True
End If
If value.fixedY <> mStyle.fixedY Then
    mStyle.fixedY = value.fixedY
    somethingChanged = True
End If
If Not value.font Is mStyle.font Then
    mStyle.font = value.font
    somethingChanged = True
    recalcBoundaries = True
End If
If value.includeInAutoscale <> mStyle.includeInAutoscale Then
    mStyle.includeInAutoscale = value.includeInAutoscale
    somethingChanged = True
    recalcBoundaries = True
    recalcRegion = True
End If
If value.paddingX <> mStyle.paddingX Then
    mStyle.paddingX = value.paddingX
    somethingChanged = True
    recalcBoundaries = True
End If
If value.paddingY <> mStyle.paddingY Then
    mStyle.paddingY = value.paddingY
    somethingChanged = True
    recalcBoundaries = True
End If

If somethingChanged Then
    If recalcBoundaries Then If mTexts.count <> 0 Then mEventProxy.recalcBoundary mSeriesId
    If mTexts.count <> 0 Then mEventProxy.paintRegion recalcRegion
End If
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                Optional ByVal key As String, _
                Optional ByVal localStyle As TextStyle) As text
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil26.TextSeries::add", _
                            "Text series is finished"
Set add = New text
If key = "" Then
    key = GenerateGUIDString
End If
add.initialise key, mEventProxy, mLayer, mCanvas, mStyle, localStyle

mTexts.add add, key
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
                ByVal style As TextStyle)
mName = name
Set mEventProxy = eventProxy
mLayer = layer
Set mCanvas = canvas
Set mStyle = style.clone    ' style cannot be shared between TextSeries
End Sub

Public Function item(ByVal index As Variant) As text
Attribute item.VB_UserMemId = 0
Set item = mTexts(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mTexts.[_NewEnum]
End Function

Friend Sub removeAll()
StartTask Me, PriorityLow
End Sub

'================================================================================
' Helper Functions
'================================================================================


