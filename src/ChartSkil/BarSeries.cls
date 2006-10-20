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

Private mFinished As Boolean

Private mBars As Collection

'Private mSurface As PictureBox
Private mCanvas As canvas
Private mLayer As Long

Private mEventProxy As ChartRegionEventProxy

Private mName As String
Private mTailThickness As Long
Private mOutlineThickness As Long
Private mUpColor As Long
Private mDownColor As Long
Private mDisplayAsCandlestick As Boolean
Private mSolidUpBody As Boolean
Private mBarThickness As Long
Private mCandleWidth As Single

Private mMinPeriodInView As Long
Private mMaxPeriodInView As Long

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mBars = New Collection
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

Public Property Get barThickness() As Long
barThickness = mBarThickness
End Property

Public Property Let barThickness(ByVal value As Long)
Dim bar As bar
If mBarThickness = value Then Exit Property
mBarThickness = value
For Each bar In mBars
    bar.barThickness = mBarThickness
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

Friend Property Let canvas(ByVal value As canvas)
Set mCanvas = value
End Property

Public Property Get candleWidth() As Single
candleWidth = mCandleWidth
End Property

Public Property Let candleWidth(ByVal value As Single)
Dim bar As bar
If mCandleWidth = value Then Exit Property
mCandleWidth = value
For Each bar In mBars
    bar.candleWidth = mCandleWidth
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

Public Property Get count() As Long
count = mBars.count
End Property

Public Property Get displayAsCandlestick() As Boolean
displayAsCandlestick = mDisplayAsCandlestick
End Property

Public Property Let displayAsCandlestick(ByVal value As Boolean)
Dim bar As bar
If mDisplayAsCandlestick = value Then Exit Property
mDisplayAsCandlestick = value
For Each bar In mBars
    bar.displayAsCandlestick = mDisplayAsCandlestick
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

Public Property Get downColor() As Long
downColor = mDownColor
End Property

Public Property Let downColor(ByVal value As Long)
Dim bar As bar
If mDownColor = value Then Exit Property
mDownColor = value
For Each bar In mBars
    bar.downColor = mDownColor
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

Friend Property Let eventProxy(ByVal value As ChartRegionEventProxy)
Set mEventProxy = value
End Property

Friend Property Let layer(ByVal value As Long)
mLayer = value
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Get maxPeriodInView() As Long
maxPeriodInView = mMaxPeriodInView
End Property

Public Property Get minPeriodInView() As Long
minPeriodInView = mMinPeriodInView
End Property

Public Property Get name() As String
name = mName
End Property

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get outlineThickness() As Long
outlineThickness = mOutlineThickness
End Property

Public Property Let outlineThickness(ByVal value As Long)
Dim bar As bar
If mOutlineThickness = value Then Exit Property
mOutlineThickness = value
For Each bar In mBars
    bar.outlineThickness = mOutlineThickness
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

Public Property Get solidUpBody() As Boolean
solidUpBody = mSolidUpBody
End Property

Public Property Let solidUpBody(ByVal value As Boolean)
Dim bar As bar
If mSolidUpBody = value Then Exit Property
mSolidUpBody = value
For Each bar In mBars
    bar.solidUpBody = mSolidUpBody
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

Public Property Get tailThickness() As Long
tailThickness = mTailThickness
End Property

Public Property Let tailThickness(ByVal value As Long)
Dim bar As bar
If mTailThickness = value Then Exit Property
mTailThickness = value
For Each bar In mBars
    bar.tailThickness = mTailThickness
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

Public Property Get upColor() As Long
upColor = mUpColor
End Property

Public Property Let upColor(ByVal value As Long)
Dim bar As bar
If mUpColor = value Then Exit Property
mUpColor = value
For Each bar In mBars
    bar.upColor = mUpColor
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

'================================================================================
' Methods
'================================================================================

Public Function addBar(ByVal periodNumber As Long) As bar
If mFinished Then err.Raise CommonErrorCodes.ErrIllegalStateException, _
                            "ChartSkil.BarSeries::add", _
                            "Bar series is finished"
Set addBar = New bar
addBar.periodNumber = periodNumber
addBar.barThickness = mBarThickness
addBar.candleWidth = mCandleWidth
addBar.displayAsCandlestick = mDisplayAsCandlestick
addBar.downColor = mDownColor
addBar.eventProxy = mEventProxy
addBar.layer = mLayer
addBar.outlineThickness = mOutlineThickness
addBar.solidUpBody = mSolidUpBody
addBar.canvas = mCanvas
addBar.tailThickness = mTailThickness
addBar.upColor = mUpColor

mBars.add addBar, CStr(periodNumber)
mEventProxy.addGraphicObject addBar
End Function

Friend Sub finish()
mFinished = True
Set mCanvas = Nothing
End Sub

Public Function item(ByVal periodNumber As Long) As bar
Set item = mBars(CStr(periodNumber))
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


