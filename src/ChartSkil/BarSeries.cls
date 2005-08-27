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
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private mBars As Collection

'Private mSurface As PictureBox
Private mCanvas As canvas
Private mLayer As Long

Private mEventProxy As ChartRegionEventProxy

Private mName As String
Private mTailThickness As Long
Private mOutlineThickness As Long
Private mUpColour As Long
Private mDownColour As Long
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

Public Property Get downColour() As Long
downColour = mDownColour
End Property

Public Property Let downColour(ByVal value As Long)
Dim bar As bar
If mDownColour = value Then Exit Property
mDownColour = value
For Each bar In mBars
    bar.downColour = mDownColour
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

Public Property Get upColour() As Long
upColour = mUpColour
End Property

Public Property Let upColour(ByVal value As Long)
Dim bar As bar
If mUpColour = value Then Exit Property
mUpColour = value
For Each bar In mBars
    bar.upColour = mUpColour
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

'================================================================================
' Methods
'================================================================================

Public Function addBar(ByVal periodNumber As Long) As bar
Set addBar = New bar
addBar.periodNumber = periodNumber
addBar.barThickness = mBarThickness
addBar.candleWidth = mCandleWidth
addBar.displayAsCandlestick = mDisplayAsCandlestick
addBar.downColour = mDownColour
addBar.eventProxy = mEventProxy
addBar.layer = mLayer
addBar.outlineThickness = mOutlineThickness
addBar.solidUpBody = mSolidUpBody
addBar.canvas = mCanvas
addBar.tailThickness = mTailThickness
addBar.upColour = mUpColour

mBars.Add addBar, CStr(periodNumber)
mEventProxy.addGraphicObject addBar
End Function

Friend Sub finish()
Set mCanvas = Nothing
End Sub

Public Function Item(ByVal periodNumber As Long) As bar
Set Item = mBars(CStr(periodNumber))
End Function

Friend Sub removeAll()
Dim i As Long
For i = mBars.count To 1 Step -1
    mBars.Remove (i)
Next
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mBars.[_NewEnum]
End Function

'================================================================================
' Helper Functions
'================================================================================


