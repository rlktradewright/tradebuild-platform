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
Private mDisplayMode As BarDisplayModes
Private mSolidUpBody As Boolean
Private mBarThickness As Long
Private mBarWidth As Single

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
Dim Bar As Bar
If mBarThickness = value Then Exit Property
mBarThickness = value
For Each Bar In mBars
    Bar.barThickness = mBarThickness
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

Friend Property Let canvas(ByVal value As canvas)
Set mCanvas = value
End Property

Public Property Get barWidth() As Single
barWidth = mBarWidth
End Property

Public Property Let barWidth(ByVal value As Single)
Dim Bar As Bar
If mBarWidth = value Then Exit Property
mBarWidth = value
For Each Bar In mBars
    Bar.barWidth = mBarWidth
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

Friend Property Let controller( _
                ByVal value As ChartController)
Set mController = value
End Property

Public Property Get count() As Long
count = mBars.count
End Property

Public Property Get displayMode() As BarDisplayModes
displayMode = mDisplayMode
End Property

Public Property Let displayMode(ByVal value As BarDisplayModes)
Dim Bar As Bar
If mDisplayMode = value Then Exit Property
mDisplayMode = value
For Each Bar In mBars
    Bar.displayMode = mDisplayMode
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

Public Property Get downColor() As Long
downColor = mDownColor
End Property

Public Property Let downColor(ByVal value As Long)
Dim Bar As Bar
If mDownColor = value Then Exit Property
mDownColor = value
For Each Bar In mBars
    Bar.downColor = mDownColor
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
Dim Bar As Bar
If mOutlineThickness = value Then Exit Property
mOutlineThickness = value
For Each Bar In mBars
    Bar.outlineThickness = mOutlineThickness
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

Public Property Get solidUpBody() As Boolean
solidUpBody = mSolidUpBody
End Property

Public Property Let solidUpBody(ByVal value As Boolean)
Dim Bar As Bar
If mSolidUpBody = value Then Exit Property
mSolidUpBody = value
For Each Bar In mBars
    Bar.solidUpBody = mSolidUpBody
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

Public Property Get tailThickness() As Long
tailThickness = mTailThickness
End Property

Public Property Let tailThickness(ByVal value As Long)
Dim Bar As Bar
If mTailThickness = value Then Exit Property
mTailThickness = value
For Each Bar In mBars
    Bar.tailThickness = mTailThickness
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

Public Property Get upColor() As Long
upColor = mUpColor
End Property

Public Property Let upColor(ByVal value As Long)
Dim Bar As Bar
If mUpColor = value Then Exit Property
mUpColor = value
For Each Bar In mBars
    Bar.upColor = mUpColor
Next
If mBars.count <> 0 Then mEventProxy.paintRegion
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                ByVal timestamp As Date, _
                Optional ByVal key As String) As Bar

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "ChartSkil25.BarSeries::add", _
                            "Bar series is finished"

Set add = New Bar

add.x = gCalculateX(timestamp, mController)
If key = "" Then
    add.key = GenerateGUIDString
Else
    add.key = key
End If
add.barThickness = mBarThickness
add.barWidth = mBarWidth
add.displayMode = mDisplayMode
add.downColor = mDownColor
add.eventProxy = mEventProxy
add.layer = mLayer
add.outlineThickness = mOutlineThickness
add.solidUpBody = mSolidUpBody
add.canvas = mCanvas
add.tailThickness = mTailThickness
add.upColor = mUpColor

mBars.add add, add.key
mEventProxy.addGraphicObject add
End Function

Friend Sub finish()
mFinished = True
Set mCanvas = Nothing
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


