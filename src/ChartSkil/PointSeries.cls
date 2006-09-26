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
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

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
Private mDisplayMode As DisplayModes
Private mIncludeInAutoscale As Boolean

Private mHistBarWidth As Single

Private mMinPeriodInView As Long
Private mMaxPeriodInView As Long

Private mPrevDataPoint As DataPoint

'================================================================================
' Enums
'================================================================================

Public Enum DisplayModes
    displayAsPoints = 1
    DisplayAsLines
    DisplayAsSteppedLines
    displayAsHistogram
End Enum

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mDataPoints = New Collection
mLineStyle = LineStyles.LineSolid
mHistBarWidth = DefaultHistBarWidth
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

Public Property Get count() As Long
count = mDataPoints.count
End Property

Public Property Get displayMode() As DisplayModes
displayMode = mDisplayMode
End Property

Public Property Let displayMode(ByVal value As DisplayModes)
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

'================================================================================
' Methods
'================================================================================

Public Function addDataPoint( _
                ByVal periodNumber As Long, _
                Optional ByVal key As Variant) As DataPoint
Set addDataPoint = New DataPoint
addDataPoint.canvas = mCanvas
addDataPoint.periodNumber = periodNumber
addDataPoint.histBarWidth = mHistBarWidth
addDataPoint.displayMode = mDisplayMode
addDataPoint.eventProxy = mEventProxy
addDataPoint.includeInAutoscale = mIncludeInAutoscale
addDataPoint.layer = mLayer
addDataPoint.lineColor = mLineColor
addDataPoint.lineStyle = mLineStyle
addDataPoint.lineThickness = mLineThickness
addDataPoint.prevDataPoint = mPrevDataPoint
If IsMissing(key) Then
    mDataPoints.add addDataPoint
Else
    mDataPoints.add addDataPoint, key
End If
Set mPrevDataPoint = addDataPoint
mEventProxy.addGraphicObject addDataPoint
End Function

Friend Sub finish()
Set mCanvas = Nothing
End Sub

Public Function item(ByVal key As Variant) As DataPoint
Attribute item.VB_UserMemId = 0
Set item = mDataPoints(CStr(key))
End Function

Friend Sub removeAll()
Dim i As Long
For i = mDataPoints.count To 1 Step -1
    mDataPoints.remove (i)
Next
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mDataPoints.[_NewEnum]
End Function

'================================================================================
' Helper Functions
'================================================================================



