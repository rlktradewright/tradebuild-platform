VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mIncludeInChart As Boolean
Private mValueName As String
Private mColor As Long
Private mDisplayMode As ChartSkil.DisplayModes
Private mLineThickness As Long
Private mIncludeInAutoscale As Boolean
Private mLineStyle As ChartSkil.LineStyles
Private mChartRegionName As String
Private mLayer As Long
Private mHistogramBarWidth As Single
Private mMultipleValuesPerBar As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

mColor = vbBlack
mDisplayMode = ChartSkil.DisplayModes.DisplayAsLines
mLineThickness = 1
mIncludeInAutoscale = True
mLineStyle = ChartSkil.LineStyles.LineSolid
mChartRegionName = PriceRegionName
mLayer = ChartSkil.LayerNumbers.LayerHIghestUser
mHistogramBarWidth = 0.5
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let chartRegionName(ByVal value As String)
mChartRegionName = value
End Property

Public Property Get chartRegionName() As String
chartRegionName = mChartRegionName
End Property

Friend Property Let color(ByVal value As Long)
mColor = value
End Property

Public Property Get color() As Long
color = mColor
End Property

Friend Property Let displayMode(ByVal value As ChartSkil.DisplayModes)
mDisplayMode = value
End Property

Public Property Get displayMode() As ChartSkil.DisplayModes
displayMode = mDisplayMode
End Property

Friend Property Let histogramBarWidth(ByVal value As Single)
mHistogramBarWidth = value
End Property

Public Property Get histogramBarWidth() As Single
histogramBarWidth = mHistogramBarWidth
End Property

Friend Property Let includeInAutoscale(ByVal value As Boolean)
mIncludeInAutoscale = value
End Property

Public Property Get includeInAutoscale() As Boolean
includeInAutoscale = mIncludeInAutoscale
End Property

Friend Property Let includeInChart(ByVal value As Boolean)
mIncludeInChart = value
End Property

Public Property Get includeInChart() As Boolean
includeInChart = mIncludeInChart
End Property

Friend Property Let layer(ByVal value As Long)
mLayer = value
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Friend Property Let lineStyle(ByVal value As ChartSkil.LineStyles)
mLineStyle = value
End Property

Public Property Get lineStyle() As ChartSkil.LineStyles
lineStyle = mLineStyle
End Property

Friend Property Let lineThickness(ByVal value As Long)
mLineThickness = value
End Property

Public Property Get lineThickness() As Long
lineThickness = mLineThickness
End Property

Friend Property Let multipleValuesPerBar(ByVal value As Boolean)
mMultipleValuesPerBar = value
End Property

Public Property Get multipleValuesPerBar() As Boolean
multipleValuesPerBar = mMultipleValuesPerBar
End Property

Friend Property Let valueName(ByVal value As String)
mValueName = value
End Property

Public Property Get valueName() As String
valueName = mValueName
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

