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
Private mDataPointDisplayMode As DataPointDisplayModes
Private mLineThickness As Long
Private mIncludeInAutoscale As Boolean
Private mLineStyle As LineStyles
Private mChartRegionName As String
Private mLayer As Long
Private mHistogramBarWidth As Single

Private mBarThickness As Long
Private mBarWidth As Single
Private mBarDisplayMode As BarDisplayModes
Private mDownColor As Long
Private mOutlineThickness As Long
Private mSolidUpBody As Boolean
Private mTailThickness As Long
Private mUpColor As Long

Private mFixedX As Boolean
Private mFixedY As Boolean
Private mExtended As Boolean

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

Private mFont As StdFont
Private mBox As Boolean
Private mBoxColor As Long
Private mBoxStyle As LineStyles
Private mBoxThickness As Long
Private mBoxFillColor As Long
Private mBoxFillStyle As FillStyles
Private mPaddingX As Double
Private mPaddingY As Double
Private mAlign As TextAlignModes
Private mOffset As Dimension

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

mColor = vbBlack
mDataPointDisplayMode = DataPointDisplayModes.DataPointDisplayModeLine
mLineThickness = 1
mIncludeInAutoscale = True
mLineStyle = LineStyles.LineSolid
'mChartRegionName = RegionNamePrice
mLayer = LayerNumbers.LayerHIghestUser
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

Public Property Get align() As TextAlignModes
align = mAlign
End Property

Public Property Let align(ByVal newValue As TextAlignModes)
mAlign = newValue
End Property

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

Public Property Let barThickness(ByVal value As Long)
mBarThickness = value
End Property

Public Property Get barThickness() As Long
barThickness = mBarThickness
End Property

Friend Property Get box() As Boolean
box = mBox
End Property

Public Property Let box(ByVal newValue As Boolean)
mBox = newValue
End Property

Public Property Get boxColor() As Long
boxColor = mBoxColor
End Property

Public Property Let boxColor(ByVal newValue As Long)
mBoxColor = newValue
End Property

Friend Property Get boxFillColor() As Long
boxFillColor = mBoxFillColor
End Property

Public Property Let boxFillColor(ByVal newValue As Long)
mBoxFillColor = newValue
End Property

Public Property Get boxFillStyle() As FillStyles
boxFillStyle = mBoxFillStyle
End Property

Public Property Let boxFillStyle(ByVal newValue As FillStyles)
mBoxFillStyle = newValue
End Property

Public Property Get boxStyle() As LineStyles
boxStyle = mBoxStyle
End Property

Public Property Let boxStyle(ByVal newValue As LineStyles)
mBoxStyle = newValue
End Property

Public Property Get boxThickness() As Long
boxThickness = mBoxThickness
End Property

Public Property Let boxThickness(ByVal newValue As Long)
mBoxThickness = newValue
End Property

Public Property Let chartRegionName(ByVal value As String)
mChartRegionName = value
End Property

Public Property Get chartRegionName() As String
chartRegionName = mChartRegionName
End Property

Public Property Let color(ByVal value As Long)
mColor = value
End Property

Public Property Get color() As Long
color = mColor
End Property

Public Property Get barWidth() As Single
barWidth = mBarWidth
End Property

Public Property Let barWidth( _
                ByVal value As Single)
mBarWidth = value
End Property

Public Property Let dataPointDisplayMode(ByVal value As DataPointDisplayModes)
mDataPointDisplayMode = value
End Property

Public Property Get dataPointDisplayMode() As DataPointDisplayModes
dataPointDisplayMode = mDataPointDisplayMode
End Property

Public Property Get barDisplayMode() As BarDisplayModes
barDisplayMode = mBarDisplayMode
End Property

Public Property Let barDisplayMode( _
                ByVal value As BarDisplayModes)
mBarDisplayMode = value
End Property

Public Property Get downColor() As Long
downColor = mDownColor
End Property

Public Property Let downColor( _
                ByVal value As Long)
mDownColor = value
End Property

Public Property Get extended() As Boolean
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
mExtendAfter = newValue
End Property

Public Property Get extendBefore() As Boolean
extendBefore = mExtendBefore
End Property

Public Property Let extendBefore(ByVal newValue As Boolean)
mExtendBefore = newValue
End Property

Public Property Get fixedX() As Boolean
fixedX = mFixedX
End Property

Public Property Let fixedX(ByVal newValue As Boolean)
mFixedX = newValue
End Property

Public Property Get fixedY() As Boolean
fixedY = mFixedY
End Property

Public Property Let fixedY(ByVal newValue As Boolean)
mFixedY = newValue
End Property

Public Property Get font() As StdFont
Set font = mFont
End Property

Public Property Let font(ByVal value As StdFont)
Set mFont = value
End Property

Public Property Let histogramBarWidth(ByVal value As Single)
mHistogramBarWidth = value
End Property

Public Property Get histogramBarWidth() As Single
histogramBarWidth = mHistogramBarWidth
End Property

Public Property Let includeInAutoscale(ByVal value As Boolean)
mIncludeInAutoscale = value
End Property

Public Property Get includeInAutoscale() As Boolean
includeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let includeInChart(ByVal value As Boolean)
mIncludeInChart = value
End Property

Public Property Get includeInChart() As Boolean
includeInChart = mIncludeInChart
End Property

Public Property Let layer(ByVal value As Long)
mLayer = value
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Let lineStyle(ByVal value As LineStyles)
mLineStyle = value
End Property

Public Property Get lineStyle() As LineStyles
lineStyle = mLineStyle
End Property

Public Property Let lineThickness(ByVal value As Long)
mLineThickness = value
End Property

Public Property Get lineThickness() As Long
lineThickness = mLineThickness
End Property

Public Property Let offset(ByVal value As Dimension)
mOffset = value
End Property

Public Property Get offset() As Dimension
offset = mOffset
End Property

Public Property Get outlineThickness() As Long
outlineThickness = mOutlineThickness
End Property

Public Property Let outlineThickness( _
                ByVal value As Long)
mOutlineThickness = value
End Property

Public Property Get paddingX() As Double
paddingX = mPaddingX
End Property

Public Property Let paddingX(ByVal newValue As Double)
mPaddingX = newValue
End Property

Public Property Get paddingY() As Double
paddingY = mPaddingY
End Property

Public Property Let paddingY(ByVal newValue As Double)
mPaddingY = newValue
End Property

Public Property Get solidUpBody() As Boolean
solidUpBody = mSolidUpBody
End Property

Public Property Let solidUpBody( _
                ByVal value As Boolean)
mSolidUpBody = value
End Property

Public Property Get tailThickness() As Long
tailThickness = mTailThickness
End Property

Public Property Let tailThickness( _
                ByVal value As Long)
mTailThickness = value
End Property

Public Property Get upColor() As Long
upColor = mUpColor
End Property

Public Property Let upColor( _
                ByVal value As Long)
mUpColor = value
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


