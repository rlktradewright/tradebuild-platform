VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartRegionStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
' @see
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mGridColor As Long
Private mGridTextColor  As Long
Private mPointerStyle As PointerStyles
Private mBackColor As Long
Private mhasGrid As Boolean
Private mhasGridText As Boolean
Private mAutoscale As Boolean
Private mIntegerYScale As Boolean
Private mYScaleQuantum As Double
Private mGridlineSpacingY As Double
Private mMinimumHeight As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get autoscale() As Boolean
autoscale = mAutoscale
End Property

Public Property Let autoscale(ByVal value As Boolean)
mAutoscale = value
End Property

Public Property Get backColor() As Long
backColor = mBackColor
End Property

Public Property Let backColor(ByVal val As Long)
mBackColor = val
End Property

Public Property Get gridColor() As Long
gridColor = mGridColor
End Property

Public Property Let gridColor(ByVal val As Long)
mGridColor = val
End Property

Public Property Get gridlineSpacingY() As Double
gridlineSpacingY = mGridlineSpacingY
End Property

Public Property Let gridlineSpacingY(ByVal value As Double)
mGridlineSpacingY = value
End Property

Public Property Get gridTextColor() As Long
gridTextColor = mGridTextColor
End Property

Public Property Let gridTextColor(ByVal val As Long)
mGridTextColor = val
End Property

Public Property Get hasGrid() As Boolean
hasGrid = mhasGrid
End Property

Public Property Let hasGrid(ByVal val As Boolean)
mhasGrid = val
End Property

Public Property Get hasGridText() As Boolean
hasGridText = mhasGridText
End Property

Public Property Let hasGridText(ByVal val As Boolean)
mhasGridText = val
End Property

Public Property Get integerYScale() As Boolean
integerYScale = mIntegerYScale
End Property

Public Property Let integerYScale(ByVal value As Boolean)
mIntegerYScale = value
End Property

Public Property Get minimumHeight() As Double
minimumHeight = mMinimumHeight
End Property

Public Property Let minimumHeight(ByVal value As Double)
mMinimumHeight = value
End Property

Public Property Get pointerStyle() As PointerStyles
pointerStyle = mPointerStyle
End Property

Public Property Let pointerStyle(ByVal value As PointerStyles)
mPointerStyle = value
End Property

Public Property Get YScaleQuantum() As Double
YScaleQuantum = mYScaleQuantum
End Property

Public Property Let YScaleQuantum(ByVal value As Double)
mYScaleQuantum = value
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As ChartRegionStyle
Set clone = New ChartRegionStyle
clone.autoscale = mAutoscale
clone.backColor = mBackColor
clone.gridColor = mGridColor
clone.gridlineSpacingY = mGridlineSpacingY
clone.gridTextColor = mGridTextColor
clone.integerYScale = mIntegerYScale
clone.minimumHeight = mMinimumHeight
clone.pointerStyle = mPointerStyle
clone.hasGrid = mhasGrid
clone.hasGridText = mhasGridText
clone.YScaleQuantum = mYScaleQuantum
End Function

'@================================================================================
' Helper Functions
'@================================================================================


