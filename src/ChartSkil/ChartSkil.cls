VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartSkil"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Public Enum ArrowStyles
    ArrowNone
    ArrowSingleOpen
    ArrowDoubleOpen
    ArrowClosed
    ArrowSingleBar
    ArrowDoubleBar
    ArrowLollipop
    ArrowDiamond
    ArrowBarb
End Enum

Public Enum BarDisplayModes
    BarDisplayModeBar
    BarDisplayModeCandlestick
    BarDisplayModeLine
End Enum

Public Enum CanvasComparisonCodes
    CantCompare = 1
    GaugeChangedX = 2
    GaugeChangedY = 4
    BoundsChangedX = 8
    BoundsChangedY = 16
End Enum

Public Enum CoordinateSystems
    CoordsLogical = 0
    CoordsRelative
    CoordsDistance        ' Measured from left or bottom of region
    CoordsCounterDistance ' Measured from right or top of region
End Enum

Public Enum DataPointDisplayModes
    DataPointDisplayModePoint = 1
    DataPointDisplayModeLine
    DataPointDisplayModeStep
    DataPointDisplayModeHistogram
End Enum

Public Enum DrawModes
    DrawModeBlackness = vbBlackness
    DrawModeCopyPen = vbCopyPen
    DrawModeInvert = vbInvert
    DrawModeMaskNotPen = vbMaskNotPen
    DrawModeMaskPen = vbMaskPen
    DrawModeMaskPenNot = vbMaskPenNot
    DrawModeMergeNotPen = vbMergeNotPen
    DrawModeMergePen = vbMergePen
    DrawModeMergePenNot = vbMergePenNot
    DrawModeNop = vbNop
    DrawModeNotCopyPen = vbNotCopyPen
    DrawModeNotMaskPen = vbNotMaskPen
    DrawModeNotMergePen = vbNotMergePen
    DrawModeNotXorPen = vbNotXorPen
    DrawModeWhiteness = vbWhiteness
    DrawModeXorPen = vbXorPen
End Enum

Public Enum FillStyles
    FillSolid = vbFSSolid ' 0 Solid
    FillTransparent = vbFSTransparent ' 1 (Default) Transparent
    FillHorizontalLine = vbHorizontalLine ' 2 Horizontal Line
    FillVerticalLine = vbVerticalLine ' 3 Vertical Line
    FillUpwardDiagonal = vbUpwardDiagonal ' 4 Upward Diagonal
    FillDownwardDiagonal = vbDownwardDiagonal ' 5 Downward Diagonal
    FillCross = vbCross ' 6 Cross
    FillDiagonalCross = vbDiagonalCross ' 7 Diagonal Cross
End Enum

Public Enum LayerNumbers
    LayerBackground = MinLayer
    LayerGrid = MinLayer + 5
    LayerGridText = LayerGrid + 3
    LayerLowestUser = MinLayer + 10
    LayerHighestUser = MaxLayer - 10
    LayerDefaultLines
    LayerDefaultTexts
    LayerTitle = MaxLayer - 5
    LayerPointer = MaxLayer
End Enum

Public Enum LineStyles
    LineSolid = vbSolid
    LineDash = vbDash
    LineDot = vbDot
    LineDashDot = vbDashDot
    LineDashDotDot = vbDashDotDot
    LineInvisible = vbInvisible
    LineInsideSolid = vbInsideSolid
End Enum

Public Enum PointerStyles
    PointerNone
    PointerCrosshairs
    PointerDisc
End Enum

Public Enum PointStyles
    PointRound
    PointSquare
End Enum

Public Enum TextAlignModes
    AlignTopLeft
    AlignCentreLeft
    AlignBottomLeft
    AlignTopCentre
    AlignCentreCentre
    AlignBottomCentre
    AlignTopRight
    AlignCentreRight
    AlignBottomRight
    AlignBoxTopLeft
    AlignBoxCentreLeft
    AlignBoxBottomLeft
    AlignBoxTopCentre
    AlignBoxCentreCentre
    AlignBoxBottomCentre
    AlignBoxTopRight
    AlignBoxCentreRight
    AlignBoxBottomRight
End Enum

Public Enum ToolTypes
    ToolPointer
    ToolLine
    ToolLineExtended
    ToolLineRay
    ToolLineHorizontal
    ToolLineVertical
    ToolFibonacciRetracement
    ToolFibonacciExtension
    ToolFibonacciCircle
    ToolFibonacciTime
    ToolRegressionChannel
    ToolRegressionEnvelope
    ToolText
    ToolPitchfork
    ToolCircle
    ToolRectangle
End Enum

Public Enum Verticals
    VerticalNot
    VerticalUp
    VerticalDown
End Enum

Public Enum Quadrants
    NE
    NW
    SW
    SE
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type TInterval
    isValid         As Boolean
    startValue      As Double
    endValue        As Double
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "ChartSkil26"
Private Const ModuleName                    As String = "ChartSkil"

'@================================================================================
' Member variables
'@================================================================================

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

'@================================================================================
' Methods
'@================================================================================

Public Function CreateBarStyle() As BarStyle
Set CreateBarStyle = gCreateBarStyle
End Function

Public Function CreateDataPointStyle() As DataPointStyle
Set CreateDataPointStyle = gCreateDataPointStyle
End Function

Public Function CreateLineStyle() As linestyle
Set CreateLineStyle = gCreateLineStyle
End Function

Public Function CreateChartRegionStyle() As ChartRegionStyle
Set CreateChartRegionStyle = gCreateChartRegionStyle
End Function

Public Function CreateTextStyle() As TextStyle
Set CreateTextStyle = gCreateTextStyle
End Function

'@================================================================================
' Helper Functions
'@================================================================================


