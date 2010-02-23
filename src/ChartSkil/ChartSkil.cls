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

Public Enum GraphicObjectCapabilities
    ComplexBounding = 1
    PartialRedraw = 2
    BlockUndraw = 4
End Enum

Public Enum CoordinateSystems
    CoordsLogical = 0
    CoordsRelative
    CoordsDistance        ' Measured from Left or Bottom of region
    CoordsCounterDistance ' Measured from Right or Top of region
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

Public Enum EllipsisModes
    EllipsisNone = 0
    EllipsisEnd = twwin32api.DT_END_ELLIPSIS
    EllipsisPath = twwin32api.DT_PATH_ELLIPSIS
    EllipsisWord = twwin32api.DT_WORD_ELLIPSIS
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
    LayerMin = LayerNumberRange.MinLayer
    LayerBackground = LayerNumberRange.MinLayer
    LayerGrid = LayerNumberRange.MinLayer + 5
    LayerGridText = LayerGrid + 3
    LayerLowestUser = LayerNumberRange.MinLayer + 10
'    LayerDefaultBars
'    LayerDefaultDataPoints
'    LayerDefaultLines
'    LayerDefaultTexts
    LayerHighestUser = LayerNumberRange.MaxLayer - 10
    LayerTitle = LayerNumberRange.MaxLayer - 5
    LayerPointer = LayerNumberRange.MaxLayer - 2
    LayerMax = LayerNumberRange.MaxLayer
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

Public Enum PointerModes
    PointerModeDefault
    PointerModeTool
    PointerModeSelection
End Enum

Public Enum PointerStyles
    PointerNone
    PointerCrosshairs
    PointerDisc
    PointerTool
    PointerCustom
    PointerSelection
    PointerWindowsDefault
End Enum

Public Enum PointStyles
    PointRound
    PointSquare
End Enum

Public Enum Quadrants
    NE
    NW
    SW
    SE
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

Public Enum TextJustifyModes
    JustifyLeft = twwin32api.DT_LEFT
    JustifyRight = twwin32api.DT_RIGHT
    JustifyCentre = twwin32api.DT_CENTER
    JustifyTop = twwin32api.DT_TOP
    JustifyVCentre = twwin32api.DT_VCENTER
    JustifyBottom = twwin32api.DT_BOTTOM
End Enum

'Public Enum ToolTypes
'    ToolPointer
'    ToolLine
'    ToolLineExtended
'    ToolLineRay
'    ToolLineHorizontal
'    ToolLineVertical
'    ToolFibonacciRetracement
'    ToolFibonacciExtension
'    ToolFibonacciCircle
'    ToolFibonacciTime
'    ToolRegressionChannel
'    ToolRegressionEnvelope
'    ToolText
'    ToolPitchfork
'    ToolCircle
'    ToolRectangle
'End Enum

Public Enum Verticals
    VerticalNot
    VerticalUp
    VerticalDown
End Enum

Public Enum ViewportComparisonCodes
    CantCompare = 1
    GaugeChangedX = 2
    GaugeChangedY = 4
    BoundsChangedX = 8
    BoundsChangedY = 16
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type TInterval
    isValid         As Boolean
    startValue      As Double
    endValue        As Double
End Type

Public Type TPoint
    X               As Double
    Y               As Double
End Type

Public Type TRectangle
    Bottom          As Double
    Left            As Double
    Right           As Double
    Top             As Double
    isValid         As Boolean
End Type

'@================================================================================
' Constants
'@================================================================================

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

Public Function DegreesToRadians( _
                ByVal degrees As Double) As Double
DegreesToRadians = gDegreesToRadians(degrees)
End Function

Public Function RadiansToDegrees( _
                ByVal radians As Double) As Double
RadiansToDegrees = gRadiansToDegrees(radians)
End Function

'@================================================================================
' Helper Functions
'@================================================================================


