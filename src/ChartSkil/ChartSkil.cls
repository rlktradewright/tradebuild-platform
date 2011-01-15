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

Public Enum BarPropertyFlags
    BarPropertyColor = 1
    BarPropertyUpColor = 2
    BarPropertyDownColor = 4
    BarPropertyDisplayMode = 8
    BarPropertySolidUpBody = &H10&
    BarPropertyThickness = &H20&
    BarPropertyWidth = &H40&
    BarPropertyTailThickness = &H80&
    BarPropertyOutlineThickness = &H100&
    BarPropertyIncludeInAutoscale = &H200&
    BarPropertyLayer = &H400&
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

Public Enum DataPointPropertyFlags
    DataPointPropertyLineThickness = 1
    DataPointPropertyColor = 2
    DataPointPropertyUpColor = 4
    DataPointPropertyDownColor = 8
    DataPointPropertyLineStyle = &H10&
    DataPointPropertyPointStyle = &H20&
    DataPointPropertyDisplayMode = &H40&
    DataPointPropertyHistWidth = &H80&
    DataPointPropertyIncludeInAutoscale = &H100&
    DataPointPropertyLayer = &H200&
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
    EllipsisEnd = TWWin32API.DT_END_ELLIPSIS
    EllipsisPath = TWWin32API.DT_PATH_ELLIPSIS
    EllipsisWord = TWWin32API.DT_WORD_ELLIPSIS
End Enum

Public Enum FillStyles
    FillSolid = vbFSSolid                       ' 0 Solid
    FillTransparent = vbFSTransparent           ' 1 (Default) Transparent
    FillHorizontalLine = vbHorizontalLine       ' 2 Horizontal Line
    FillVerticalLine = vbVerticalLine           ' 3 Vertical Line
    FillUpwardDiagonal = vbUpwardDiagonal       ' 4 Upward Diagonal
    FillDownwardDiagonal = vbDownwardDiagonal   ' 5 Downward Diagonal
    FillCross = vbCross                         ' 6 Cross
    FillDiagonalCross = vbDiagonalCross         ' 7 Diagonal Cross
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

Public Enum LinePropertyFlags
    LinePropertyColor = 1
    LinePropertyThickness = 2
    LinePropertyLineStyle = 4
    LinePropertyExtendBefore = 8
    LinePropertyExtendAfter = &H10&
    LinePropertyArrowStartStyle = &H20&
    LinePropertyArrowStartLength = &H40&
    LinePropertyArrowStartWidth = &H80&
    LinePropertyArrowStartColor = &H100&
    LinePropertyArrowStartFillColor = &H200&
    LinePropertyArrowStartFillStyle = &H400&
    LinePropertyArrowEndStyle = &H800&
    LinePropertyArrowEndLength = &H1000&
    LinePropertyArrowEndWidth = &H2000&
    LinePropertyArrowEndColor = &H4000&
    LinePropertyArrowEndFillColor = &H8000&
    LinePropertyArrowEndFillStyle = &H10000
    LinePropertyFixedX = &H20000
    LinePropertyFixedY = &H40000
    LinePropertyIncludeInAutoscale = &H80000
    LinePropertyExtended = &H100000
    LinePropertyLayer = &H200000
    LinePropertyOffset1 = &H400000
    LinePropertyOffset2 = &H800000
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
    JustifyLeft = TWWin32API.DT_LEFT
    JustifyRight = TWWin32API.DT_RIGHT
    JustifyCentre = TWWin32API.DT_CENTER
    JustifyTop = TWWin32API.DT_TOP
    JustifyVCentre = TWWin32API.DT_VCENTER
    JustifyBottom = TWWin32API.DT_BOTTOM
End Enum

Public Enum TextPropertyFlags
    TextPropertyColor = 1
    TextPropertyBox = 2
    TextPropertyBoxColor = 4
    TextPropertyBoxStyle = 8
    TextPropertyBoxThickness = &H10&
    TextPropertyBoxFillColor = &H20&
    TextPropertyBoxFillStyle = &H40&
    TextPropertyAlign = &H80&
    TextPropertyPaddingX = &H100&
    TextPropertyPaddingY = &H200&
    TextPropertyFont = &H400&
    TextPropertyBoxFillWithBackgroundColor = &H800&
    TextPropertyFixedX = &H1000
    TextPropertyFixedY = &H2000
    TextPropertyIncludeInAutoscale = &H4000
    TextPropertyExtended = &H8000
    TextPropertyLayer = &H10000
    TextPropertySize = &H20000
    TextPropertyAngle = &H40000
    TextPropertyJustification = &H80000
    TextPropertyMultiLine = &H100000
    TextPropertyEllipsis = &H200000
    TextPropertyExpandTabs = &H400000
    TextPropertyTabWidth = &H800000
    TextPropertyWordWrap = &H1000000
    TextPropertyLeftMargin = &H2000000
    TextPropertyRightMargin = &H4000000
    TextPropertyOffset = &H8000000
    TextPropertyHideIfBlank = &H10000000
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

Public Property Get GraphicObjectStyleManager() As GraphicObjectStyleManager
Set GraphicObjectStyleManager = gGraphicObjectStyleManager
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function DegreesToRadians( _
                ByVal degrees As Double) As Double
DegreesToRadians = gDegreesToRadians(degrees)
End Function

Public Function GetDefaultBarStyle() As BarStyle
Set GetDefaultBarStyle = gDefaultBarStyle
End Function

Public Function GetDefaultChartRegionStyle() As ChartRegionStyle
Set GetDefaultChartRegionStyle = gDefaultChartRegionStyle
End Function

Public Function GetDefaultDataPointStyle() As DataPointStyle
Set GetDefaultDataPointStyle = gDefaultDataPointStyle
End Function

Public Function GetDefaultLineStyle() As LineStyle
Set GetDefaultLineStyle = gDefaultLineStyle
End Function

Public Function GetDefaultTextStyle() As TextStyle
Set GetDefaultTextStyle = gDefaultTextStyle
End Function

Public Function LoadDimensionFromConfig( _
                ByVal pConfig As ConfigurationSection) As Dimension
Const ProcName As String = "LoadDimensionFromConfig"
On Error GoTo Err

Set LoadDimensionFromConfig = gLoadDimensionFromConfig(pConfig)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function LoadSizeFromConfig( _
                ByVal pConfig As ConfigurationSection) As Size
Const ProcName As String = "LoadSizeFromConfig"
On Error GoTo Err

Set LoadSizeFromConfig = gLoadSizeFromConfig(pConfig)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function NewDimension( _
                ByVal pLength As Double, _
                Optional ByVal pCoordSystem As CoordinateSystems = CoordsDistance) As Dimension
Const ProcName As String = "NewDimension"
Dim failpoint As String
On Error GoTo Err

Set NewDimension = gNewDimension(pLength, pCoordSystem)
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function NewPoint( _
                ByVal X As Double, _
                ByVal Y As Double, _
                Optional ByVal pCoordSystemX As CoordinateSystems = CoordsLogical, _
                Optional ByVal pCoordSystemY As CoordinateSystems = CoordsLogical, _
                Optional ByVal Offset As Size) As Point
Const ProcName As String = "NewPoint"
On Error GoTo Err

Set NewPoint = gNewPoint(X, Y, pCoordSystemX, pCoordSystemY)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function NewSize( _
                ByVal X As Double, _
                ByVal Y As Double, _
                Optional ByVal coordSystemX As CoordinateSystems = CoordsDistance, _
                Optional ByVal coordSystemY As CoordinateSystems = CoordsDistance) As Size
Const ProcName As String = "NewSize"
Dim failpoint As String
On Error GoTo Err

Set NewSize = gNewSize(X, Y, coordSystemX, coordSystemY)
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function RadiansToDegrees( _
                ByVal radians As Double) As Double
RadiansToDegrees = gRadiansToDegrees(radians)
End Function

'@================================================================================
' Helper Functions
'@================================================================================


