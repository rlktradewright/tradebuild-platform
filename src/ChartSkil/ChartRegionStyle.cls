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

Private Const ProjectName                As String = "ChartSkil26"
Private Const ModuleName                As String = "ChartRegionStyle"

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
Private mhasGrid As Boolean
Private mhasGridText As Boolean
Private mAutoscale As Boolean
Private mIntegerYScale As Boolean
Private mYScaleQuantum As Double
Private mGridlineSpacingY As Double
Private mMinimumHeight As Double
Private mSnapCursorToTickBoundaries As Boolean

Private mBackGradientFillColors() As Long

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
backColor = mBackGradientFillColors(0)
End Property

Public Property Let backColor(ByVal value As Long)
ReDim mBackGradientFillColors(0) As Long
mBackGradientFillColors(0) = value
End Property

Public Property Get backGradientFillColors() As Long()
backGradientFillColors = mBackGradientFillColors
End Property

Public Property Let backGradientFillColors(ByRef value() As Long)
Dim numColors As Long

On Error Resume Next
numColors = UBound(value) + 1
On Error GoTo 0

If numColors = 0 Then
    Erase mBackGradientFillColors
    Exit Property
End If

If numColors < 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "BackGradientFillColors", _
            "At least 1 background gradient fill color must be supplied"
End If

If numColors > 5 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "BackGradientFillColors", _
            "At most 5 background gradient fill colors may be supplied"
End If

mBackGradientFillColors = value
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

Public Property Let snapCursorToTickBoundaries(ByVal value As Boolean)
mSnapCursorToTickBoundaries = value
End Property

Public Property Get snapCursorToTickBoundaries() As Boolean
snapCursorToTickBoundaries = mSnapCursorToTickBoundaries
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
clone.backGradientFillColors = mBackGradientFillColors
clone.gridColor = mGridColor
clone.gridlineSpacingY = mGridlineSpacingY
clone.gridTextColor = mGridTextColor
clone.integerYScale = mIntegerYScale
clone.minimumHeight = mMinimumHeight
clone.pointerStyle = mPointerStyle
clone.hasGrid = mhasGrid
clone.hasGridText = mhasGridText
clone.snapCursorToTickBoundaries = mSnapCursorToTickBoundaries
clone.YScaleQuantum = mYScaleQuantum
End Function

'@================================================================================
' Helper Functions
'@================================================================================


