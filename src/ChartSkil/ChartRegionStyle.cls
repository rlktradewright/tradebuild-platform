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
'Private mPointerStyle As PointerStyles

' indicates whether there is a grid, not whether it
' is currently visible
Private mHasGrid As Boolean
Private mHasGridText As Boolean

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

Private Sub Class_Initialize()
mGridColor = vbBlack
mGridTextColor = vbBlack
mHasGrid = False
mHasGridText = False
mAutoscale = False
mIntegerYScale = False
mYScaleQuantum = 1
mGridlineSpacingY = 1
mMinimumHeight = 1
mSnapCursorToTickBoundaries = False
ReDim mBackGradientFillColors(0) As Long
mBackGradientFillColors(0) = vbBlack
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Autoscale() As Boolean
Autoscale = mAutoscale
End Property

Public Property Let Autoscale(ByVal value As Boolean)
mAutoscale = value
End Property

Public Property Get BackColor() As Long
BackColor = mBackGradientFillColors(0)
End Property

Public Property Let BackColor(ByVal value As Long)
ReDim mBackGradientFillColors(0) As Long
mBackGradientFillColors(0) = value
End Property

Public Property Get BackGradientFillColors() As Long()
BackGradientFillColors = mBackGradientFillColors
End Property

Public Property Let BackGradientFillColors(ByRef value() As Long)
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

Public Property Get GridColor() As Long
GridColor = mGridColor
End Property

Public Property Let GridColor(ByVal val As Long)
mGridColor = val
End Property

Public Property Get GridlineSpacingY() As Double
GridlineSpacingY = mGridlineSpacingY
End Property

Public Property Let GridlineSpacingY(ByVal value As Double)
mGridlineSpacingY = value
End Property

Public Property Get GridTextColor() As Long
GridTextColor = mGridTextColor
End Property

Public Property Let GridTextColor(ByVal val As Long)
mGridTextColor = val
End Property

Public Property Get HasGrid() As Boolean
HasGrid = mHasGrid
End Property

Public Property Let HasGrid(ByVal val As Boolean)
mHasGrid = val
End Property

Public Property Get HasGridText() As Boolean
HasGridText = mHasGridText
End Property

Public Property Let HasGridText(ByVal val As Boolean)
mHasGridText = val
End Property

Public Property Get IntegerYScale() As Boolean
IntegerYScale = mIntegerYScale
End Property

Public Property Let IntegerYScale(ByVal value As Boolean)
mIntegerYScale = value
End Property

Public Property Get MinimumHeight() As Double
MinimumHeight = mMinimumHeight
End Property

Public Property Let MinimumHeight(ByVal value As Double)
mMinimumHeight = value
End Property

'Public Property Get PointerStyle() As PointerStyles
'PointerStyle = mPointerStyle
'End Property
'
'Public Property Let PointerStyle(ByVal value As PointerStyles)
'mPointerStyle = value
'End Property

Public Property Let SnapCursorToTickBoundaries(ByVal value As Boolean)
mSnapCursorToTickBoundaries = value
End Property

Public Property Get SnapCursorToTickBoundaries() As Boolean
SnapCursorToTickBoundaries = mSnapCursorToTickBoundaries
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
clone.Autoscale = mAutoscale
clone.BackGradientFillColors = mBackGradientFillColors
clone.GridColor = mGridColor
clone.GridlineSpacingY = mGridlineSpacingY
clone.GridTextColor = mGridTextColor
clone.IntegerYScale = mIntegerYScale
clone.MinimumHeight = mMinimumHeight
clone.HasGrid = mHasGrid
clone.HasGridText = mHasGridText
clone.SnapCursorToTickBoundaries = mSnapCursorToTickBoundaries
clone.YScaleQuantum = mYScaleQuantum
End Function

'@================================================================================
' Helper Functions
'@================================================================================


