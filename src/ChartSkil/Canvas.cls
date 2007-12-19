VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Canvas"
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

Private Type POINTAPI
        x As Long
        y As Long
End Type

'================================================================================
' Declares
'================================================================================

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, _
                                            lpPoint As POINTAPI, _
                                            ByVal nCount As Long) As Long

'================================================================================
' Member variables
'================================================================================

Private mSurface As PictureBox
Private mBoundary As TRectangle
Private mGaugeX As Double
Private mGaugeY As Double

' pen attributes
Private mPenColor As Long
Private mPenWidth As Long
Private mPenStyle As LineStyles
Private mPenMode As DrawModes

' brush attributes
Private mBrushColor As Long
Private mBrushStyle As FillStyles

' text attributes
Private mTextColor As Long
Private mTextFont As StdFont
Private mTextTransparent As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
rectInitialise mBoundary
End Sub

Private Sub Class_Terminate()
Debug.Print "Canvas terminated"
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

Friend Property Let backColor(ByVal value As Long)
mSurface.backColor = value
End Property

Public Property Get backColor() As Long
backColor = mSurface.backColor
End Property

Public Property Let bottom(ByVal value As Single)
mBoundary.bottom = value
rectValidate mBoundary
If Not mBoundary.isValid Then Exit Property

setScale
End Property

Public Property Get bottom() As Single
bottom = mBoundary.bottom
End Property

Public Property Get boundary() As TRectangle
boundary = mBoundary
End Property

Public Property Get gaugeX() As Double
gaugeX = mGaugeX
End Property

Public Property Get gaugeY() As Double
gaugeY = mGaugeY
End Property

Public Property Get hdc() As Long
hdc = mSurface.hdc
End Property

Public Property Let height(ByVal value As Single)
mBoundary.bottom = mBoundary.top - value
rectValidate mBoundary
If Not mBoundary.isValid Then Exit Property
setScale
End Property

Public Property Get height() As Single
height = mBoundary.top - mBoundary.bottom
End Property

Public Property Get heightCm() As Single
heightCm = mSurface.height / TwipsPerCm
End Property

Public Property Let left(ByVal value As Single)
mBoundary.left = value
rectValidate mBoundary
If Not mBoundary.isValid Then Exit Property
setScale
End Property

Public Property Get left() As Single
left = mBoundary.left
End Property

Public Property Let MousePointer(ByVal value As VBRUN.MousePointerConstants)
mSurface.MousePointer = value
End Property

Public Property Get MousePointer() As VBRUN.MousePointerConstants
MousePointer = mSurface.MousePointer
End Property

Friend Sub resizedY()
' notifies that the physical height of the surface has changed.
If Not mBoundary.isValid Then Exit Sub
mSurface.ScaleTop = mBoundary.top
setScale
End Sub

Public Property Let right(ByVal value As Single)
mBoundary.right = value
rectValidate mBoundary
If Not mBoundary.isValid Then Exit Property
setScale
End Property

Friend Property Let surface(ByRef value As PictureBox)
Set mSurface = value
End Property

Public Property Let top(ByVal value As Single)
mBoundary.top = value
rectValidate mBoundary
If Not mBoundary.isValid Then Exit Property
setScale
End Property

Public Property Get top() As Single
top = mBoundary.top
End Property

Public Property Get width() As Single
width = mBoundary.right - mBoundary.left
End Property

Public Property Let width(ByVal value As Single)
mBoundary.left = mBoundary.right - value
rectValidate mBoundary
If Not mBoundary.isValid Then Exit Property
mSurface.ScaleLeft = mBoundary.left
setScale
End Property

Public Property Get widthCm() As Single
widthCm = mSurface.width / TwipsPerCm
End Property

'================================================================================
' Methods
'================================================================================

Public Sub clear()
mSurface.Cls
End Sub

Public Function compareTo( _
                ByVal gaugeX As Double, _
                ByVal gaugeY As Double, _
                ByRef rect As TRectangle) As Long
If Not rect.isValid Or Not mBoundary.isValid Then
    compareTo = CanvasComparisonCodes.CantCompare
    Exit Function
End If
If gaugeX <> mGaugeX Then
    compareTo = CanvasComparisonCodes.GaugeChangedX
End If
If gaugeY <> mGaugeY Then
    compareTo = compareTo Or CanvasComparisonCodes.GaugeChangedY
End If
If rect.left <> mBoundary.left Or _
    rect.right <> mBoundary.right _
Then
    compareTo = compareTo Or CanvasComparisonCodes.BoundsChangedX
End If
If rect.bottom <> mBoundary.bottom Or _
    rect.top <> mBoundary.top _
Then
    compareTo = compareTo Or CanvasComparisonCodes.BoundsChangedY
End If

End Function

Public Function convertLogicalToCounterDistanceX(ByVal value As Single)
convertLogicalToCounterDistanceX = (mBoundary.right - value) / (mBoundary.right - mBoundary.left) * mSurface.width / TwipsPerCm
End Function

Public Function convertLogicalToCounterDistanceY(ByVal value As Single)
convertLogicalToCounterDistanceY = (mBoundary.top - value) / (mBoundary.top - mBoundary.bottom) * mSurface.height / TwipsPerCm
End Function

Public Function convertLogicalToDistanceX(ByVal value As Single)
convertLogicalToDistanceX = (value - mBoundary.left) / (mBoundary.right - mBoundary.left) * mSurface.width / TwipsPerCm
End Function

Public Function convertLogicalToDistanceY(ByVal value As Single)
convertLogicalToDistanceY = (value - mBoundary.bottom) / (mBoundary.top - mBoundary.bottom) * mSurface.height / TwipsPerCm
End Function

Public Function convertLogicalToPixelsX(ByVal value As Single)
convertLogicalToPixelsX = mSurface.ScaleX(value, vbUser, vbPixels)
End Function

Public Function convertLogicalToPixelsY(ByVal value As Single)
convertLogicalToPixelsY = -mSurface.ScaleY(value, vbUser, vbPixels)
End Function

Public Function convertLogicalToRelativeX(ByVal value As Single)
convertLogicalToRelativeX = (value - mBoundary.left) / (mBoundary.right - mBoundary.left) * 100
End Function

Public Function convertLogicalToRelativeY(ByVal value As Single)
convertLogicalToRelativeY = (value - mBoundary.bottom) / (mBoundary.top - mBoundary.bottom) * 100
End Function

Public Function convertCounterDistanceToDistanceX(ByVal value As Single)
convertCounterDistanceToDistanceX = mSurface.width / TwipsPerCm - value     ' ((100 - 100 * value * TwipsPerCm / mSurface.width) * mSurface.width) / 100 / TwipsPerCm
End Function

Public Function convertCounterDistanceToDistanceY(ByVal value As Single)
convertCounterDistanceToDistanceY = mSurface.height / TwipsPerCm - value    ' ((100 - 100 * value * TwipsPerCm / mSurface.height) * mSurface.height) * 100 / TwipsPerCm
End Function

Public Function convertCounterDistanceToLogicalX(ByVal value As Single)
convertCounterDistanceToLogicalX = mBoundary.right - ((value * TwipsPerCm) / mSurface.width * (mBoundary.right - mBoundary.left))
End Function

Public Function convertCounterDistanceToLogicalY(ByVal value As Single)
convertCounterDistanceToLogicalY = mBoundary.top - ((value * TwipsPerCm) / mSurface.height * (mBoundary.top - mBoundary.bottom))
End Function

Public Function convertCounterDistanceToRelativeX(ByVal value As Single)
convertCounterDistanceToRelativeX = 100 - 100 * value * TwipsPerCm / mSurface.width
End Function

Public Function convertCounterDistanceToRelativeY(ByVal value As Single)
convertCounterDistanceToRelativeY = 100 - 100 * value * TwipsPerCm / mSurface.height
End Function

Public Function convertDistanceToLogicalX(ByVal value As Single)
convertDistanceToLogicalX = mBoundary.left + (value * TwipsPerCm) / mSurface.width * (mBoundary.right - mBoundary.left)
End Function

Public Function convertDistanceToLogicalY(ByVal value As Single)
convertDistanceToLogicalY = mBoundary.bottom + (value * TwipsPerCm) / mSurface.height * (mBoundary.top - mBoundary.bottom)
End Function

Public Function convertDistanceToCounterDistanceX(ByVal value As Single)
convertDistanceToCounterDistanceX = mSurface.width / TwipsPerCm - value     '  ((100 - (100 * value * TwipsPerCm / mSurface.width)) * mSurface.width) / 100 / TwipsPerCm
End Function

Public Function convertDistanceToCounterDistanceY(ByVal value As Single)
convertDistanceToCounterDistanceY = mSurface.height / TwipsPerCm - value    ' ((100 - (100 * value * TwipsPerCm / mSurface.height)) * mSurface.height) / 100 / TwipsPerCm
End Function

Public Function convertDistanceToRelativeX(ByVal value As Single)
convertDistanceToRelativeX = 100 * value * TwipsPerCm / mSurface.width
End Function

Public Function convertDistanceToRelativeY(ByVal value As Single)
convertDistanceToRelativeY = 100 * value * TwipsPerCm / mSurface.height
End Function

Public Function convertPixelsToLogicalX(ByVal value As Single)
convertPixelsToLogicalX = value * mGaugeX
End Function

Public Function convertPixelsToLogicalY(ByVal value As Single)
convertPixelsToLogicalY = value * mGaugeY
End Function

Public Function convertRelativeToLogicalX(ByVal value As Single)
convertRelativeToLogicalX = mBoundary.left + (value * (mBoundary.right - mBoundary.left) / 100)
End Function

Public Function convertRelativeToLogicalY(ByVal value As Single)
convertRelativeToLogicalY = mBoundary.bottom + (value * (mBoundary.top - mBoundary.bottom) / 100)
End Function

Public Function convertRelativeToCounterDistanceX(ByVal value As Single)
convertRelativeToCounterDistanceX = ((100 - value) * mSurface.width) / 100 / TwipsPerCm
End Function

Public Function convertRelativeToCounterDistanceY(ByVal value As Single)
convertRelativeToCounterDistanceY = ((100 - value) * mSurface.height) / 100 / TwipsPerCm
End Function

Public Function convertRelativeToDistanceX(ByVal value As Single)
convertRelativeToDistanceX = (value * mSurface.width) / 100 / TwipsPerCm
End Function

Public Function convertRelativeToDistanceY(ByVal value As Single)
convertRelativeToDistanceY = (value * mSurface.height) * 100 / TwipsPerCm
End Function

Public Sub drawCircle( _
                ByVal x1 As Single, _
                ByVal y1 As Single, _
                ByVal r As Single)
applyPenSettings
applyFillSettings
mSurface.Circle (x1, y1), r
End Sub

Public Sub drawLine( _
                ByVal x1 As Single, _
                ByVal y1 As Single, _
                ByVal x2 As Single, _
                ByVal y2 As Single)
applyPenSettings
mSurface.Line (x1, y1)-(x2, y2)
End Sub

Public Sub drawPoint( _
                ByVal x1 As Single, _
                ByVal y1 As Single)
applyPenSettings
mSurface.PSet (x1, y1)
End Sub

Public Sub drawPolygon( _
                ByRef points() As Point)
Dim apiPoints() As POINTAPI
ReDim apiPoints(UBound(points)) As POINTAPI
Dim i As Long

For i = 0 To UBound(points)
    apiPoints(i).x = convertLogicalToPixelsX(points(i).XLogical - mBoundary.left)
    apiPoints(i).y = convertLogicalToPixelsY(mBoundary.top - points(i).YLogical)
Next

applyPenSettings
applyFillSettings
Polygon mSurface.hdc, apiPoints(0), UBound(apiPoints) + 1
End Sub

Public Sub drawRectangle( _
                ByVal x1 As Single, _
                ByVal y1 As Single, _
                ByVal x2 As Single, _
                ByVal y2 As Single)
applyPenSettings
applyFillSettings
mSurface.Line (x1, y1)-(x2, y2), , B
End Sub

Public Sub drawText( _
                ByVal text As String, _
                ByVal x As Single, _
                ByVal y As Single)
Dim prevFont As StdFont
Dim textLines() As String
Dim i As Long
                
setfont
    
mSurface.FontTransparent = mTextTransparent
mSurface.CurrentX = x
mSurface.CurrentY = y
mSurface.ForeColor = mTextColor
textLines = Split(text, vbCrLf)
For i = 0 To UBound(textLines)
    mSurface.Print textLines(i)
    mSurface.CurrentX = x
Next

End Sub

Public Function newDimension(ByVal x As Double, _
                        ByVal y As Double) As Dimension
Set newDimension = New Dimension
newDimension.Canvas = Me
newDimension.x = x
newDimension.y = y
End Function

Public Function newPoint(ByVal x As Double, _
                        ByVal y As Double, _
                        Optional ByVal coordSystemX As CoordinateSystems = CoordsLogical, _
                        Optional ByVal coordSystemY As CoordinateSystems = CoordsLogical) As Point
Set newPoint = New Point
newPoint.Canvas = Me
newPoint.x = x
newPoint.y = y
newPoint.CoordinateSystemX = coordSystemX
newPoint.CoordinateSystemY = coordSystemY
End Function

Public Sub setPenAttributes( _
                ByVal Color As Long, _
                ByVal width As Long, _
                ByVal style As LineStyles, _
                ByVal mode As DrawModes)
mPenColor = Color
mPenWidth = width
mPenStyle = style
mPenMode = mode
End Sub

Public Sub setTextAttributes( _
                ByVal Color As Long, _
                ByVal font As StdFont, _
                ByVal transparent As Boolean)
mTextColor = Color
Set mTextFont = font
mTextTransparent = transparent
End Sub
                
Public Sub setBrushAttributes( _
                ByVal Color As Long, _
                ByVal style As FillStyles)
mBrushColor = Color
mBrushStyle = style
End Sub
                
Public Function TextHeight(ByVal value As String) As Single
setfont
TextHeight = Abs(mSurface.TextHeight(value))
End Function

Public Function TextWidth(ByVal value As String) As Single
setfont
TextWidth = mSurface.TextWidth(value)
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub applyFillSettings()
mSurface.FillColor = mBrushColor
mSurface.FillStyle = mBrushStyle
End Sub

Private Sub applyPenSettings()
mSurface.ForeColor = mPenColor
mSurface.DrawWidth = mPenWidth
mSurface.DrawStyle = mPenStyle
mSurface.DrawMode = mPenMode
End Sub

Private Sub setfont()
If Not mTextFont Is Nothing Then
    With mSurface.font
        If mTextFont.Size < 8 Then .Size = mTextFont.Size
       .name = mTextFont.name
       .Bold = mTextFont.Bold
       .Italic = mTextFont.Italic
       .Size = mTextFont.Size
       .Charset = mTextFont.Charset
       .Strikethrough = mTextFont.Strikethrough
       .Underline = mTextFont.Underline
       .Weight = mTextFont.Weight
    End With
End If
End Sub

Private Sub setScale()
mSurface.ScaleHeight = mBoundary.bottom - mBoundary.top
mSurface.ScaleWidth = mBoundary.right - mBoundary.left + 1
mSurface.ScaleTop = mBoundary.top
mSurface.ScaleLeft = mBoundary.left
mGaugeX = (CDbl(mSurface.ScaleWidth) / CDbl(mSurface.width)) * Screen.TwipsPerPixelX
mGaugeY = -(CDbl(mSurface.ScaleHeight) / CDbl(mSurface.height)) * Screen.TwipsPerPixelY
End Sub

