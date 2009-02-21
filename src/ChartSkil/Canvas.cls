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

Private Const ProjectName                As String = "ChartSkil26"
Private Const ModuleName                As String = "Canvas"

Private Const GRADIENT_FILL_RECT_H As Long = 0
Private Const GRADIENT_FILL_RECT_V As Long = 1
Private Const GRADIENT_FILL_TRIANGLE  As Long = 2

Private Const NULLREGION = 1
Private Const SIMPLEREGION = 2
Private Const COMPLEXREGION = 3

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type GRADIENT_RECT
    upperLeft                   As Long
    lowerRight                  As Long
End Type

Private Type GRADIENT_TRIANGLE
    vertex1   As Long
    vertex2   As Long
    vertex3   As Long
End Type

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type rect
    left                As Long
    top                 As Long
    right               As Long
    bottom              As Long
End Type

Private Type TRIVERTEX
  x         As Long
  y         As Long
  Filler1   As Byte
  Red       As Byte
  Filler2   As Byte
  Green     As Byte
  Filler3   As Byte
  Blue      As Byte
  Filler4   As Byte
  Alpha     As Byte
End Type

'================================================================================
' External FUnction Declarations
'================================================================================

Private Declare Function CreateRectRgn Lib "gdi32" ( _
                ByVal x1 As Long, _
                ByVal y1 As Long, _
                ByVal x2 As Long, _
                ByVal y2 As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" ( _
                ByVal crColor As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" ( _
                ByVal hObject As Long) As Long

Private Declare Function FillRect Lib "user32" ( _
                ByVal hdc As Long, _
                lpRect As rect, _
                ByVal hBrush As Long) As Long

Private Declare Function GetBkColor Lib "gdi32" ( _
                ByVal hdc As Long) As Long
                
Private Declare Function GetClipRgn Lib "gdi32" ( _
                ByVal hdc As Long, _
                ByVal hRgn As Long) As Long
                
Private Declare Sub GradientFill Lib "msimg32" ( _
                ByVal hdc As Long, _
                ByVal pVertices As Long, _
                ByVal numVertices As Long, _
                ByVal pMesh As Long, _
                ByVal numMesh As Long, _
                ByVal dwMode As Long)

Private Declare Function Polygon Lib "gdi32" ( _
                ByVal hdc As Long, _
                lpPoint As POINTAPI, _
                ByVal nCount As Long) As Long

Private Declare Function SelectClipRgn Lib "gdi32" ( _
                ByVal hdc As Long, _
                ByVal hRgn As Long) As Long

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

Private mClipRegion As Long

Private mGradientFillMode As Long
Private mBackGradientFillColors() As Long
Private mVertices() As TRIVERTEX
Private mMeshTriangle() As GRADIENT_TRIANGLE
Private mMeshRect() As GRADIENT_RECT

'================================================================================
' Class Event Handlers
'================================================================================

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

Friend Property Let BackColor(ByVal value As Long)
ReDim mBackGradientFillColors(0) As Long
mBackGradientFillColors(0) = value
calculateGradientFillParameters
End Property

Public Property Get BackColor() As Long
BackColor = mBackGradientFillColors(0)
End Property

Friend Property Let bottom(ByVal value As Single)
mBoundary.bottom = value
rectValidate mBoundary
If Not mBoundary.isValid Then Exit Property

setScale
End Property

Public Property Get bottom() As Single
If mBoundary.bottom = PlusInfinityDouble Then
    bottom = PlusInfinitySingle
ElseIf mBoundary.bottom = MinusInfinityDouble Then
    bottom = MinusInfinitySingle
Else
    bottom = mBoundary.bottom
End If
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

Public Property Get gradientFillColors() As Long()
gradientFillColors = mBackGradientFillColors
End Property

Friend Property Let gradientFillColors(ByRef value() As Long)
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
            ProjectName & "." & ModuleName & ":" & "gradientFillColors", _
            "At least 1 background gradient fill color must be supplied"
End If

If numColors > 5 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "gradientFillColors", _
            "At most 5 background gradient fill colors may be supplied"
End If

mBackGradientFillColors = value

calculateGradientFillParameters
End Property

Public Property Get hdc() As Long
hdc = mSurface.hdc
End Property

Friend Property Let height(ByVal value As Single)
mBoundary.bottom = mBoundary.top - value
rectValidate mBoundary
If Not mBoundary.isValid Then Exit Property
setScale
End Property

Public Property Get height() As Single
If mBoundary.top = PlusInfinityDouble Or mBoundary.bottom = PlusInfinityDouble Then
    height = PlusInfinitySingle
ElseIf mBoundary.top = MinusInfinityDouble Or mBoundary.bottom = MinusInfinityDouble Then
    height = MinusInfinitySingle
Else
    height = mBoundary.top - mBoundary.bottom
End If
End Property

Public Property Get heightCm() As Single
heightCm = mSurface.height / TwipsPerCm
End Property

Friend Property Let left(ByVal value As Single)
mBoundary.left = value
rectValidate mBoundary
If Not mBoundary.isValid Then Exit Property
setScale
End Property

Public Property Get left() As Single
If mBoundary.left = PlusInfinityDouble Then
    left = PlusInfinitySingle
ElseIf mBoundary.left = MinusInfinityDouble Then
    left = MinusInfinitySingle
Else
    left = mBoundary.left
End If
End Property

Friend Property Let MouseIcon(ByVal value As IPictureDisp)
mSurface.MouseIcon = value
End Property

Friend Property Let MousePointer(ByVal value As VBRUN.MousePointerConstants)
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

Friend Property Let right(ByVal value As Single)
mBoundary.right = value
rectValidate mBoundary
If Not mBoundary.isValid Then Exit Property
setScale
End Property

Public Property Get right() As Single
If mBoundary.right = PlusInfinityDouble Then
    right = PlusInfinitySingle
ElseIf mBoundary.right = MinusInfinityDouble Then
    right = MinusInfinitySingle
Else
    right = mBoundary.right
End If
End Property

Friend Property Let surface(ByRef value As PictureBox)
rectInitialise mBoundary
Set mSurface = value
ReDim mBackGradientFillColors(0) As Long
mBackGradientFillColors(0) = vbWhite
left = 0
right = 1
bottom = 0
top = 1
End Property

Friend Property Let top(ByVal value As Single)
mBoundary.top = value
rectValidate mBoundary
If Not mBoundary.isValid Then Exit Property
setScale
End Property

Public Property Get top() As Single
If mBoundary.top = PlusInfinityDouble Then
    top = PlusInfinitySingle
ElseIf mBoundary.top = MinusInfinityDouble Then
    top = MinusInfinitySingle
Else
    top = mBoundary.top
End If
End Property

Public Property Get width() As Single
If mBoundary.right = PlusInfinityDouble Or mBoundary.left = PlusInfinityDouble Then
    width = PlusInfinitySingle
ElseIf mBoundary.right = MinusInfinityDouble Or mBoundary.left = MinusInfinityDouble Then
    width = MinusInfinitySingle
Else
    width = mBoundary.right - mBoundary.left
End If
End Property

Friend Property Let width(ByVal value As Single)
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

Friend Sub clear()
Dim failpoint As Long
On Error GoTo Err

mSurface.Cls

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "clear" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub clearClippingRegion()
Dim failpoint As Long
On Error GoTo Err

SelectClipRgn mSurface.hdc, 0
DeleteObject mClipRegion

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "clearClippingRegion" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Function compareTo( _
                ByVal gaugeX As Double, _
                ByVal gaugeY As Double, _
                ByRef rect As TRectangle) As Long
Dim failpoint As Long
On Error GoTo Err

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

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "compareTo" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription



End Function

Public Function convertCounterDistanceToDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertCounterDistanceToDistanceX = CDbl(mSurface.width) / TwipsPerCm - value

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertCounterDistanceToDistanceX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertCounterDistanceToDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertCounterDistanceToDistanceY = CDbl(mSurface.height) / TwipsPerCm - value    ' ((100 - 100 * value * TwipsPerCm / mSurface.height) * mSurface.height) * 100 / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertCounterDistanceToDistanceY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertCounterDistanceToLogicalX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertCounterDistanceToLogicalX = mBoundary.right - ((value * TwipsPerCm) / CDbl(mSurface.width) * (mBoundary.right - mBoundary.left))

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertCounterDistanceToLogicalX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertCounterDistanceToLogicalY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertCounterDistanceToLogicalY = mBoundary.top - ((value * TwipsPerCm) / CDbl(mSurface.height) * (mBoundary.top - mBoundary.bottom))

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertCounterDistanceToLogicalY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertCounterDistanceToRelativeX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertCounterDistanceToRelativeX = 100# - 100# * value * TwipsPerCm / CDbl(mSurface.width)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertCounterDistanceToRelativeX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertCounterDistanceToRelativeY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertCounterDistanceToRelativeY = 100# - 100# * value * TwipsPerCm / CDbl(mSurface.height)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertCounterDistanceToRelativeY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertDistanceToLogicalX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertDistanceToLogicalX = mBoundary.left + (value * TwipsPerCm) / CDbl(mSurface.width) * (mBoundary.right - mBoundary.left)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertDistanceToLogicalX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertDistanceToLogicalY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertDistanceToLogicalY = mBoundary.bottom + (value * TwipsPerCm) / CDbl(mSurface.height) * (mBoundary.top - mBoundary.bottom)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertDistanceToLogicalY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertDistanceToCounterDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertDistanceToCounterDistanceX = CDbl(mSurface.width) / TwipsPerCm - value

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertDistanceToCounterDistanceX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertDistanceToCounterDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertDistanceToCounterDistanceY = CDbl(mSurface.height) / TwipsPerCm - value

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertDistanceToCounterDistanceY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertDistanceToRelativeX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertDistanceToRelativeX = 100# * value * TwipsPerCm / CDbl(mSurface.width)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertDistanceToRelativeX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertDistanceToRelativeY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertDistanceToRelativeY = 100# * value * TwipsPerCm / CDbl(mSurface.height)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertDistanceToRelativeY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertLogicalToCounterDistanceX(ByVal value As Single)
Dim failpoint As Long
On Error GoTo Err

convertLogicalToCounterDistanceX = (mBoundary.right - value) / (mBoundary.right - mBoundary.left) * mSurface.width / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertLogicalToCounterDistanceX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertLogicalToCounterDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertLogicalToCounterDistanceY = (mBoundary.top - value) / (mBoundary.top - mBoundary.bottom) * CDbl(mSurface.height) / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertLogicalToCounterDistanceY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertLogicalToDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertLogicalToDistanceX = (value - mBoundary.left) / (mBoundary.right - mBoundary.left) * CDbl(mSurface.width) / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertLogicalToDistanceX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertLogicalToDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertLogicalToDistanceY = (value - mBoundary.bottom) / (mBoundary.top - mBoundary.bottom) * CDbl(mSurface.height) / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertLogicalToDistanceY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertLogicalToPixelsX(ByVal value As Double) As Long
Dim failpoint As Long
On Error GoTo Err

convertLogicalToPixelsX = mSurface.ScaleX(value, vbUser, vbPixels)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertLogicalToPixelsX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertLogicalToPixelsY(ByVal value As Double) As Long
Dim failpoint As Long
On Error GoTo Err

convertLogicalToPixelsY = -mSurface.ScaleY(value, vbUser, vbPixels)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertLogicalToPixelsY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertLogicalToRelativeX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertLogicalToRelativeX = (value - mBoundary.left) / (mBoundary.right - mBoundary.left) * 100#

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertLogicalToRelativeX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertLogicalToRelativeY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertLogicalToRelativeY = (value - mBoundary.bottom) / (mBoundary.top - mBoundary.bottom) * 100#

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertLogicalToRelativeY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertPixelsToLogicalX(ByVal value As Long) As Double
Dim failpoint As Long
On Error GoTo Err

convertPixelsToLogicalX = CDbl(value) * mGaugeX

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertPixelsToLogicalX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertPixelsToLogicalY(ByVal value As Long) As Double
Dim failpoint As Long
On Error GoTo Err

convertPixelsToLogicalY = CDbl(value) * mGaugeY

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertPixelsToLogicalY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertRelativeToLogicalX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertRelativeToLogicalX = mBoundary.left + (value * (mBoundary.right - mBoundary.left) / 100#)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertRelativeToLogicalX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertRelativeToLogicalY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertRelativeToLogicalY = mBoundary.bottom + (value * (mBoundary.top - mBoundary.bottom) / 100#)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertRelativeToLogicalY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertRelativeToCounterDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertRelativeToCounterDistanceX = ((100# - value) * CDbl(mSurface.width)) / 100# / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertRelativeToCounterDistanceX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertRelativeToCounterDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertRelativeToCounterDistanceY = ((100# - value) * CDbl(mSurface.height)) / 100# / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertRelativeToCounterDistanceY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertRelativeToDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertRelativeToDistanceX = (value * CDbl(mSurface.width)) / 100# / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertRelativeToDistanceX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function convertRelativeToDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

convertRelativeToDistanceY = (value * CDbl(mSurface.height)) * 100# / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "convertRelativeToDistanceY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Sub drawCircle( _
                ByVal centre As Point, _
                ByVal radius As Double)
Dim failpoint As Long
On Error GoTo Err

applyPenSettings
applyFillSettings
mSurface.Circle (centre.XLogical, centre.YLogical), radius

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "drawCircle" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub drawLine( _
                ByVal startPoint As Point, _
                ByVal endPoint As Point)
Dim failpoint As Long
On Error GoTo Err

applyPenSettings
mSurface.Line (startPoint.XLogical, startPoint.YLogical)-(endPoint.XLogical, endPoint.YLogical)

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "drawLine" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub drawPoint( _
                ByVal p As Point)
Dim failpoint As Long
On Error GoTo Err

applyPenSettings
mSurface.PSet (p.XLogical, p.YLogical)

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "drawPoint" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub drawPolygon( _
                ByRef points() As Point)
Dim apiPoints() As POINTAPI
Dim failpoint As Long
On Error GoTo Err

ReDim apiPoints(UBound(points)) As POINTAPI
Dim i As Long

For i = 0 To UBound(points)
    apiPoints(i).x = convertLogicalToPixelsX(points(i).XLogical - mBoundary.left)
    apiPoints(i).y = convertLogicalToPixelsY(mBoundary.top - points(i).YLogical)
Next

applyPenSettings
applyFillSettings
Polygon mSurface.hdc, apiPoints(0), UBound(apiPoints) + 1

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "drawPolygon" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub drawRectangle( _
                ByVal p1 As Point, _
                ByVal p2 As Point)
Dim failpoint As Long
On Error GoTo Err

applyPenSettings
applyFillSettings
mSurface.Line (p1.XLogical, p1.YLogical)-(p2.XLogical, p2.YLogical), , B

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "drawRectangle" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub drawText( _
                ByVal text As String, _
                ByVal location As Point)
Dim prevFont As StdFont
Dim textLines() As String
Dim i As Long
                
Dim failpoint As Long
On Error GoTo Err

setfont
    
mSurface.FontTransparent = mTextTransparent
mSurface.CurrentX = location.XLogical
mSurface.CurrentY = location.YLogical
mSurface.ForeColor = mTextColor
textLines = Split(text, vbCrLf)
For i = 0 To UBound(textLines)
    mSurface.Print textLines(i)
    mSurface.CurrentX = location.XLogical
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "drawText" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription



End Sub

Public Function newDimension(ByVal x As Double, _
                        ByVal y As Double) As Dimension
Dim failpoint As Long
On Error GoTo Err

Set newDimension = New Dimension
newDimension.Canvas = Me
newDimension.x = x
newDimension.y = y

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "newDimension" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function newPoint(ByVal x As Double, _
                        ByVal y As Double, _
                        Optional ByVal coordSystemX As CoordinateSystems = CoordsLogical, _
                        Optional ByVal coordSystemY As CoordinateSystems = CoordsLogical) As Point
Dim failpoint As Long
On Error GoTo Err

Set newPoint = New Point
newPoint.Canvas = Me
newPoint.x = x
newPoint.y = y
newPoint.CoordinateSystemX = coordSystemX
newPoint.CoordinateSystemY = coordSystemY

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "newPoint" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Sub paintBackground()
Dim failpoint As Long
On Error GoTo Err

'mSurface.Visible = True
If mGradientFillMode = GRADIENT_FILL_TRIANGLE Then
    GradientFill mSurface.hdc, _
                VarPtr(mVertices(0)), _
                UBound(mVertices) + 1, _
                VarPtr(mMeshTriangle(0)), _
                UBound(mMeshTriangle) + 1, _
                GRADIENT_FILL_TRIANGLE
Else
    GradientFill mSurface.hdc, _
                VarPtr(mVertices(0)), _
                UBound(mVertices) + 1, _
                VarPtr(mMeshRect(0)), _
                UBound(mMeshRect) + 1, _
                GRADIENT_FILL_RECT_H
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "paintBackground" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub setBrushAttributes( _
                ByVal Color As Long, _
                ByVal style As FillStyles)
Dim failpoint As Long
On Error GoTo Err

mBrushColor = Color
mBrushStyle = style

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "setBrushAttributes" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub setClippingRegion( _
                ByRef clipArea As TRectangle)
Dim res As Long
Dim failpoint As Long
On Error GoTo Err

mClipRegion = CreateRectRgn(Int(convertLogicalToPixelsX(clipArea.left - left)), _
                        Int(convertLogicalToPixelsY(top - clipArea.top)), _
                        Int(convertLogicalToPixelsX(clipArea.right - left) + 0.99999) + 1, _
                        Int(convertLogicalToPixelsY(top - clipArea.bottom) + 0.99999) + 1)
res = SelectClipRgn(mSurface.hdc, mClipRegion)

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "setClippingRegion" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub setPenAttributes( _
                ByVal Color As Long, _
                ByVal width As Long, _
                ByVal style As LineStyles, _
                ByVal mode As DrawModes)
Dim failpoint As Long
On Error GoTo Err

mPenColor = Color
mPenWidth = width
mPenStyle = style
mPenMode = mode

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "setPenAttributes" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub setTextAttributes( _
                ByVal Color As Long, _
                ByVal font As StdFont, _
                ByVal transparent As Boolean)
Dim failpoint As Long
On Error GoTo Err

mTextColor = Color
Set mTextFont = font
mTextTransparent = transparent

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "setTextAttributes" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Function TextHeight(ByVal value As String) As Single
Dim failpoint As Long
On Error GoTo Err

setfont
TextHeight = Abs(mSurface.TextHeight(value))

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "TextHeight" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function TextWidth(ByVal value As String) As Single
Dim failpoint As Long
On Error GoTo Err

setfont
TextWidth = mSurface.TextWidth(value)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "TextWidth" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ZOrder( _
                ByVal position As Integer)
mSurface.ZOrder position
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

Private Sub calculateGradientFillParameters()
Select Case UBound(mBackGradientFillColors)
Case 0
    ReDim mVertices(1) As TRIVERTEX
    ReDim mMeshRect(0) As GRADIENT_RECT
    setVertex 0, left, top, 0
    setVertex 1, right + 1, bottom, 0
    setMeshRect 0, 0, 1
    mGradientFillMode = GRADIENT_FILL_RECT_H
Case 1
    ReDim mVertices(1) As TRIVERTEX
    ReDim mMeshRect(0) As GRADIENT_RECT
    setVertex 0, left, top, 0
    setVertex 1, right + 1, bottom, 1
    setMeshRect 0, 0, 1
    mGradientFillMode = GRADIENT_FILL_RECT_H
Case 2
    ReDim mVertices(4) As TRIVERTEX
    ReDim mMeshTriangle(3) As GRADIENT_TRIANGLE
    setVertex 0, left, top, 0
    setVertex 1, right, top, 1
    setVertex 2, (left + right + 1) / 2, (top + bottom) / 2, 2
    setVertex 3, right + 1, bottom, 1
    setVertex 4, left, bottom, 0
    setMeshTriangle 0, 0, 1, 2
    setMeshTriangle 1, 1, 3, 2
    setMeshTriangle 2, 3, 4, 2
    setMeshTriangle 3, 4, 0, 2
    mGradientFillMode = GRADIENT_FILL_TRIANGLE
Case 3
    ReDim mVertices(3) As TRIVERTEX
    ReDim mMeshTriangle(1) As GRADIENT_TRIANGLE
    setVertex 0, left, top, 0
    setVertex 1, right + 1, top, 1
    setVertex 2, right + 1, bottom, 2
    setVertex 3, left, bottom, 3
    setMeshTriangle 0, 0, 1, 2
    setMeshTriangle 1, 2, 3, 0
    mGradientFillMode = GRADIENT_FILL_TRIANGLE
Case 4
    ReDim mVertices(4) As TRIVERTEX
    ReDim mMeshTriangle(3) As GRADIENT_TRIANGLE
    setVertex 0, left, top, 0
    setVertex 1, right + 1, top, 1
    setVertex 2, right + 1, bottom, 2
    setVertex 3, left, bottom, 3
    setVertex 4, (left + right + 1) / 2, (top + bottom) / 2, 4
    setMeshTriangle 0, 0, 1, 4
    setMeshTriangle 1, 1, 2, 4
    setMeshTriangle 2, 2, 3, 4
    setMeshTriangle 3, 3, 0, 4
    mGradientFillMode = GRADIENT_FILL_TRIANGLE
End Select
End Sub

Private Sub setfont()
If Not mTextFont Is Nothing Then
    With mSurface.font
        If mTextFont.size < 8 Then .size = mTextFont.size
       .name = mTextFont.name
       .Bold = mTextFont.Bold
       .Italic = mTextFont.Italic
       .size = mTextFont.size
       .Charset = mTextFont.Charset
       .Strikethrough = mTextFont.Strikethrough
       .Underline = mTextFont.Underline
       .Weight = mTextFont.Weight
    End With
End If
End Sub

Private Sub setMeshRect( _
                ByVal index As Long, _
                ByVal upperLeft As Long, _
                ByVal lowerRight As Long)
mMeshRect(index).lowerRight = lowerRight
mMeshRect(index).upperLeft = upperLeft
End Sub

Private Sub setMeshTriangle( _
                ByVal index As Long, _
                ByVal vertex1 As Long, _
                ByVal vertex2 As Long, _
                ByVal vertex3 As Long)
mMeshTriangle(index).vertex1 = vertex1
mMeshTriangle(index).vertex2 = vertex2
mMeshTriangle(index).vertex3 = vertex3
End Sub

Private Sub setScale()
If mSurface Is Nothing Then Exit Sub
mSurface.ScaleHeight = mBoundary.bottom - mBoundary.top
mSurface.ScaleWidth = mBoundary.right - mBoundary.left + 1
mSurface.ScaleTop = mBoundary.top
mSurface.ScaleLeft = mBoundary.left
mGaugeX = mSurface.ScaleWidth / mSurface.ScaleX(mSurface.ScaleWidth, ScaleModeConstants.vbUser, ScaleModeConstants.vbPixels)
'mGaugeX = (CDbl(mSurface.ScaleWidth) * CDbl(Screen.TwipsPerPixelX)) / CDbl(mSurface.width)
mGaugeY = -mSurface.ScaleHeight / mSurface.ScaleY(mSurface.ScaleHeight, ScaleModeConstants.vbUser, ScaleModeConstants.vbPixels)
'mGaugeY = -(CDbl(mSurface.ScaleHeight) * CDbl(Screen.TwipsPerPixelY)) / CDbl(mSurface.height)

calculateGradientFillParameters

End Sub

Private Sub setVertex( _
                ByRef index As Long, _
                ByVal x As Single, _
                ByVal y As Single, _
                ByVal colorIndex As Long)
mVertices(index).x = convertLogicalToPixelsX(x - left)
mVertices(index).y = convertLogicalToPixelsY(top - y)
mVertices(index).Red = mBackGradientFillColors(colorIndex) And &HFF&
mVertices(index).Green = (mBackGradientFillColors(colorIndex) And &HFF00&) / &H100&
mVertices(index).Blue = (mBackGradientFillColors(colorIndex) And &HFF0000) / &H10000
mVertices(index).Alpha = 0
End Sub

