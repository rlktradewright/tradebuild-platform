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
        X As Long
        Y As Long
End Type

Private Type rect
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Type TRIVERTEX
  X         As Long
  Y         As Long
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
                ByVal Hdc As Long, _
                lpRect As rect, _
                ByVal hBrush As Long) As Long

Private Declare Function GetBkColor Lib "gdi32" ( _
                ByVal Hdc As Long) As Long
                
Private Declare Function GetClipRgn Lib "gdi32" ( _
                ByVal Hdc As Long, _
                ByVal hRgn As Long) As Long
                
Private Declare Sub GradientFill Lib "msimg32" ( _
                ByVal Hdc As Long, _
                ByVal pVertices As Long, _
                ByVal numVertices As Long, _
                ByVal pMesh As Long, _
                ByVal numMesh As Long, _
                ByVal dwMode As Long)

Private Declare Function Polygon Lib "gdi32" ( _
                ByVal Hdc As Long, _
                lpPoint As POINTAPI, _
                ByVal nCount As Long) As Long

Private Declare Function SelectClipRgn Lib "gdi32" ( _
                ByVal Hdc As Long, _
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

' Text attributes
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

Friend Property Let BackColor(ByVal value As OLE_COLOR)
ReDim mBackGradientFillColors(0) As Long
mBackGradientFillColors(0) = value
calculateGradientFillParameters
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = mBackGradientFillColors(0)
End Property

Friend Property Let Bottom(ByVal value As Double)
mBoundary.Bottom = value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property

setScale
End Property

Public Property Get Bottom() As Double
Bottom = mBoundary.Bottom
End Property

Public Property Get Boundary() As TRectangle
Boundary = mBoundary
End Property

Public Property Get GaugeX() As Double
GaugeX = mGaugeX
End Property

Public Property Get GaugeY() As Double
GaugeY = mGaugeY
End Property

Public Property Get GradientFillColors() As Long()
GradientFillColors = mBackGradientFillColors
End Property

Friend Property Let GradientFillColors(ByRef value() As Long)
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
            ProjectName & "." & ModuleName & ":" & "GradientFillColors", _
            "At least 1 background gradient fill color must be supplied"
End If

If numColors > 5 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "GradientFillColors", _
            "At most 5 background gradient fill colors may be supplied"
End If

mBackGradientFillColors = value

calculateGradientFillParameters
End Property

Public Property Get Hdc() As Long
Hdc = mSurface.Hdc
End Property

Friend Property Let Height(ByVal value As Double)
mBoundary.Bottom = mBoundary.Top - value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
setScale
End Property

Public Property Get Height() As Double
If mBoundary.Top = PlusInfinityDouble Or mBoundary.Bottom = PlusInfinityDouble Then
    Height = PlusInfinityDouble
ElseIf mBoundary.Top = MinusInfinityDouble Or mBoundary.Bottom = MinusInfinityDouble Then
    Height = PlusInfinityDouble
Else
    Height = mBoundary.Top - mBoundary.Bottom
End If
End Property

Public Property Get HeightCm() As Double
HeightCm = mSurface.Height / TwipsPerCm
End Property

Friend Property Let Left(ByVal value As Double)
mBoundary.Left = value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
setScale
End Property

Public Property Get Left() As Double
Left = mBoundary.Left
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
' notifies that the physical Height of the Surface has changed.
If Not mBoundary.isValid Then Exit Sub
mSurface.ScaleTop = mBoundary.Top
setScale
End Sub

Friend Property Let Right(ByVal value As Double)
mBoundary.Right = value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
setScale
End Property

Public Property Get Right() As Double
Right = mBoundary.Right
End Property

Friend Property Let Surface(ByRef value As PictureBox)
rectInitialise mBoundary
Set mSurface = value
ReDim mBackGradientFillColors(0) As Long
mBackGradientFillColors(0) = vbWhite
Left = 0
Right = 1
Bottom = 0
Top = 1
End Property

Friend Property Let Top(ByVal value As Double)
mBoundary.Top = value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
setScale
End Property

Public Property Get Top() As Double
Top = mBoundary.Top
End Property

Public Property Get Width() As Double
If mBoundary.Right = PlusInfinityDouble Or mBoundary.Left = PlusInfinityDouble Then
    Width = PlusInfinitySingle
ElseIf mBoundary.Right = MinusInfinityDouble Or mBoundary.Left = MinusInfinityDouble Then
    Width = PlusInfinitySingle
Else
    Width = mBoundary.Right - mBoundary.Left
End If
End Property

Friend Property Let Width(ByVal value As Double)
mBoundary.Left = mBoundary.Right - value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
mSurface.ScaleLeft = mBoundary.Left
setScale
End Property

Public Property Get WidthCm() As Double
WidthCm = mSurface.Width / TwipsPerCm
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Clear()
Dim failpoint As Long
On Error GoTo Err

mSurface.Cls

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Clear" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub ClearClippingRegion()
Dim failpoint As Long
On Error GoTo Err

SelectClipRgn mSurface.Hdc, 0
DeleteObject mClipRegion

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ClearClippingRegion" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Function CompareTo( _
                ByVal GaugeX As Double, _
                ByVal GaugeY As Double, _
                ByRef rect As TRectangle) As Long
Dim failpoint As Long
On Error GoTo Err

If Not rect.isValid Or Not mBoundary.isValid Then
    CompareTo = CanvasComparisonCodes.CantCompare
    Exit Function
End If
If GaugeX <> mGaugeX Then
    CompareTo = CanvasComparisonCodes.GaugeChangedX
End If
If GaugeY <> mGaugeY Then
    CompareTo = CompareTo Or CanvasComparisonCodes.GaugeChangedY
End If
If rect.Left <> mBoundary.Left Or _
    rect.Right <> mBoundary.Right _
Then
    CompareTo = CompareTo Or CanvasComparisonCodes.BoundsChangedX
End If
If rect.Bottom <> mBoundary.Bottom Or _
    rect.Top <> mBoundary.Top _
Then
    CompareTo = CompareTo Or CanvasComparisonCodes.BoundsChangedY
End If

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "CompareTo" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription



End Function

Public Function ConvertCounterDistanceToDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertCounterDistanceToDistanceX = CDbl(mSurface.Width) / TwipsPerCm - value

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertCounterDistanceToDistanceX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertCounterDistanceToDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertCounterDistanceToDistanceY = CDbl(mSurface.Height) / TwipsPerCm - value    ' ((100 - 100 * value * TwipsPerCm / mSurface.Height) * mSurface.Height) * 100 / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertCounterDistanceToDistanceY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertCounterDistanceToLogicalX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertCounterDistanceToLogicalX = mBoundary.Right - ((value * TwipsPerCm) / CDbl(mSurface.Width) * (mBoundary.Right - mBoundary.Left))

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertCounterDistanceToLogicalX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertCounterDistanceToLogicalY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertCounterDistanceToLogicalY = mBoundary.Top - ((value * TwipsPerCm) / CDbl(mSurface.Height) * (mBoundary.Top - mBoundary.Bottom))

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertCounterDistanceToLogicalY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertCounterDistanceToRelativeX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertCounterDistanceToRelativeX = 100# - 100# * value * TwipsPerCm / CDbl(mSurface.Width)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertCounterDistanceToRelativeX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertCounterDistanceToRelativeY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertCounterDistanceToRelativeY = 100# - 100# * value * TwipsPerCm / CDbl(mSurface.Height)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertCounterDistanceToRelativeY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertDistanceToLogicalX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertDistanceToLogicalX = mBoundary.Left + (value * TwipsPerCm) / CDbl(mSurface.Width) * (mBoundary.Right - mBoundary.Left)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertDistanceToLogicalX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertDistanceToLogicalY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertDistanceToLogicalY = mBoundary.Bottom + (value * TwipsPerCm) / CDbl(mSurface.Height) * (mBoundary.Top - mBoundary.Bottom)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertDistanceToLogicalY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertDistanceToCounterDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertDistanceToCounterDistanceX = CDbl(mSurface.Width) / TwipsPerCm - value

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertDistanceToCounterDistanceX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertDistanceToCounterDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertDistanceToCounterDistanceY = CDbl(mSurface.Height) / TwipsPerCm - value

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertDistanceToCounterDistanceY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertDistanceToRelativeX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertDistanceToRelativeX = 100# * value * TwipsPerCm / CDbl(mSurface.Width)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertDistanceToRelativeX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertDistanceToRelativeY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertDistanceToRelativeY = 100# * value * TwipsPerCm / CDbl(mSurface.Height)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertDistanceToRelativeY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertLogicalToCounterDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertLogicalToCounterDistanceX = (mBoundary.Right - value) / (mBoundary.Right - mBoundary.Left) * mSurface.Width / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertLogicalToCounterDistanceX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertLogicalToCounterDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertLogicalToCounterDistanceY = (mBoundary.Top - value) / (mBoundary.Top - mBoundary.Bottom) * CDbl(mSurface.Height) / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertLogicalToCounterDistanceY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertLogicalToDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertLogicalToDistanceX = (value - mBoundary.Left) / (mBoundary.Right - mBoundary.Left) * CDbl(mSurface.Width) / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertLogicalToDistanceX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertLogicalToDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertLogicalToDistanceY = (value - mBoundary.Bottom) / (mBoundary.Top - mBoundary.Bottom) * CDbl(mSurface.Height) / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertLogicalToDistanceY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertLogicalToPixelsX(ByVal value As Double) As Long
Dim failpoint As Long
On Error GoTo Err

ConvertLogicalToPixelsX = mSurface.ScaleX(value, vbUser, vbPixels)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertLogicalToPixelsX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertLogicalToPixelsY(ByVal value As Double) As Long
Dim failpoint As Long
On Error GoTo Err

ConvertLogicalToPixelsY = -mSurface.ScaleY(value, vbUser, vbPixels)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertLogicalToPixelsY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertLogicalToRelativeX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertLogicalToRelativeX = (value - mBoundary.Left) / (mBoundary.Right - mBoundary.Left) * 100#

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertLogicalToRelativeX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertLogicalToRelativeY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertLogicalToRelativeY = (value - mBoundary.Bottom) / (mBoundary.Top - mBoundary.Bottom) * 100#

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertLogicalToRelativeY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertPixelsToLogicalX(ByVal value As Long) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertPixelsToLogicalX = CDbl(value) * mGaugeX

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertPixelsToLogicalX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertPixelsToLogicalY(ByVal value As Long) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertPixelsToLogicalY = CDbl(value) * mGaugeY

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertPixelsToLogicalY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertRelativeToLogicalX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertRelativeToLogicalX = mBoundary.Left + (value * (mBoundary.Right - mBoundary.Left) / 100#)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertRelativeToLogicalX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertRelativeToLogicalY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertRelativeToLogicalY = mBoundary.Bottom + (value * (mBoundary.Top - mBoundary.Bottom) / 100#)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertRelativeToLogicalY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertRelativeToCounterDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertRelativeToCounterDistanceX = ((100# - value) * CDbl(mSurface.Width)) / 100# / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertRelativeToCounterDistanceX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertRelativeToCounterDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertRelativeToCounterDistanceY = ((100# - value) * CDbl(mSurface.Height)) / 100# / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertRelativeToCounterDistanceY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertRelativeToDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertRelativeToDistanceX = (value * CDbl(mSurface.Width)) / 100# / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertRelativeToDistanceX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertRelativeToDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertRelativeToDistanceY = (value * CDbl(mSurface.Height)) * 100# / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ConvertRelativeToDistanceY" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Sub DrawCircle( _
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
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "DrawCircle" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub DrawLine( _
                ByVal startPoint As Point, _
                ByVal endPoint As Point)
Dim failpoint As Long
On Error GoTo Err

applyPenSettings
mSurface.Line (startPoint.XLogical, startPoint.YLogical)-(endPoint.XLogical, endPoint.YLogical)

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "DrawLine" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub DrawPoint( _
                ByVal p As Point)
Dim failpoint As Long
On Error GoTo Err

applyPenSettings
mSurface.PSet (p.XLogical, p.YLogical)

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "DrawPoint" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub DrawPolygon( _
                ByRef points() As Point)
Dim apiPoints() As POINTAPI
Dim failpoint As Long
On Error GoTo Err

ReDim apiPoints(UBound(points)) As POINTAPI
Dim i As Long

For i = 0 To UBound(points)
    apiPoints(i).X = ConvertLogicalToPixelsX(points(i).XLogical - mBoundary.Left)
    apiPoints(i).Y = ConvertLogicalToPixelsY(mBoundary.Top - points(i).YLogical)
Next

applyPenSettings
applyFillSettings
Polygon mSurface.Hdc, apiPoints(0), UBound(apiPoints) + 1

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "DrawPolygon" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub DrawRectangle( _
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
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "DrawRectangle" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub DrawText( _
                ByVal Text As String, _
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
textLines = Split(Text, vbCrLf)
For i = 0 To UBound(textLines)
    mSurface.Print textLines(i)
    mSurface.CurrentX = location.XLogical
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "DrawText" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription



End Sub

Public Function NewDimension(ByVal X As Double, _
                        ByVal Y As Double) As Dimension
Dim failpoint As Long
On Error GoTo Err

Set NewDimension = New Dimension
NewDimension.Canvas = Me
NewDimension.X = X
NewDimension.Y = Y

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "NewDimension" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function NewPoint(ByVal X As Double, _
                        ByVal Y As Double, _
                        Optional ByVal coordSystemX As CoordinateSystems = CoordsLogical, _
                        Optional ByVal coordSystemY As CoordinateSystems = CoordsLogical) As Point
Dim failpoint As Long
On Error GoTo Err

Set NewPoint = New Point
NewPoint.Canvas = Me
NewPoint.X = X
NewPoint.Y = Y
NewPoint.CoordinateSystemX = coordSystemX
NewPoint.CoordinateSystemY = coordSystemY

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "NewPoint" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Sub PaintBackground()
Dim failpoint As Long
On Error GoTo Err

'mSurface.Visible = True
If mGradientFillMode = GRADIENT_FILL_TRIANGLE Then
    GradientFill mSurface.Hdc, _
                VarPtr(mVertices(0)), _
                UBound(mVertices) + 1, _
                VarPtr(mMeshTriangle(0)), _
                UBound(mMeshTriangle) + 1, _
                GRADIENT_FILL_TRIANGLE
Else
    GradientFill mSurface.Hdc, _
                VarPtr(mVertices(0)), _
                UBound(mVertices) + 1, _
                VarPtr(mMeshRect(0)), _
                UBound(mMeshRect) + 1, _
                GRADIENT_FILL_RECT_H
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "PaintBackground" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub SetBrushAttributes( _
                ByVal Color As Long, _
                ByVal Style As FillStyles)
Dim failpoint As Long
On Error GoTo Err

mBrushColor = Color
mBrushStyle = Style

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetBrushAttributes" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub SetClippingRegion( _
                ByRef clipArea As TRectangle)
Dim res As Long
Dim failpoint As Long
On Error GoTo Err

mClipRegion = CreateRectRgn(Int(ConvertLogicalToPixelsX(clipArea.Left - Left)), _
                        Int(ConvertLogicalToPixelsY(Top - clipArea.Top)), _
                        Int(ConvertLogicalToPixelsX(clipArea.Right - Left) + 0.99999) + 1, _
                        Int(ConvertLogicalToPixelsY(Top - clipArea.Bottom) + 0.99999) + 1)
If mClipRegion = 0 Then
    Err.Raise ErrorCodes.ErrRuntimeException, _
            ProjectName & "." & ModuleName & ":" & "SetClippingRegion", _
            "Failed to create GDI region"
End If

res = SelectClipRgn(mSurface.Hdc, mClipRegion)

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetClippingRegion" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub SetPenAttributes( _
                ByVal Color As Long, _
                ByVal Width As Long, _
                ByVal Style As LineStyles, _
                ByVal mode As DrawModes)
Dim failpoint As Long
On Error GoTo Err

mPenColor = Color
mPenWidth = Width
mPenStyle = Style
mPenMode = mode

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetPenAttributes" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub SetTextAttributes( _
                ByVal Color As Long, _
                ByVal Font As StdFont, _
                ByVal transparent As Boolean)
Dim failpoint As Long
On Error GoTo Err

mTextColor = Color
Set mTextFont = Font
mTextTransparent = transparent

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetTextAttributes" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Function TextHeight(ByVal value As String) As Double
Dim failpoint As Long
On Error GoTo Err

setfont
TextHeight = Abs(mSurface.TextHeight(value))

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "TextHeight" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function TextWidth(ByVal value As String) As Double
Dim failpoint As Long
On Error GoTo Err

setfont
TextWidth = mSurface.TextWidth(value)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "TextWidth" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
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
    setVertex 0, Left, Top, 0
    setVertex 1, Right + 1, Bottom, 0
    setMeshRect 0, 0, 1
    mGradientFillMode = GRADIENT_FILL_RECT_H
Case 1
    ReDim mVertices(1) As TRIVERTEX
    ReDim mMeshRect(0) As GRADIENT_RECT
    setVertex 0, Left, Top, 0
    setVertex 1, Right + 1, Bottom, 1
    setMeshRect 0, 0, 1
    mGradientFillMode = GRADIENT_FILL_RECT_H
Case 2
    ReDim mVertices(4) As TRIVERTEX
    ReDim mMeshTriangle(3) As GRADIENT_TRIANGLE
    setVertex 0, Left, Top, 0
    setVertex 1, Right, Top, 1
    setVertex 2, (Left + Right + 1) / 2, (Top + Bottom) / 2, 2
    setVertex 3, Right + 1, Bottom, 1
    setVertex 4, Left, Bottom, 0
    setMeshTriangle 0, 0, 1, 2
    setMeshTriangle 1, 1, 3, 2
    setMeshTriangle 2, 3, 4, 2
    setMeshTriangle 3, 4, 0, 2
    mGradientFillMode = GRADIENT_FILL_TRIANGLE
Case 3
    ReDim mVertices(3) As TRIVERTEX
    ReDim mMeshTriangle(1) As GRADIENT_TRIANGLE
    setVertex 0, Left, Top, 0
    setVertex 1, Right + 1, Top, 1
    setVertex 2, Right + 1, Bottom, 2
    setVertex 3, Left, Bottom, 3
    setMeshTriangle 0, 0, 1, 2
    setMeshTriangle 1, 2, 3, 0
    mGradientFillMode = GRADIENT_FILL_TRIANGLE
Case 4
    ReDim mVertices(4) As TRIVERTEX
    ReDim mMeshTriangle(3) As GRADIENT_TRIANGLE
    setVertex 0, Left, Top, 0
    setVertex 1, Right + 1, Top, 1
    setVertex 2, Right + 1, Bottom, 2
    setVertex 3, Left, Bottom, 3
    setVertex 4, (Left + Right + 1) / 2, (Top + Bottom) / 2, 4
    setMeshTriangle 0, 0, 1, 4
    setMeshTriangle 1, 1, 2, 4
    setMeshTriangle 2, 2, 3, 4
    setMeshTriangle 3, 3, 0, 4
    mGradientFillMode = GRADIENT_FILL_TRIANGLE
End Select
End Sub

Public Sub checkValidDimensions()
Dim midPoint As Double
Dim newRange As Double

rectValidate mBoundary
If mBoundary.isValid Then
    Do While CSng(mBoundary.Top) - CSng(mBoundary.Bottom) = 0
        midPoint = mBoundary.Top - mBoundary.Bottom
        newRange = 10 * mBoundary.Top - 0.1 * mBoundary.Bottom
        mBoundary.Top = midPoint + newRange / 2
        mBoundary.Bottom = midPoint - newRange / 2
    Loop
    Do While CSng(mBoundary.Right) - CSng(mBoundary.Left) = 0
        midPoint = mBoundary.Right - mBoundary.Left
        newRange = 10 * mBoundary.Right - 0.1 * mBoundary.Left
        mBoundary.Top = midPoint + newRange / 2
        mBoundary.Bottom = midPoint - newRange / 2
    Loop
End If
End Sub

Private Sub setfont()
If Not mTextFont Is Nothing Then
    With mSurface.Font
        If mTextFont.size < 8 Then .size = mTextFont.size
       .Name = mTextFont.Name
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
mSurface.ScaleHeight = mBoundary.Bottom - mBoundary.Top
mSurface.ScaleWidth = mBoundary.Right - mBoundary.Left + 1
mSurface.ScaleTop = mBoundary.Top
mSurface.ScaleLeft = mBoundary.Left
mGaugeX = mSurface.ScaleWidth / mSurface.ScaleX(mSurface.ScaleWidth, ScaleModeConstants.vbUser, ScaleModeConstants.vbPixels)
'mGaugeX = (CDbl(mSurface.ScaleWidth) * CDbl(Screen.TwipsPerPixelX)) / CDbl(mSurface.Width)
mGaugeY = -mSurface.ScaleHeight / mSurface.ScaleY(mSurface.ScaleHeight, ScaleModeConstants.vbUser, ScaleModeConstants.vbPixels)
'mGaugeY = -(CDbl(mSurface.ScaleHeight) * CDbl(Screen.TwipsPerPixelY)) / CDbl(mSurface.Height)

calculateGradientFillParameters

End Sub

Private Sub setVertex( _
                ByRef index As Long, _
                ByVal X As Double, _
                ByVal Y As Double, _
                ByVal colorIndex As Long)
mVertices(index).X = ConvertLogicalToPixelsX(X - Left)
mVertices(index).Y = ConvertLogicalToPixelsY(Top - Y)
mVertices(index).Red = mBackGradientFillColors(colorIndex) And &HFF&
mVertices(index).Green = (mBackGradientFillColors(colorIndex) And &HFF00&) / &H100&
mVertices(index).Blue = (mBackGradientFillColors(colorIndex) And &HFF0000) / &H10000
mVertices(index).Alpha = 0
End Sub

