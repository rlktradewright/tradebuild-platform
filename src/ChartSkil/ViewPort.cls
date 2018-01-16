VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViewPort"
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


Private Const ModuleName                As String = "Viewport"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' External FUnction Declarations
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mCanvas As Canvas
Attribute mCanvas.VB_VarHelpID = -1
Private mBoundary As TRectangle
Private mGaugeX As Double
Private mGaugeY As Double

Private mRegionType As RegionTypes

Private mGradientFill As GradientFill

Private mTransform As XForm

Private mRegionRef As WeakReference

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mGradientFill = New GradientFill
End Sub

Private Sub Class_Terminate()
gLogger.Log pLogLevel:=LogLevelHighDetail, pProcName:="Class_Terminate", pModName:=ModuleName, pMsg:="Viewport terminated (" & gRegionTypeToString(mRegionType) & ")"
Debug.Print "Viewport terminated (" & gRegionTypeToString(mRegionType) & ")"
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get Bottom() As Double
Const ProcName As String = "Bottom"

On Error GoTo Err

Bottom = mBoundary.Bottom

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Boundary() As TRectangle
Boundary = mBoundary
End Property

Public Property Get Canvas() As Canvas
Set Canvas = mCanvas
End Property

Public Property Get GaugeX() As Double
GaugeX = mGaugeX
End Property

Public Property Get GaugeY() As Double
GaugeY = mGaugeY
End Property

Public Property Get GradientFillColors() As Long()
Const ProcName As String = "GradientFillColors"

On Error GoTo Err

GradientFillColors = mGradientFill.Colors

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let GradientFillColors(ByRef Value() As Long)
Dim ar() As Long
Const ProcName As String = "GradientFillColors"

On Error GoTo Err

ar = Value
mGradientFill.Colors = ar

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Height(ByVal Value As Double)
Const ProcName As String = "Height"

On Error GoTo Err

mBoundary.Bottom = mBoundary.Top - Value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
SetScale

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Height() As Double
Const ProcName As String = "Height"

On Error GoTo Err

If mBoundary.Top = PlusInfinityDouble Or mBoundary.Bottom = PlusInfinityDouble Then
    Height = PlusInfinityDouble
ElseIf mBoundary.Top = MinusInfinityDouble Or mBoundary.Bottom = MinusInfinityDouble Then
    Height = PlusInfinityDouble
Else
    Height = mBoundary.Top - mBoundary.Bottom
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HeightCm() As Double
Const ProcName As String = "HeightCm"

On Error GoTo Err

HeightCm = mCanvas.Height / TwipsPerCm

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Left(ByVal Value As Double)
Const ProcName As String = "Left"

On Error GoTo Err

mBoundary.Left = Value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
SetScale

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Left() As Double
Const ProcName As String = "Left"

On Error GoTo Err

Left = mBoundary.Left

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let MouseIcon(ByVal Value As IPictureDisp)
Const ProcName As String = "MouseIcon"

On Error GoTo Err

mCanvas.MouseIcon = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let MousePointer(ByVal Value As VBRUN.MousePointerConstants)
Const ProcName As String = "MousePointer"

On Error GoTo Err

mCanvas.MousePointer = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MousePointer() As VBRUN.MousePointerConstants
Const ProcName As String = "MousePointer"

On Error GoTo Err

MousePointer = mCanvas.MousePointer

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Region() As ChartRegion
Set Region = mRegionRef.Target
End Property

Friend Property Let RegionType(ByVal Value As RegionTypes)
mRegionType = Value
End Property

Friend Property Let Right(ByVal Value As Double)
Const ProcName As String = "Right"

On Error GoTo Err

mBoundary.Right = Value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
SetScale

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Right() As Double
Const ProcName As String = "Right"

On Error GoTo Err

Right = mBoundary.Right

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Top() As Double
Top = mBoundary.Top
End Property

Public Property Get Width() As Double
Const ProcName As String = "Width"

On Error GoTo Err

If mBoundary.Right = PlusInfinityDouble Or mBoundary.Left = PlusInfinityDouble Then
    Width = PlusInfinitySingle
ElseIf mBoundary.Right = MinusInfinityDouble Or mBoundary.Left = MinusInfinityDouble Then
    Width = PlusInfinitySingle
Else
    Width = mBoundary.Right - mBoundary.Left
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Width(ByVal Value As Double)
Const ProcName As String = "Width"

On Error GoTo Err

mBoundary.Left = mBoundary.Right - Value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
SetScale

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get WidthCm() As Double
Const ProcName As String = "WidthCm"

On Error GoTo Err

WidthCm = mCanvas.Width / TwipsPerCm

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Clear()
Const ProcName As String = "Clear"

On Error GoTo Err

mCanvas.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function CompareTo( _
                ByVal GaugeX As Double, _
                ByVal GaugeY As Double, _
                ByRef pRect As TRectangle) As ViewportComparisonCodes
Const ProcName As String = "CompareTo"

On Error GoTo Err

If Not pRect.isValid Or Not mBoundary.isValid Then
    CompareTo = ViewportComparisonCodes.CantCompare
    Exit Function
End If
If GaugeX <> mGaugeX Then
    CompareTo = ViewportComparisonCodes.GaugeChangedX
End If
If GaugeY <> mGaugeY Then
    CompareTo = CompareTo Or ViewportComparisonCodes.GaugeChangedY
End If
If pRect.Left <> mBoundary.Left Or _
    pRect.Right <> mBoundary.Right _
Then
    CompareTo = CompareTo Or ViewportComparisonCodes.BoundsChangedX
End If
If pRect.Bottom <> mBoundary.Bottom Or _
    pRect.Top <> mBoundary.Top _
Then
    CompareTo = CompareTo Or ViewportComparisonCodes.BoundsChangedY
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertCoordsLogicalToPixelsX(ByVal Value As Double) As Long
Const ProcName As String = "ConvertCoordsLogicalToPixelsX"

On Error GoTo Err

ConvertCoordsLogicalToPixelsX = ConvertLogicalToPixelsX(Value - mBoundary.Left)
ConvertCoordsLogicalToPixelsX = ConvertCoordsLogicalToPixelsX - mTransform.eDx
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertCoordsLogicalToPixelsY(ByVal Value As Double) As Long
Const ProcName As String = "ConvertCoordsLogicalToPixelsY"

On Error GoTo Err

ConvertCoordsLogicalToPixelsY = ConvertLogicalToPixelsY(mBoundary.Top - Value)
ConvertCoordsLogicalToPixelsY = ConvertCoordsLogicalToPixelsY - mTransform.eDy

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertCounterDistanceToDistanceX(ByVal Value As Double) As Double
Const ProcName As String = "ConvertCounterDistanceToDistanceX"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertCounterDistanceToDistanceX = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertCounterDistanceToDistanceX = MinusInfinityDouble
Else
    ConvertCounterDistanceToDistanceX = CDbl(mCanvas.Width) / TwipsPerCm - Value
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertCounterDistanceToDistanceY(ByVal Value As Double) As Double
Const ProcName As String = "ConvertCounterDistanceToDistanceY"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertCounterDistanceToDistanceY = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertCounterDistanceToDistanceY = MinusInfinityDouble
Else
    ConvertCounterDistanceToDistanceY = CDbl(mCanvas.Height) / TwipsPerCm - Value
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertCounterDistanceToLogicalX(ByVal Value As Double) As Double
Const ProcName As String = "ConvertCounterDistanceToLogicalX"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertCounterDistanceToLogicalX = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertCounterDistanceToLogicalX = MinusInfinityDouble
Else
    ConvertCounterDistanceToLogicalX = (mBoundary.Right - mBoundary.Left) - ((Value * TwipsPerCm) / CDbl(mCanvas.Width) * (mBoundary.Right - mBoundary.Left))
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertCounterDistanceToLogicalY(ByVal Value As Double) As Double
Const ProcName As String = "ConvertCounterDistanceToLogicalY"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertCounterDistanceToLogicalY = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertCounterDistanceToLogicalY = MinusInfinityDouble
Else
    ConvertCounterDistanceToLogicalY = (mBoundary.Top - mBoundary.Bottom) - ((Value * TwipsPerCm) / CDbl(mCanvas.Height) * (mBoundary.Top - mBoundary.Bottom))
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertCounterDistanceToRelativeX(ByVal Value As Double) As Double
Const ProcName As String = "ConvertCounterDistanceToRelativeX"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertCounterDistanceToRelativeX = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertCounterDistanceToRelativeX = MinusInfinityDouble
Else
    ConvertCounterDistanceToRelativeX = 100# - 100# * Value * TwipsPerCm / CDbl(mCanvas.Width)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertCounterDistanceToRelativeY(ByVal Value As Double) As Double
Const ProcName As String = "ConvertCounterDistanceToRelativeY"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertCounterDistanceToRelativeY = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertCounterDistanceToRelativeY = MinusInfinityDouble
Else
    ConvertCounterDistanceToRelativeY = 100# - 100# * Value * TwipsPerCm / CDbl(mCanvas.Height)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertDistanceToCounterDistanceX(ByVal Value As Double) As Double
Const ProcName As String = "ConvertDistanceToCounterDistanceX"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertDistanceToCounterDistanceX = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertDistanceToCounterDistanceX = MinusInfinityDouble
Else
    ConvertDistanceToCounterDistanceX = CDbl(mCanvas.Width) / TwipsPerCm - Value
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertDistanceToCounterDistanceY(ByVal Value As Double) As Double
Const ProcName As String = "ConvertDistanceToCounterDistanceY"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertDistanceToCounterDistanceY = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertDistanceToCounterDistanceY = MinusInfinityDouble
Else
    ConvertDistanceToCounterDistanceY = CDbl(mCanvas.Height) / TwipsPerCm - Value
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertDistanceToLogicalX(ByVal Value As Double) As Double
Const ProcName As String = "ConvertDistanceToLogicalX"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertDistanceToLogicalX = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertDistanceToLogicalX = MinusInfinityDouble
Else
    ConvertDistanceToLogicalX = (Value * TwipsPerCm) / CDbl(mCanvas.Width) * (mBoundary.Right - mBoundary.Left)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertDistanceToLogicalY(ByVal Value As Double) As Double
Const ProcName As String = "ConvertDistanceToLogicalY"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertDistanceToLogicalY = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertDistanceToLogicalY = MinusInfinityDouble
Else
    ConvertDistanceToLogicalY = (Value * TwipsPerCm) / CDbl(mCanvas.Height) * (mBoundary.Top - mBoundary.Bottom)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertDistanceToPixelsX(ByVal Value As Double) As Long
Const ProcName As String = "ConvertDistanceToPixelsX"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertDistanceToPixelsX = PlusInfinityLong
ElseIf Value = MinusInfinityDouble Then
    ConvertDistanceToPixelsX = MinusInfinityLong
Else
    ConvertDistanceToPixelsX = Round((Value * TwipsPerCm) / Screen.TwipsPerPixelX, 0)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertDistanceToPixelsY(ByVal Value As Double) As Long
Const ProcName As String = "ConvertDistanceToPixelsY"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertDistanceToPixelsY = PlusInfinityLong
ElseIf Value = MinusInfinityDouble Then
    ConvertDistanceToPixelsY = MinusInfinityLong
Else
    ConvertDistanceToPixelsY = Round((Value * TwipsPerCm) / Screen.TwipsPerPixelY, 0)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertDistanceToRelativeX(ByVal Value As Double) As Double
Const ProcName As String = "ConvertDistanceToRelativeX"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertDistanceToRelativeX = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertDistanceToRelativeX = MinusInfinityDouble
Else
    ConvertDistanceToRelativeX = 100# * Value * TwipsPerCm / CDbl(mCanvas.Width)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertDistanceToRelativeY(ByVal Value As Double) As Double
Const ProcName As String = "ConvertDistanceToRelativeY"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertDistanceToRelativeY = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertDistanceToRelativeY = MinusInfinityDouble
Else
    ConvertDistanceToRelativeY = 100# * Value * TwipsPerCm / CDbl(mCanvas.Height)
End If
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertLogicalToCounterDistanceX(ByVal Value As Double) As Double
Const ProcName As String = "ConvertLogicalToCounterDistanceX"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertLogicalToCounterDistanceX = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertLogicalToCounterDistanceX = MinusInfinityDouble
Else
    ConvertLogicalToCounterDistanceX = (mBoundary.Right - Value) / (mBoundary.Right - mBoundary.Left) * mCanvas.Width / TwipsPerCm
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertLogicalToCounterDistanceY(ByVal Value As Double) As Double
Const ProcName As String = "ConvertLogicalToCounterDistanceY"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertLogicalToCounterDistanceY = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertLogicalToCounterDistanceY = MinusInfinityDouble
Else
    ConvertLogicalToCounterDistanceY = (mBoundary.Top - Value) / (mBoundary.Top - mBoundary.Bottom) * CDbl(mCanvas.Height) / TwipsPerCm
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertLogicalToDistanceX(ByVal Value As Double) As Double
Const ProcName As String = "ConvertLogicalToDistanceX"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertLogicalToDistanceX = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertLogicalToDistanceX = MinusInfinityDouble
Else
    ConvertLogicalToDistanceX = (Value - mBoundary.Left) / (mBoundary.Right - mBoundary.Left) * CDbl(mCanvas.Width) / TwipsPerCm
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertLogicalToDistanceY(ByVal Value As Double) As Double
Const ProcName As String = "ConvertLogicalToDistanceY"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertLogicalToDistanceY = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertLogicalToDistanceY = MinusInfinityDouble
Else
    ConvertLogicalToDistanceY = (Value - mBoundary.Bottom) / (mBoundary.Top - mBoundary.Bottom) * CDbl(mCanvas.Height) / TwipsPerCm
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertLogicalToPixelsX(ByVal Value As Double) As Long
Const ProcName As String = "ConvertLogicalToPixelsX"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertLogicalToPixelsX = PlusInfinityLong
ElseIf Value = MinusInfinityDouble Then
    ConvertLogicalToPixelsX = MinusInfinityLong
Else
    ConvertLogicalToPixelsX = mCanvas.Surface.ScaleX(Value, vbUser, vbPixels)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
    End Function

Public Function ConvertLogicalToPixelsY(ByVal Value As Double) As Long
Const ProcName As String = "ConvertLogicalToPixelsY"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertLogicalToPixelsY = PlusInfinityLong
ElseIf Value = MinusInfinityDouble Then
    ConvertLogicalToPixelsY = MinusInfinityLong
Else
    ConvertLogicalToPixelsY = -Canvas.Surface.ScaleY(Value, vbUser, vbPixels)
End If
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertLogicalToRelativeX(ByVal Value As Double) As Double
Const ProcName As String = "ConvertLogicalToRelativeX"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertLogicalToRelativeX = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertLogicalToRelativeX = MinusInfinityDouble
Else
    ConvertLogicalToRelativeX = (Value - mBoundary.Left) / (mBoundary.Right - mBoundary.Left) * 100#
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertLogicalToRelativeY(ByVal Value As Double) As Double
Const ProcName As String = "ConvertLogicalToRelativeY"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertLogicalToRelativeY = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertLogicalToRelativeY = MinusInfinityDouble
Else
    ConvertLogicalToRelativeY = (Value - mBoundary.Bottom) / (mBoundary.Top - mBoundary.Bottom) * 100#
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertPixelsToDistanceX(ByVal Value As Long) As Double
Const ProcName As String = "ConvertPixelsToDistanceX"

On Error GoTo Err

If Value = PlusInfinityLong Then
    ConvertPixelsToDistanceX = PlusInfinityDouble
ElseIf Value = MinusInfinityLong Then
    ConvertPixelsToDistanceX = MinusInfinityDouble
Else
    ConvertPixelsToDistanceX = Value * Screen.TwipsPerPixelX / TwipsPerCm
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertPixelsToDistanceY(ByVal Value As Long) As Double
Const ProcName As String = "ConvertPixelsToDistanceY"

On Error GoTo Err

If Value = PlusInfinityLong Then
    ConvertPixelsToDistanceY = PlusInfinityDouble
ElseIf Value = MinusInfinityLong Then
    ConvertPixelsToDistanceY = MinusInfinityDouble
Else
    ConvertPixelsToDistanceY = Value * Screen.TwipsPerPixelY / TwipsPerCm
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertPixelsToLogicalX(ByVal Value As Long) As Double
Const ProcName As String = "ConvertPixelsToLogicalX"

On Error GoTo Err

If Value = PlusInfinityLong Then
    ConvertPixelsToLogicalX = PlusInfinityDouble
ElseIf Value = MinusInfinityLong Then
    ConvertPixelsToLogicalX = MinusInfinityDouble
Else
    ConvertPixelsToLogicalX = Canvas.Surface.ScaleX(Value, vbPixels, vbUser)
    'ConvertPixelsToLogicalX = CDbl(value) * mGaugeX
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertPixelsToLogicalY(ByVal Value As Long) As Double
Const ProcName As String = "ConvertPixelsToLogicalY"

On Error GoTo Err

If Value = PlusInfinityLong Then
    ConvertPixelsToLogicalY = PlusInfinityDouble
ElseIf Value = MinusInfinityLong Then
    ConvertPixelsToLogicalY = MinusInfinityDouble
Else
    ConvertPixelsToLogicalY = -Canvas.Surface.ScaleY(Value, vbPixels, vbUser)
    'ConvertPixelsToLogicalY = CDbl(value) * mGaugeY
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertPointsToWin32Points( _
                ByRef points() As Point) As GDI_POINT()
Dim GDI_POINTs() As GDI_POINT
Const ProcName As String = "ConvertPointsToWin32Points"

On Error GoTo Err

ReDim GDI_POINTs(UBound(points)) As GDI_POINT
Dim i As Long

For i = 0 To UBound(points)
    GDI_POINTs(i).X = ConvertCoordsLogicalToPixelsX(points(i).XLogical(Me.Region))
    GDI_POINTs(i).Y = ConvertCoordsLogicalToPixelsY(points(i).yLogical(Me.Region))
Next

ConvertPointsToWin32Points = GDI_POINTs

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertRectToTRectangle( _
                ByRef pRect As GDI_RECT) As TRectangle
Dim lTRect As TRectangle
Const ProcName As String = "ConvertRectToTRectangle"

On Error GoTo Err

If pRect.Left = MinusInfinityLong Then
    lTRect.Left = MinusInfinityDouble
Else
    lTRect.Left = Left + ConvertPixelsToLogicalX(pRect.Left)
End If

If pRect.Top = MinusInfinityLong Then
    lTRect.Top = PlusInfinityDouble
Else
    lTRect.Top = Top - ConvertPixelsToLogicalY(pRect.Top)
End If

If pRect.Right = PlusInfinityLong Then
    lTRect.Right = PlusInfinityDouble
Else
    'lTRect.Right = Left + 0.999999999999999 * ConvertPixelsToLogicalX(pRect.Right + 1)
    lTRect.Right = Left + ConvertPixelsToLogicalX(pRect.Right)
End If

If pRect.Bottom = PlusInfinityLong Then
    lTRect.Bottom = MinusInfinityDouble
Else
    'lTRect.Bottom = Top - 0.999999999999999 * ConvertPixelsToLogicalY(pRect.Bottom + 1)
    lTRect.Bottom = Top - ConvertPixelsToLogicalY(pRect.Bottom)
End If

RectValidate lTRect
ConvertRectToTRectangle = lTRect

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertRelativeToCounterDistanceX(ByVal Value As Double) As Double
Const ProcName As String = "ConvertRelativeToCounterDistanceX"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertRelativeToCounterDistanceX = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertRelativeToCounterDistanceX = MinusInfinityDouble
Else
    ConvertRelativeToCounterDistanceX = ((100# - Value) * CDbl(mCanvas.Width)) / 100# / TwipsPerCm
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertRelativeToCounterDistanceY(ByVal Value As Double) As Double
Const ProcName As String = "ConvertRelativeToCounterDistanceY"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertRelativeToCounterDistanceY = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertRelativeToCounterDistanceY = MinusInfinityDouble
Else
    ConvertRelativeToCounterDistanceY = ((100# - Value) * CDbl(mCanvas.Height)) / 100# / TwipsPerCm
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertRelativeToDistanceX(ByVal Value As Double) As Double
Const ProcName As String = "ConvertRelativeToDistanceX"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertRelativeToDistanceX = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertRelativeToDistanceX = MinusInfinityDouble
Else
    ConvertRelativeToDistanceX = (Value * CDbl(mCanvas.Width)) / 100# / TwipsPerCm
End If
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertRelativeToDistanceY(ByVal Value As Double) As Double
Const ProcName As String = "ConvertRelativeToDistanceY"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertRelativeToDistanceY = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertRelativeToDistanceY = MinusInfinityDouble
Else
    ConvertRelativeToDistanceY = (Value * CDbl(mCanvas.Height)) * 100# / TwipsPerCm
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertRelativeToLogicalX(ByVal Value As Double) As Double
Const ProcName As String = "ConvertRelativeToLogicalX"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertRelativeToLogicalX = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertRelativeToLogicalX = MinusInfinityDouble
Else
    'ConvertRelativeToLogicalX = mBoundary.Left + (value * (mBoundary.Right - mBoundary.Left) / 100#)
    ConvertRelativeToLogicalX = (Value * (mBoundary.Right - mBoundary.Left) / 100#)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertRelativeToLogicalY(ByVal Value As Double) As Double
Const ProcName As String = "ConvertRelativeToLogicalY"

On Error GoTo Err

If Value = PlusInfinityDouble Then
    ConvertRelativeToLogicalY = PlusInfinityDouble
ElseIf Value = MinusInfinityDouble Then
    ConvertRelativeToLogicalY = MinusInfinityDouble
Else
    'ConvertRelativeToLogicalY = mBoundary.Bottom + (value * (mBoundary.Top - mBoundary.Bottom) / 100#)
    ConvertRelativeToLogicalY = (Value * (mBoundary.Top - mBoundary.Bottom) / 100#)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ConvertTRectangleToRect( _
                ByRef pTRect As TRectangle) As GDI_RECT
Dim lRect As GDI_RECT

Const ProcName As String = "ConvertTRectangleToRect"

On Error GoTo Err

If pTRect.Left = MinusInfinityDouble Then
    lRect.Left = MinusInfinityLong
Else
    lRect.Left = ConvertCoordsLogicalToPixelsX(pTRect.Left)
End If

If pTRect.Top = PlusInfinityDouble Then
    lRect.Top = MinusInfinityLong
Else
    lRect.Top = ConvertCoordsLogicalToPixelsY(pTRect.Top)
End If

If pTRect.Right = PlusInfinityDouble Then
    lRect.Right = PlusInfinityLong
Else
    lRect.Right = ConvertCoordsLogicalToPixelsX(pTRect.Right)
End If

If pTRect.Bottom = MinusInfinityDouble Then
    lRect.Bottom = PlusInfinityLong
Else
    lRect.Bottom = ConvertCoordsLogicalToPixelsY(pTRect.Bottom)
End If

ConvertTRectangleToRect = lRect

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Finish()
Const ProcName As String = "Finish"

On Error GoTo Err

gLogger.Log "Viewport finished", ProcName, ModuleName, LogLevelHighDetail
mCanvas.Finish
Set mCanvas = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pCanvas As Canvas, _
                ByVal pRegion As ChartRegion, _
                ByVal pRegionType As RegionTypes)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mCanvas = pCanvas
mCanvas.ViewPort = Me

Set mRegionRef = CreateWeakReference(pRegion)
mRegionType = pRegionType

setGradientFillArea

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyResizedY()
' notifies that the physical Height of the Surface has changed.
Const ProcName As String = "NotifyResizedY"

On Error GoTo Err

If Not mBoundary.isValid Then Exit Sub
SetScale

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PaintBackground()
Const ProcName As String = "PaintBackground"

On Error GoTo Err

mGradientFill.Fill mCanvas.hDC

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Reset()
ModifyWorldTransform mCanvas.hDC, mTransform, MWT_IDENTITY
mTransform.eM11 = 1
mTransform.eM12 = 0
mTransform.eM21 = 0
mTransform.eM22 = 1
mTransform.eDx = 0
mTransform.eDy = 0
End Sub

Public Sub Rotate(ByVal angleRadians As Double)
Dim pt As TPoint
RotateAboutPoint angleRadians, pt
End Sub

Public Sub RotateAboutPoint(ByVal angleRadians As Double, ByRef pPoint As TPoint)
mTransform.eM11 = Cos(-angleRadians)
mTransform.eM12 = Sin(-angleRadians)
mTransform.eM21 = -mTransform.eM12
mTransform.eM22 = mTransform.eM11
mTransform.eDx = ConvertCoordsLogicalToPixelsX(pPoint.X)
mTransform.eDy = ConvertCoordsLogicalToPixelsY(pPoint.Y)
SetGraphicsMode mCanvas.hDC, GM_ADVANCED
SetWorldTransform mCanvas.hDC, mTransform
End Sub

Friend Sub SetVerticalBounds( _
                ByVal pBottom As Double, _
                ByVal pTop As Double)
Const ProcName As String = "SetVerticalBounds"
On Error GoTo Err

Assert pTop <> 65535, "pTop is 65535"
mBoundary.Bottom = pBottom
mBoundary.Top = pTop

checkValidDimensions
If Not mBoundary.isValid Then Exit Sub
SetScale

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ZOrder( _
                ByVal Position As Integer)
Const ProcName As String = "ZOrder"

On Error GoTo Err

mCanvas.ZOrder Position

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub UndrawArea( _
                ByRef undrawnArea As TRectangle)
Const ProcName As String = "UndrawArea"

On Error GoTo Err

Canvas.SetClippingRegion ConvertTRectangleToRect(undrawnArea)
PaintBackground
Canvas.ClearClippingRegion

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'================================================================================
' Helper Functions
'================================================================================

Private Sub checkValidDimensions()
Const ProcName As String = "checkValidDimensions"
On Error GoTo Err

RectValidate mBoundary
If mBoundary.isValid Then
    Dim midPoint As Double
    Dim range As Double
    Do While CSng(mBoundary.Top) - CSng(mBoundary.Bottom) = 0
        range = mBoundary.Top - mBoundary.Bottom
        midPoint = mBoundary.Bottom + range / 2
        mBoundary.Top = midPoint + 5 * range
        mBoundary.Bottom = midPoint - 5 * range
    Loop
    Do While CSng(mBoundary.Right) - CSng(mBoundary.Left) = 0
        range = mBoundary.Right - mBoundary.Left
        midPoint = mBoundary.Left + range / 2
        mBoundary.Left = midPoint + 5 * range
        mBoundary.Right = midPoint - 5 * range
    Loop
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setGradientFillArea()
Dim lRect As GDI_RECT
Const ProcName As String = "setGradientFillArea"

On Error GoTo Err

lRect.Right = mCanvas.Width / Screen.TwipsPerPixelX
lRect.Bottom = mCanvas.Height / Screen.TwipsPerPixelY
mGradientFill.Area = lRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub SetScale()
Const ProcName As String = "SetScale"

On Error GoTo Err

If Not mCanvas Is Nothing Then mCanvas.SetScale mBoundary

mGaugeX = ((mBoundary.Right - mBoundary.Left) * CDbl(Screen.TwipsPerPixelX)) / CDbl(mCanvas.Width)
mGaugeY = -((mBoundary.Bottom - mBoundary.Top) * CDbl(Screen.TwipsPerPixelY)) / CDbl(mCanvas.Height)

setGradientFillArea

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



