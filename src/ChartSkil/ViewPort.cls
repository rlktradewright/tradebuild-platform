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

Event BoundaryChanged()

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

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mGradientFill = New GradientFill
End Sub

Private Sub Class_Terminate()
gLogger.Log LogLevelHighDetail, "Viewport terminated (" & gRegionTypeToString(mRegionType) & ")"
Debug.Print "Viewport terminated (" & gRegionTypeToString(mRegionType) & ")"
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' mCanvas Event Handlers
'================================================================================

Private Sub mCanvas_SizeChanged()
setGradientFillArea
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let BackColor(ByVal value As Long)
mGradientFill.Color = value
End Property

Public Property Get BackColor() As Long
BackColor = mGradientFill.Color
End Property

Friend Property Let Bottom(ByVal value As Double)
mBoundary.Bottom = value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property

SetScale
RaiseEvent BoundaryChanged
End Property

Public Property Get Bottom() As Double
Bottom = mBoundary.Bottom
End Property

Public Property Get Boundary() As TRectangle
Boundary = mBoundary
End Property

Friend Property Let Canvas(ByRef value As Canvas)
Set mCanvas = value
setGradientFillArea
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
GradientFillColors = mGradientFill.Colors
End Property

Friend Property Let GradientFillColors(ByRef value() As Long)
Dim ar() As Long
ar = value
mGradientFill.Colors = ar
End Property

Friend Property Let Height(ByVal value As Double)
mBoundary.Bottom = mBoundary.Top - value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
SetScale
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
HeightCm = mCanvas.Height / TwipsPerCm
End Property

Friend Property Let Left(ByVal value As Double)
mBoundary.Left = value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
SetScale
RaiseEvent BoundaryChanged
End Property

Public Property Get Left() As Double
Left = mBoundary.Left
End Property

Friend Property Let MouseIcon(ByVal value As IPictureDisp)
mCanvas.MouseIcon = value
End Property

Friend Property Let MousePointer(ByVal value As VBRUN.MousePointerConstants)
mCanvas.MousePointer = value
End Property

Public Property Get MousePointer() As VBRUN.MousePointerConstants
MousePointer = mCanvas.MousePointer
End Property

Friend Property Let RegionType(ByVal value As RegionTypes)
mRegionType = value
End Property

Friend Sub ResizedY()
' notifies that the physical Height of the Surface has changed.
If Not mBoundary.isValid Then Exit Sub
SetScale
End Sub

Friend Property Let Right(ByVal value As Double)
mBoundary.Right = value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
SetScale
RaiseEvent BoundaryChanged
End Property

Public Property Get Right() As Double
Right = mBoundary.Right
End Property

Friend Property Let Top(ByVal value As Double)
mBoundary.Top = value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
SetScale
RaiseEvent BoundaryChanged
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
SetScale
End Property

Public Property Get WidthCm() As Double
WidthCm = mCanvas.Width / TwipsPerCm
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Clear()
Dim failpoint As Long
On Error GoTo Err

mCanvas.Clear

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Clear" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Friend Function CompareTo( _
                ByVal GaugeX As Double, _
                ByVal GaugeY As Double, _
                ByRef pRect As TRectangle) As Long
Dim failpoint As Long
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
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "CompareTo" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Public Function ConvertCounterDistanceToDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertCounterDistanceToDistanceX = CDbl(mCanvas.Width) / TwipsPerCm - value

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertCounterDistanceToDistanceX" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertCounterDistanceToDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertCounterDistanceToDistanceY = CDbl(mCanvas.Height) / TwipsPerCm - value    ' ((100 - 100 * value * TwipsPerCm / mCanvas.Height) * mCanvas.Height) * 100 / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertCounterDistanceToDistanceY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertCounterDistanceToLogicalX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertCounterDistanceToLogicalX = mBoundary.Right - ((value * TwipsPerCm) / CDbl(mCanvas.Width) * (mBoundary.Right - mBoundary.Left))

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertCounterDistanceToLogicalX" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertCounterDistanceToLogicalY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertCounterDistanceToLogicalY = mBoundary.Top - ((value * TwipsPerCm) / CDbl(mCanvas.Height) * (mBoundary.Top - mBoundary.Bottom))

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertCounterDistanceToLogicalY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertCounterDistanceToRelativeX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertCounterDistanceToRelativeX = 100# - 100# * value * TwipsPerCm / CDbl(mCanvas.Width)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertCounterDistanceToRelativeX" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertCounterDistanceToRelativeY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertCounterDistanceToRelativeY = 100# - 100# * value * TwipsPerCm / CDbl(mCanvas.Height)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertCounterDistanceToRelativeY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertDistanceToCounterDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertDistanceToCounterDistanceX = CDbl(mCanvas.Width) / TwipsPerCm - value

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertDistanceToCounterDistanceX" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertDistanceToCounterDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertDistanceToCounterDistanceY = CDbl(mCanvas.Height) / TwipsPerCm - value

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertDistanceToCounterDistanceY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertDistanceToLogicalX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertDistanceToLogicalX = mBoundary.Left + (value * TwipsPerCm) / CDbl(mCanvas.Width) * (mBoundary.Right - mBoundary.Left)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertDistanceToLogicalX" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertDistanceToLogicalY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertDistanceToLogicalY = mBoundary.Bottom + (value * TwipsPerCm) / CDbl(mCanvas.Height) * (mBoundary.Top - mBoundary.Bottom)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertDistanceToLogicalY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertDistanceToPixelsX(ByVal value As Double) As Long
Dim failpoint As Long
On Error GoTo Err

ConvertDistanceToPixelsX = Round((value * TwipsPerCm) / Screen.TwipsPerPixelX, 0)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertDistanceToPixelsX" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertDistanceToPixelsY(ByVal value As Double) As Long
Dim failpoint As Long
On Error GoTo Err

ConvertDistanceToPixelsY = Round((value * TwipsPerCm) / Screen.TwipsPerPixelY, 0)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertDistanceToPixelsY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertDistanceToRelativeX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertDistanceToRelativeX = 100# * value * TwipsPerCm / CDbl(mCanvas.Width)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertDistanceToRelativeX" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertDistanceToRelativeY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertDistanceToRelativeY = 100# * value * TwipsPerCm / CDbl(mCanvas.Height)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertDistanceToRelativeY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertLogicalToCounterDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertLogicalToCounterDistanceX = (mBoundary.Right - value) / (mBoundary.Right - mBoundary.Left) * mCanvas.Width / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertLogicalToCounterDistanceX" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertLogicalToCounterDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertLogicalToCounterDistanceY = (mBoundary.Top - value) / (mBoundary.Top - mBoundary.Bottom) * CDbl(mCanvas.Height) / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertLogicalToCounterDistanceY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertLogicalToDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertLogicalToDistanceX = (value - mBoundary.Left) / (mBoundary.Right - mBoundary.Left) * CDbl(mCanvas.Width) / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertLogicalToDistanceX" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertLogicalToDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertLogicalToDistanceY = (value - mBoundary.Bottom) / (mBoundary.Top - mBoundary.Bottom) * CDbl(mCanvas.Height) / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertLogicalToDistanceY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertLogicalToPixelsX(ByVal value As Double) As Long
Dim failpoint As Long
On Error GoTo Err

'ConvertLogicalToPixelsX = Round(mCanvas.Surface.ScaleX(value, vbUser, vbPixels), 0)
ConvertLogicalToPixelsX = Int(value / mGaugeX)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertLogicalToPixelsX" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertLogicalToPixelsY(ByVal value As Double) As Long
Dim failpoint As Long
On Error GoTo Err

'ConvertLogicalToPixelsY = Round(-Canvas.Surface.ScaleY(value, vbUser, vbPixels), 0)
ConvertLogicalToPixelsY = Int(value / mGaugeY)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertLogicalToPixelsY" & "." & failpoint
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
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertLogicalToRelativeX" & "." & failpoint
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
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertLogicalToRelativeY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertPixelsDistanceX(ByVal value As Long) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertPixelsDistanceX = value * Screen.TwipsPerPixelX / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertPixelsDistanceX" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertPixelsDistanceY(ByVal value As Long) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertPixelsDistanceY = value * Screen.TwipsPerPixelY / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertPixelsDistanceY" & "." & failpoint
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
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertPixelsToLogicalX" & "." & failpoint
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
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertPixelsToLogicalY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertPointsToWin32Points( _
                ByRef points() As Point) As W32Point()
Dim w32Points() As W32Point
ReDim w32Points(UBound(points)) As W32Point
Dim i As Long

For i = 0 To UBound(points)
    w32Points(i).X = ConvertLogicalToPixelsX(points(i).XLogical - mBoundary.Left)
    w32Points(i).Y = ConvertLogicalToPixelsY(mBoundary.Top - points(i).YLogical)
Next

ConvertPointsToWin32Points = w32Points
End Function

Public Function ConvertRectToTRectangle( _
                ByRef pRect As RECT) As TRectangle
Dim lTRect As TRectangle
lTRect.Left = Left + ConvertPixelsToLogicalX(pRect.Left)
lTRect.Top = Top - ConvertPixelsToLogicalY(pRect.Top)
lTRect.Right = Left + ConvertPixelsToLogicalX(pRect.Right + 1)
lTRect.Bottom = Top - ConvertPixelsToLogicalY(pRect.Bottom + 1)
RectValidate lTRect
ConvertRectToTRectangle = lTRect
End Function

Public Function ConvertRelativeToLogicalX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertRelativeToLogicalX = mBoundary.Left + (value * (mBoundary.Right - mBoundary.Left) / 100#)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertRelativeToLogicalX" & "." & failpoint
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
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertRelativeToLogicalY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertRelativeToCounterDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertRelativeToCounterDistanceX = ((100# - value) * CDbl(mCanvas.Width)) / 100# / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertRelativeToCounterDistanceX" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertRelativeToCounterDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertRelativeToCounterDistanceY = ((100# - value) * CDbl(mCanvas.Height)) / 100# / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertRelativeToCounterDistanceY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertRelativeToDistanceX(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertRelativeToDistanceX = (value * CDbl(mCanvas.Width)) / 100# / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertRelativeToDistanceX" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertRelativeToDistanceY(ByVal value As Double) As Double
Dim failpoint As Long
On Error GoTo Err

ConvertRelativeToDistanceY = (value * CDbl(mCanvas.Height)) * 100# / TwipsPerCm

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConvertRelativeToDistanceY" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Function ConvertTRectangleToRect( _
                ByRef pTRect As TRectangle) As RECT
Dim lRect As RECT

lRect.Left = ConvertLogicalToPixelsX(pTRect.Left - Left)
lRect.Top = ConvertLogicalToPixelsY(Top - pTRect.Top)
lRect.Right = ConvertLogicalToPixelsX(pTRect.Right - Left) + 1
lRect.Bottom = ConvertLogicalToPixelsY(Top - pTRect.Bottom) + 1

ConvertTRectangleToRect = lRect
End Function

Friend Sub Finish()
gLogger.Log LogLevelHighDetail, "Viewport finished"
mCanvas.Finish
Set mCanvas = Nothing
End Sub

Public Function NewDimension(ByVal X As Double, _
                        ByVal Y As Double) As Dimension
Dim failpoint As Long
On Error GoTo Err

Set NewDimension = New Dimension
NewDimension.Viewport = Me
NewDimension.X = X
NewDimension.Y = Y

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "NewDimension" & "." & failpoint
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
NewPoint.Viewport = Me
NewPoint.X = X
NewPoint.Y = Y
NewPoint.CoordinateSystemX = coordSystemX
NewPoint.CoordinateSystemY = coordSystemY

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "NewPoint" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Function

Public Sub PaintBackground()
Dim failpoint As Long
On Error GoTo Err

mGradientFill.Fill mCanvas.hDC

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "PaintBackground" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Function ZOrder( _
                ByVal position As Integer)
mCanvas.ZOrder position
End Function

Public Sub UndrawArea( _
                ByRef undrawnArea As TRectangle)
Canvas.SetClippingRegion ConvertTRectangleToRect(undrawnArea)
PaintBackground
Canvas.ClearClippingRegion
End Sub


'================================================================================
' Helper Functions
'================================================================================

Private Sub checkValidDimensions()
Dim midPoint As Double
Dim range As Double

RectValidate mBoundary
If mBoundary.isValid Then
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
End Sub

Private Sub setGradientFillArea()
Dim lRect As RECT
lRect.Right = mCanvas.Width / Screen.TwipsPerPixelX
lRect.Bottom = mCanvas.Height / Screen.TwipsPerPixelY
mGradientFill.Area = lRect
End Sub

Private Sub SetScale()
If Not mCanvas Is Nothing Then mCanvas.SetScale mBoundary

'mGaugeX = mSurface.ScaleWidth / mSurface.ScaleX(mSurface.ScaleWidth, ScaleModeConstants.vbUser, ScaleModeConstants.vbPixels)
mGaugeX = ((mBoundary.Right - mBoundary.Left) * CDbl(Screen.TwipsPerPixelX)) / CDbl(mCanvas.Width)
'mGaugeY = -mSurface.ScaleHeight / mSurface.ScaleY(mSurface.ScaleHeight, ScaleModeConstants.vbUser, ScaleModeConstants.vbPixels)
mGaugeY = -((mBoundary.Bottom - mBoundary.Top) * CDbl(Screen.TwipsPerPixelY)) / CDbl(mCanvas.Height)

setGradientFillArea
End Sub



