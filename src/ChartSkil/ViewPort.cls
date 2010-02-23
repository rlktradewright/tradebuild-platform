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

Private mTransform As XForm

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mGradientFill = New GradientFill
End Sub

Private Sub Class_Terminate()
'gLogger.Log LogLevelHighDetail, "Viewport terminated (" & gRegionTypeToString(mRegionType) & ")"
Debug.Print "Viewport terminated (" & gRegionTypeToString(mRegionType) & ")"
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let BackColor(ByVal value As Long)
mGradientFill.Color = value
End Property

Public Property Get BackColor() As Long
Const ProcName As String = "BackColor"
Dim failpoint As String
On Error GoTo Err

BackColor = mGradientFill.Color

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let Bottom(ByVal value As Double)
Const ProcName As String = "Bottom"
Dim failpoint As String
On Error GoTo Err

mBoundary.Bottom = value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property

SetScale
RaiseEvent BoundaryChanged

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Bottom() As Double
Const ProcName As String = "Bottom"
Dim failpoint As String
On Error GoTo Err

Bottom = mBoundary.Bottom

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Boundary() As TRectangle
Boundary = mBoundary
End Property

Friend Property Let Canvas(ByRef value As Canvas)
Const ProcName As String = "Canvas"
Dim failpoint As String
On Error GoTo Err

Set mCanvas = value
mCanvas.Viewport = Me
setGradientFillArea

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Dim failpoint As String
On Error GoTo Err

GradientFillColors = mGradientFill.Colors

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let GradientFillColors(ByRef value() As Long)
Dim ar() As Long
Const ProcName As String = "GradientFillColors"
Dim failpoint As String
On Error GoTo Err

ar = value
mGradientFill.Colors = ar

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let Height(ByVal value As Double)
Const ProcName As String = "Height"
Dim failpoint As String
On Error GoTo Err

mBoundary.Bottom = mBoundary.Top - value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
SetScale

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Height() As Double
Const ProcName As String = "Height"
Dim failpoint As String
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get HeightCm() As Double
Const ProcName As String = "HeightCm"
Dim failpoint As String
On Error GoTo Err

HeightCm = mCanvas.Height / TwipsPerCm

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let Left(ByVal value As Double)
Const ProcName As String = "Left"
Dim failpoint As String
On Error GoTo Err

mBoundary.Left = value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
SetScale
RaiseEvent BoundaryChanged

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Left() As Double
Const ProcName As String = "Left"
Dim failpoint As String
On Error GoTo Err

Left = mBoundary.Left

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let MouseIcon(ByVal value As IPictureDisp)
Const ProcName As String = "MouseIcon"
Dim failpoint As String
On Error GoTo Err

mCanvas.MouseIcon = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let MousePointer(ByVal value As VBRUN.MousePointerConstants)
Const ProcName As String = "MousePointer"
Dim failpoint As String
On Error GoTo Err

mCanvas.MousePointer = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get MousePointer() As VBRUN.MousePointerConstants
Const ProcName As String = "MousePointer"
Dim failpoint As String
On Error GoTo Err

MousePointer = mCanvas.MousePointer

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let RegionType(ByVal value As RegionTypes)
mRegionType = value
End Property

Friend Sub NotifyResizedY()
' notifies that the physical Height of the Surface has changed.
Const ProcName As String = "NotifyResizedY"
Dim failpoint As String
On Error GoTo Err

If Not mBoundary.isValid Then Exit Sub
SetScale

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Property Let Right(ByVal value As Double)
Const ProcName As String = "Right"
Dim failpoint As String
On Error GoTo Err

mBoundary.Right = value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
SetScale
RaiseEvent BoundaryChanged

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Right() As Double
Const ProcName As String = "Right"
Dim failpoint As String
On Error GoTo Err

Right = mBoundary.Right

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let Top(ByVal value As Double)
Const ProcName As String = "Top"
Dim failpoint As String
On Error GoTo Err

mBoundary.Top = value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
SetScale
RaiseEvent BoundaryChanged

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Top() As Double
Top = mBoundary.Top
End Property

Public Property Get Width() As Double
Const ProcName As String = "Width"
Dim failpoint As String
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let Width(ByVal value As Double)
Const ProcName As String = "Width"
Dim failpoint As String
On Error GoTo Err

mBoundary.Left = mBoundary.Right - value
checkValidDimensions
If Not mBoundary.isValid Then Exit Property
SetScale

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get WidthCm() As Double
Const ProcName As String = "WidthCm"
Dim failpoint As String
On Error GoTo Err

WidthCm = mCanvas.Width / TwipsPerCm

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Clear()
Const ProcName As String = "Clear"
Dim failpoint As String
On Error GoTo Err

mCanvas.Clear

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Function CompareTo( _
                ByVal GaugeX As Double, _
                ByVal GaugeY As Double, _
                ByRef pRect As TRectangle) As ViewportComparisonCodes
Const ProcName As String = "CompareTo"
Dim failpoint As String
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertCoordsLogicalToPixelsX(ByVal value As Double) As Long
Const ProcName As String = "ConvertCoordsLogicalToPixelsX"
Dim failpoint As String
On Error GoTo Err

ConvertCoordsLogicalToPixelsX = ConvertLogicalToPixelsX(value - mBoundary.Left)
ConvertCoordsLogicalToPixelsX = ConvertCoordsLogicalToPixelsX - mTransform.eDx
Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertCoordsLogicalToPixelsY(ByVal value As Double) As Long
Const ProcName As String = "ConvertCoordsLogicalToPixelsY"
Dim failpoint As String
On Error GoTo Err

ConvertCoordsLogicalToPixelsY = ConvertLogicalToPixelsY(mBoundary.Top - value)
ConvertCoordsLogicalToPixelsY = ConvertCoordsLogicalToPixelsY - mTransform.eDy

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertCounterDistanceToDistanceX(ByVal value As Double) As Double
Const ProcName As String = "ConvertCounterDistanceToDistanceX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertCounterDistanceToDistanceX = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertCounterDistanceToDistanceX = MinusInfinityDouble
Else
    ConvertCounterDistanceToDistanceX = CDbl(mCanvas.Width) / TwipsPerCm - value
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertCounterDistanceToDistanceY(ByVal value As Double) As Double
Const ProcName As String = "ConvertCounterDistanceToDistanceY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertCounterDistanceToDistanceY = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertCounterDistanceToDistanceY = MinusInfinityDouble
Else
    ConvertCounterDistanceToDistanceY = CDbl(mCanvas.Height) / TwipsPerCm - value
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertCounterDistanceToLogicalX(ByVal value As Double) As Double
Const ProcName As String = "ConvertCounterDistanceToLogicalX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertCounterDistanceToLogicalX = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertCounterDistanceToLogicalX = MinusInfinityDouble
Else
    ConvertCounterDistanceToLogicalX = (mBoundary.Right - mBoundary.Left) - ((value * TwipsPerCm) / CDbl(mCanvas.Width) * (mBoundary.Right - mBoundary.Left))
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertCounterDistanceToLogicalY(ByVal value As Double) As Double
Const ProcName As String = "ConvertCounterDistanceToLogicalY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertCounterDistanceToLogicalY = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertCounterDistanceToLogicalY = MinusInfinityDouble
Else
    ConvertCounterDistanceToLogicalY = (mBoundary.Top - mBoundary.Bottom) - ((value * TwipsPerCm) / CDbl(mCanvas.Height) * (mBoundary.Top - mBoundary.Bottom))
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertCounterDistanceToRelativeX(ByVal value As Double) As Double
Const ProcName As String = "ConvertCounterDistanceToRelativeX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertCounterDistanceToRelativeX = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertCounterDistanceToRelativeX = MinusInfinityDouble
Else
    ConvertCounterDistanceToRelativeX = 100# - 100# * value * TwipsPerCm / CDbl(mCanvas.Width)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertCounterDistanceToRelativeY(ByVal value As Double) As Double
Const ProcName As String = "ConvertCounterDistanceToRelativeY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertCounterDistanceToRelativeY = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertCounterDistanceToRelativeY = MinusInfinityDouble
Else
    ConvertCounterDistanceToRelativeY = 100# - 100# * value * TwipsPerCm / CDbl(mCanvas.Height)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertDistanceToCounterDistanceX(ByVal value As Double) As Double
Const ProcName As String = "ConvertDistanceToCounterDistanceX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertDistanceToCounterDistanceX = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertDistanceToCounterDistanceX = MinusInfinityDouble
Else
    ConvertDistanceToCounterDistanceX = CDbl(mCanvas.Width) / TwipsPerCm - value
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertDistanceToCounterDistanceY(ByVal value As Double) As Double
Const ProcName As String = "ConvertDistanceToCounterDistanceY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertDistanceToCounterDistanceY = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertDistanceToCounterDistanceY = MinusInfinityDouble
Else
    ConvertDistanceToCounterDistanceY = CDbl(mCanvas.Height) / TwipsPerCm - value
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertDistanceToLogicalX(ByVal value As Double) As Double
Const ProcName As String = "ConvertDistanceToLogicalX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertDistanceToLogicalX = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertDistanceToLogicalX = MinusInfinityDouble
Else
    ConvertDistanceToLogicalX = (value * TwipsPerCm) / CDbl(mCanvas.Width) * (mBoundary.Right - mBoundary.Left)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertDistanceToLogicalY(ByVal value As Double) As Double
Const ProcName As String = "ConvertDistanceToLogicalY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertDistanceToLogicalY = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertDistanceToLogicalY = MinusInfinityDouble
Else
    ConvertDistanceToLogicalY = (value * TwipsPerCm) / CDbl(mCanvas.Height) * (mBoundary.Top - mBoundary.Bottom)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertDistanceToPixelsX(ByVal value As Double) As Long
Const ProcName As String = "ConvertDistanceToPixelsX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertDistanceToPixelsX = PlusInfinityLong
ElseIf value = MinusInfinityDouble Then
    ConvertDistanceToPixelsX = MinusInfinityLong
Else
    ConvertDistanceToPixelsX = Round((value * TwipsPerCm) / Screen.TwipsPerPixelX, 0)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertDistanceToPixelsY(ByVal value As Double) As Long
Const ProcName As String = "ConvertDistanceToPixelsY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertDistanceToPixelsY = PlusInfinityLong
ElseIf value = MinusInfinityDouble Then
    ConvertDistanceToPixelsY = MinusInfinityLong
Else
    ConvertDistanceToPixelsY = Round((value * TwipsPerCm) / Screen.TwipsPerPixelY, 0)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertDistanceToRelativeX(ByVal value As Double) As Double
Const ProcName As String = "ConvertDistanceToRelativeX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertDistanceToRelativeX = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertDistanceToRelativeX = MinusInfinityDouble
Else
    ConvertDistanceToRelativeX = 100# * value * TwipsPerCm / CDbl(mCanvas.Width)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertDistanceToRelativeY(ByVal value As Double) As Double
Const ProcName As String = "ConvertDistanceToRelativeY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertDistanceToRelativeY = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertDistanceToRelativeY = MinusInfinityDouble
Else
    ConvertDistanceToRelativeY = 100# * value * TwipsPerCm / CDbl(mCanvas.Height)
End If
Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertLogicalToCounterDistanceX(ByVal value As Double) As Double
Const ProcName As String = "ConvertLogicalToCounterDistanceX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertLogicalToCounterDistanceX = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertLogicalToCounterDistanceX = MinusInfinityDouble
Else
    ConvertLogicalToCounterDistanceX = (mBoundary.Right - value) / (mBoundary.Right - mBoundary.Left) * mCanvas.Width / TwipsPerCm
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertLogicalToCounterDistanceY(ByVal value As Double) As Double
Const ProcName As String = "ConvertLogicalToCounterDistanceY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertLogicalToCounterDistanceY = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertLogicalToCounterDistanceY = MinusInfinityDouble
Else
    ConvertLogicalToCounterDistanceY = (mBoundary.Top - value) / (mBoundary.Top - mBoundary.Bottom) * CDbl(mCanvas.Height) / TwipsPerCm
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertLogicalToDistanceX(ByVal value As Double) As Double
Const ProcName As String = "ConvertLogicalToDistanceX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertLogicalToDistanceX = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertLogicalToDistanceX = MinusInfinityDouble
Else
    ConvertLogicalToDistanceX = (value - mBoundary.Left) / (mBoundary.Right - mBoundary.Left) * CDbl(mCanvas.Width) / TwipsPerCm
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertLogicalToDistanceY(ByVal value As Double) As Double
Const ProcName As String = "ConvertLogicalToDistanceY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertLogicalToDistanceY = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertLogicalToDistanceY = MinusInfinityDouble
Else
    ConvertLogicalToDistanceY = (value - mBoundary.Bottom) / (mBoundary.Top - mBoundary.Bottom) * CDbl(mCanvas.Height) / TwipsPerCm
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertLogicalToPixelsX(ByVal value As Double) As Long
Const ProcName As String = "ConvertLogicalToPixelsX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertLogicalToPixelsX = PlusInfinityLong
ElseIf value = MinusInfinityDouble Then
    ConvertLogicalToPixelsX = MinusInfinityLong
Else
    ConvertLogicalToPixelsX = mCanvas.Surface.ScaleX(value, vbUser, vbPixels)
    'ConvertLogicalToPixelsX = value / mGaugeX
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertLogicalToPixelsY(ByVal value As Double) As Long
Const ProcName As String = "ConvertLogicalToPixelsY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertLogicalToPixelsY = PlusInfinityLong
ElseIf value = MinusInfinityDouble Then
    ConvertLogicalToPixelsY = MinusInfinityLong
Else
    ConvertLogicalToPixelsY = -Canvas.Surface.ScaleY(value, vbUser, vbPixels)
    'ConvertLogicalToPixelsY = value / mGaugeY
End If
Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertLogicalToRelativeX(ByVal value As Double) As Double
Const ProcName As String = "ConvertLogicalToRelativeX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertLogicalToRelativeX = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertLogicalToRelativeX = MinusInfinityDouble
Else
    ConvertLogicalToRelativeX = (value - mBoundary.Left) / (mBoundary.Right - mBoundary.Left) * 100#
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertLogicalToRelativeY(ByVal value As Double) As Double
Const ProcName As String = "ConvertLogicalToRelativeY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertLogicalToRelativeY = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertLogicalToRelativeY = MinusInfinityDouble
Else
    ConvertLogicalToRelativeY = (value - mBoundary.Bottom) / (mBoundary.Top - mBoundary.Bottom) * 100#
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertPixelsToDistanceX(ByVal value As Long) As Double
Const ProcName As String = "ConvertPixelsToDistanceX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityLong Then
    ConvertPixelsToDistanceX = PlusInfinityDouble
ElseIf value = MinusInfinityLong Then
    ConvertPixelsToDistanceX = MinusInfinityDouble
Else
    ConvertPixelsToDistanceX = value * Screen.TwipsPerPixelX / TwipsPerCm
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertPixelsToDistanceY(ByVal value As Long) As Double
Const ProcName As String = "ConvertPixelsToDistanceY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityLong Then
    ConvertPixelsToDistanceY = PlusInfinityDouble
ElseIf value = MinusInfinityLong Then
    ConvertPixelsToDistanceY = MinusInfinityDouble
Else
    ConvertPixelsToDistanceY = value * Screen.TwipsPerPixelY / TwipsPerCm
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertPixelsToLogicalX(ByVal value As Long) As Double
Const ProcName As String = "ConvertPixelsToLogicalX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityLong Then
    ConvertPixelsToLogicalX = PlusInfinityDouble
ElseIf value = MinusInfinityLong Then
    ConvertPixelsToLogicalX = MinusInfinityDouble
Else
    ConvertPixelsToLogicalX = Canvas.Surface.ScaleX(value, vbPixels, vbUser)
    'ConvertPixelsToLogicalX = CDbl(value) * mGaugeX
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertPixelsToLogicalY(ByVal value As Long) As Double
Const ProcName As String = "ConvertPixelsToLogicalY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityLong Then
    ConvertPixelsToLogicalY = PlusInfinityDouble
ElseIf value = MinusInfinityLong Then
    ConvertPixelsToLogicalY = MinusInfinityDouble
Else
    ConvertPixelsToLogicalY = -Canvas.Surface.ScaleY(value, vbPixels, vbUser)
    'ConvertPixelsToLogicalY = CDbl(value) * mGaugeY
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertPointsToWin32Points( _
                ByRef points() As Point) As W32Point()
Dim w32Points() As W32Point
Const ProcName As String = "ConvertPointsToWin32Points"
Dim failpoint As String
On Error GoTo Err

ReDim w32Points(UBound(points)) As W32Point
Dim i As Long

For i = 0 To UBound(points)
    w32Points(i).X = ConvertCoordsLogicalToPixelsX(points(i).XLogical)
    w32Points(i).Y = ConvertCoordsLogicalToPixelsY(points(i).yLogical)
Next

ConvertPointsToWin32Points = w32Points

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertRectToTRectangle( _
                ByRef pRect As RECT) As TRectangle
Dim lTRect As TRectangle
Const ProcName As String = "ConvertRectToTRectangle"
Dim failpoint As String
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertRelativeToCounterDistanceX(ByVal value As Double) As Double
Const ProcName As String = "ConvertRelativeToCounterDistanceX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertRelativeToCounterDistanceX = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertRelativeToCounterDistanceX = MinusInfinityDouble
Else
    ConvertRelativeToCounterDistanceX = ((100# - value) * CDbl(mCanvas.Width)) / 100# / TwipsPerCm
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertRelativeToCounterDistanceY(ByVal value As Double) As Double
Const ProcName As String = "ConvertRelativeToCounterDistanceY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertRelativeToCounterDistanceY = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertRelativeToCounterDistanceY = MinusInfinityDouble
Else
    ConvertRelativeToCounterDistanceY = ((100# - value) * CDbl(mCanvas.Height)) / 100# / TwipsPerCm
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertRelativeToDistanceX(ByVal value As Double) As Double
Const ProcName As String = "ConvertRelativeToDistanceX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertRelativeToDistanceX = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertRelativeToDistanceX = MinusInfinityDouble
Else
    ConvertRelativeToDistanceX = (value * CDbl(mCanvas.Width)) / 100# / TwipsPerCm
End If
Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertRelativeToDistanceY(ByVal value As Double) As Double
Const ProcName As String = "ConvertRelativeToDistanceY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertRelativeToDistanceY = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertRelativeToDistanceY = MinusInfinityDouble
Else
    ConvertRelativeToDistanceY = (value * CDbl(mCanvas.Height)) * 100# / TwipsPerCm
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertRelativeToLogicalX(ByVal value As Double) As Double
Const ProcName As String = "ConvertRelativeToLogicalX"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertRelativeToLogicalX = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertRelativeToLogicalX = MinusInfinityDouble
Else
    'ConvertRelativeToLogicalX = mBoundary.Left + (value * (mBoundary.Right - mBoundary.Left) / 100#)
    ConvertRelativeToLogicalX = (value * (mBoundary.Right - mBoundary.Left) / 100#)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertRelativeToLogicalY(ByVal value As Double) As Double
Const ProcName As String = "ConvertRelativeToLogicalY"
Dim failpoint As String
On Error GoTo Err

If value = PlusInfinityDouble Then
    ConvertRelativeToLogicalY = PlusInfinityDouble
ElseIf value = MinusInfinityDouble Then
    ConvertRelativeToLogicalY = MinusInfinityDouble
Else
    'ConvertRelativeToLogicalY = mBoundary.Bottom + (value * (mBoundary.Top - mBoundary.Bottom) / 100#)
    ConvertRelativeToLogicalY = (value * (mBoundary.Top - mBoundary.Bottom) / 100#)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ConvertTRectangleToRect( _
                ByRef pTRect As TRectangle) As RECT
Dim lRect As RECT

Const ProcName As String = "ConvertTRectangleToRect"
Dim failpoint As String
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub Finish()
Const ProcName As String = "Finish"
Dim failpoint As String
On Error GoTo Err

gLogger.Log LogLevelHighDetail, "Viewport finished"
mCanvas.Finish
Set mCanvas = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function NewDimension( _
                ByVal X As Double, _
                ByVal Y As Double, _
                Optional ByVal coordSystemX As CoordinateSystems = CoordsDistance, _
                Optional ByVal coordSystemY As CoordinateSystems = CoordsDistance) As Dimension
Const ProcName As String = "NewDimension"
Dim failpoint As String
On Error GoTo Err

Set NewDimension = New Dimension
NewDimension.Initialise X, Y, coordSystemX, coordSystemY, Me
Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function NewDimensionX( _
                ByVal X As Double, _
                Optional ByVal coordSystemX As CoordinateSystems = CoordsDistance) As DimensionX
Const ProcName As String = "NewDimensionX"
Dim failpoint As String
On Error GoTo Err

Set NewDimensionX = New DimensionX
NewDimensionX.Initialise X, coordSystemX, Me
Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function NewDimensionY( _
                ByVal Y As Double, _
                Optional ByVal coordSystemY As CoordinateSystems = CoordsDistance) As DimensionY
Const ProcName As String = "NewDimensionX"
Dim failpoint As String
On Error GoTo Err

Set NewDimensionY = New DimensionY
NewDimensionY.Initialise Y, coordSystemY, Me
Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function NewPoint(ByVal X As Double, _
                        ByVal Y As Double, _
                        Optional ByVal coordSystemX As CoordinateSystems = CoordsLogical, _
                        Optional ByVal coordSystemY As CoordinateSystems = CoordsLogical, _
                        Optional ByVal Offset As Dimension) As Point
Const ProcName As String = "NewPoint"
Dim failpoint As String
On Error GoTo Err

Set NewPoint = New Point
NewPoint.Initialise X, Y, coordSystemX, coordSystemY, Offset, Me

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub PaintBackground()
Const ProcName As String = "PaintBackground"
Dim failpoint As String
On Error GoTo Err

mGradientFill.Fill mCanvas.hDC

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub Reset()
ModifyWorldTransform mCanvas.hDC, 0, MWT_IDENTITY
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
SetWorldTransform mCanvas.hDC, VarPtr(mTransform)
End Sub

Friend Function ZOrder( _
                ByVal Position As Integer)
Const ProcName As String = "ZOrder"
Dim failpoint As String
On Error GoTo Err

mCanvas.ZOrder Position

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub UndrawArea( _
                ByRef undrawnArea As TRectangle)
Const ProcName As String = "UndrawArea"
Dim failpoint As String
On Error GoTo Err

Canvas.SetClippingRegion ConvertTRectangleToRect(undrawnArea)
PaintBackground
Canvas.ClearClippingRegion

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub


'================================================================================
' Helper Functions
'================================================================================

Private Sub checkValidDimensions()
Dim midPoint As Double
Dim range As Double

Const ProcName As String = "checkValidDimensions"
Dim failpoint As String
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setGradientFillArea()
Dim lRect As RECT
Const ProcName As String = "setGradientFillArea"
Dim failpoint As String
On Error GoTo Err

lRect.Right = mCanvas.Width / Screen.TwipsPerPixelX
lRect.Bottom = mCanvas.Height / Screen.TwipsPerPixelY
mGradientFill.Area = lRect

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub SetScale()
Const ProcName As String = "SetScale"
Dim failpoint As String
On Error GoTo Err

If Not mCanvas Is Nothing Then mCanvas.SetScale mBoundary

mGaugeX = ((mBoundary.Right - mBoundary.Left) * CDbl(Screen.TwipsPerPixelX)) / CDbl(mCanvas.Width)
mGaugeY = -((mBoundary.Bottom - mBoundary.Top) * CDbl(Screen.TwipsPerPixelY)) / CDbl(mCanvas.Height)

setGradientFillArea

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub



