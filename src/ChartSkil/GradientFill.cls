VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GradientFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GradientFill"

'================================================================================
' External FUnction Declarations
'================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mArea                               As GDI_RECT
Private mGradientFillMode                   As Long
Private mColors()                           As Long
Private mVertices()                         As TRIVERTEX
Private mMeshTriangle()                     As GRADIENT_TRIANGLE
Private mMeshRect()                         As GRADIENT_RECT

Private mRecalcRequired                     As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mColors(0) As Long
mColors(0) = vbWhite
mArea.Left = 0
mArea.Right = 1
mArea.Bottom = 0
mArea.Top = 1
calculateGradientFillParameters
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

Public Property Let Area( _
                ByRef Value As GDI_RECT)
mArea = Value
mRecalcRequired = True
End Property

Public Property Get Area() As GDI_RECT
Area = mArea
End Property

Public Property Let Color(ByVal Value As Long)
Const ProcName As String = "Color"

On Error GoTo Err

ReDim mColors(0) As Long
mColors(0) = Value
calculateGradientFillParameters

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Color() As Long
Color = mColors(0)
End Property

Public Property Let Colors(ByRef Value() As Long)
Dim numColors As Long

Const ProcName As String = "Colors"

On Error GoTo Err

On Error Resume Next
numColors = UBound(Value) + 1
On Error GoTo Err

If numColors = 0 Then
    Erase mColors
    Exit Property
End If

If numColors < 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "At least 1 background gradient fill Color must be supplied"
End If

If numColors > 5 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "At most 5 background gradient fill colors may be supplied"
End If

mColors = Value

calculateGradientFillParameters

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Colors() As Long()
Colors = mColors
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Fill( _
                ByVal hDC As Long)
Const ProcName As String = "Fill"

On Error GoTo Err

If mRecalcRequired Then calculateGradientFillParameters

If mGradientFillMode = GRADIENT_FILL_TRIANGLE Then
    GradientFill hDC, _
                mVertices(0), _
                UBound(mVertices) + 1, _
                VarPtr(mMeshTriangle(0)), _
                UBound(mMeshTriangle) + 1, _
                GRADIENT_FILL_TRIANGLE
Else
    GradientFill hDC, _
                mVertices(0), _
                UBound(mVertices) + 1, _
                VarPtr(mMeshRect(0)), _
                UBound(mMeshRect) + 1, _
                GRADIENT_FILL_RECT_H
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calculateGradientFillParameters()
Const ProcName As String = "calculateGradientFillParameters"

On Error GoTo Err

Select Case UBound(mColors)
Case 0
    ReDim mVertices(1) As TRIVERTEX
    ReDim mMeshRect(0) As GRADIENT_RECT
    setVertex 0, mArea.Left, mArea.Top, 0
    setVertex 1, mArea.Right + 1, mArea.Bottom, 0
    setMeshRect 0, 0, 1
    mGradientFillMode = GRADIENT_FILL_RECT_H
Case 1
    ReDim mVertices(1) As TRIVERTEX
    ReDim mMeshRect(0) As GRADIENT_RECT
    setVertex 0, mArea.Left, mArea.Top, 0
    setVertex 1, mArea.Right + 1, mArea.Bottom, 1
    setMeshRect 0, 0, 1
    mGradientFillMode = GRADIENT_FILL_RECT_H
Case 2
    ReDim mVertices(4) As TRIVERTEX
    ReDim mMeshTriangle(3) As GRADIENT_TRIANGLE
    setVertex 0, mArea.Left, mArea.Top, 0
    setVertex 1, mArea.Right, mArea.Top, 1
    setVertex 2, (mArea.Left + mArea.Right + 1) / 2, (mArea.Top + mArea.Bottom) / 2, 2
    setVertex 3, mArea.Right + 1, mArea.Bottom, 1
    setVertex 4, mArea.Left, mArea.Bottom, 0
    setMeshTriangle 0, 0, 1, 2
    setMeshTriangle 1, 1, 3, 2
    setMeshTriangle 2, 3, 4, 2
    setMeshTriangle 3, 4, 0, 2
    mGradientFillMode = GRADIENT_FILL_TRIANGLE
Case 3
    ReDim mVertices(4) As TRIVERTEX
    ReDim mMeshTriangle(3) As GRADIENT_TRIANGLE
    setVertex 0, mArea.Left, mArea.Top, 0
    setVertex 1, mArea.Right + 1, mArea.Top, 1
    setVertex 2, mArea.Right + 1, mArea.Bottom, 2
    setVertex 3, mArea.Left, mArea.Bottom, 3
    setVertex 4, (mArea.Left + mArea.Right + 1) / 2, (mArea.Top + mArea.Bottom) / 2, 0, calculateMidColor
    setMeshTriangle 0, 0, 1, 4
    setMeshTriangle 1, 1, 2, 4
    setMeshTriangle 2, 2, 3, 4
    setMeshTriangle 3, 3, 0, 4
    mGradientFillMode = GRADIENT_FILL_TRIANGLE
Case 4
    ReDim mVertices(4) As TRIVERTEX
    ReDim mMeshTriangle(3) As GRADIENT_TRIANGLE
    setVertex 0, mArea.Left, mArea.Top, 0
    setVertex 1, mArea.Right + 1, mArea.Top, 1
    setVertex 2, mArea.Right + 1, mArea.Bottom, 2
    setVertex 3, mArea.Left, mArea.Bottom, 3
    setVertex 4, (mArea.Left + mArea.Right + 1) / 2, (mArea.Top + mArea.Bottom) / 2, 4
    setMeshTriangle 0, 0, 1, 4
    setMeshTriangle 1, 1, 2, 4
    setMeshTriangle 2, 2, 3, 4
    setMeshTriangle 3, 3, 0, 4
    mGradientFillMode = GRADIENT_FILL_TRIANGLE
End Select
mRecalcRequired = False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function calculateMidColor() As Long
Dim sumColors As Currency
Dim i As Long
Const ProcName As String = "calculateMidColor"

On Error GoTo Err

For i = 0 To UBound(mColors)
    sumColors = sumColors + mColors(i)
Next
calculateMidColor = CLng(sumColors / (UBound(mColors) + 1))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub setMeshRect( _
                ByVal index As Long, _
                ByVal upperLeft As Long, _
                ByVal lowerRight As Long)
Const ProcName As String = "setMeshRect"

On Error GoTo Err

mMeshRect(index).lowerRight = lowerRight
mMeshRect(index).upperLeft = upperLeft

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setMeshTriangle( _
                ByVal index As Long, _
                ByVal vertex1 As Long, _
                ByVal vertex2 As Long, _
                ByVal vertex3 As Long)
Const ProcName As String = "setMeshTriangle"

On Error GoTo Err

mMeshTriangle(index).vertex1 = vertex1
mMeshTriangle(index).vertex2 = vertex2
mMeshTriangle(index).vertex3 = vertex3

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setVertex( _
                ByRef index As Long, _
                ByVal X As Double, _
                ByVal Y As Double, _
                ByVal colorIndex As Long, _
                Optional ByVal Color As Long = -1)
Const ProcName As String = "setVertex"

On Error GoTo Err

mVertices(index).X = X
mVertices(index).Y = Y
If Color = -1 Then Color = mColors(colorIndex)
mVertices(index).Red = Color And &HFF&
mVertices(index).Green = (Color And &HFF00&) / &H100&
mVertices(index).Blue = (Color And &HFF0000) / &H10000
mVertices(index).Alpha = 0

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


