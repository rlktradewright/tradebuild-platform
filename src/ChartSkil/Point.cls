VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "Point"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mX As Double
Private mY As Double
Private mCoordSystemX As CoordinateSystems
Private mCoordSystemY As CoordinateSystems
Private mViewport As Viewport
Private mOffset As Dimension

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Get Viewport() As Viewport
Set Viewport = mViewport
End Property

Public Property Get CoordinateSystemX() As CoordinateSystems
CoordinateSystemX = mCoordSystemX
End Property

Public Property Get CoordinateSystemY() As CoordinateSystems
CoordinateSystemY = mCoordSystemY
End Property

Public Property Get Offset() As Dimension
Set Offset = mOffset
End Property

Public Property Get X() As Double
X = mX
End Property

Public Property Get XLogical() As Double
Const ProcName As String = "XLogical"
Dim failpoint As String
On Error GoTo Err

Select Case mCoordSystemX
Case CoordsLogical
    XLogical = mX
Case CoordsRelative
    XLogical = mViewport.ConvertRelativeToLogicalX(mX) + mViewport.Boundary.Left
Case CoordsDistance
    XLogical = mViewport.ConvertDistanceToLogicalX(mX) + mViewport.Boundary.Left
Case CoordsCounterDistance
    XLogical = mViewport.ConvertCounterDistanceToLogicalY(mX) + mViewport.Boundary.Left
End Select

If Not mOffset Is Nothing Then XLogical = XLogical + mOffset.XLogical

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Y() As Double
Y = mY
End Property

Public Property Get yLogical() As Double
Const ProcName As String = "YLogical"
Dim failpoint As String
On Error GoTo Err

Select Case mCoordSystemY
Case CoordsLogical
    yLogical = mY
Case CoordsRelative
    yLogical = mViewport.ConvertRelativeToLogicalY(mY) + mViewport.Boundary.Bottom
Case CoordsDistance
    yLogical = mViewport.ConvertDistanceToLogicalY(mY) + mViewport.Boundary.Bottom
Case CoordsCounterDistance
    yLogical = mViewport.ConvertCounterDistanceToLogicalY(mY) + mViewport.Boundary.Bottom
End Select

If Not mOffset Is Nothing Then yLogical = yLogical + mOffset.yLogical

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'================================================================================
' Methods
'================================================================================

Public Function Equals(ByVal aPoint As Point) As Boolean
Const ProcName As String = "Equals"
Dim failpoint As String
On Error GoTo Err

If aPoint Is Nothing Then Exit Function
If aPoint Is Me Then
    Equals = True
    Exit Function
End If
With aPoint
    If Not .Viewport Is mViewport Then Exit Function
    If .CoordinateSystemX <> mCoordSystemX Then Exit Function
    If .CoordinateSystemY <> mCoordSystemY Then Exit Function
    If .X <> mX Then Exit Function
    If .Y <> mY Then Exit Function
End With
Equals = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ChangeCoordinateSystem( _
                ByVal pCoordSystemX As CoordinateSystems, _
                ByVal pCoordSystemy As CoordinateSystems) As Point
Dim newX As Double
Dim newY As Double

Const ProcName As String = "ChangeCoordinateSystem"
On Error GoTo Err

If pCoordSystemX = mCoordSystemX And pCoordSystemy = mCoordSystemY Then
    Set ChangeCoordinateSystem = Me
    Exit Function
End If

newX = transformCoordX(pCoordSystemX)
newY = transformCoordY(pCoordSystemy)

Set ChangeCoordinateSystem = mViewport.NewPoint(newX, newY, pCoordSystemX, pCoordSystemy)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub Initialise( _
                ByVal X As Double, _
                ByVal Y As Double, _
                ByVal coordSystemX As CoordinateSystems, _
                ByVal coordSystemY As CoordinateSystems, _
                ByVal Offset As Dimension, _
                ByVal pViewport As Viewport)
mX = X
mY = Y
mCoordSystemX = coordSystemX
mCoordSystemY = coordSystemY
Set mOffset = Offset
Set mViewport = pViewport
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function transformCoordX(value As CoordinateSystems) As Double
Const ProcName As String = "transformCoordX"
Dim failpoint As String
On Error GoTo Err

If mCoordSystemX = value Then
    transformCoordX = mX
    Exit Function
End If

Select Case mCoordSystemX
Case CoordsLogical
    If value = CoordsCounterDistance Then
        transformCoordX = mViewport.ConvertLogicalToCounterDistanceX(mX)
    ElseIf value = CoordsDistance Then
        transformCoordX = mViewport.ConvertLogicalToDistanceX(mX)
    ElseIf value = CoordsRelative Then
        transformCoordX = mViewport.ConvertLogicalToRelativeX(mX)
    End If
Case CoordsRelative
    If value = CoordsLogical Then
        transformCoordX = XLogical
    ElseIf value = CoordsCounterDistance Then
        transformCoordX = mViewport.ConvertRelativeToCounterDistanceX(mX)
    ElseIf value = CoordsDistance Then
        transformCoordX = mViewport.ConvertRelativeToDistanceX(mX)
    End If
Case CoordsDistance
    If value = CoordsLogical Then
        transformCoordX = XLogical
    ElseIf value = CoordsCounterDistance Then
        transformCoordX = mViewport.ConvertDistanceToCounterDistanceX(mX)
    ElseIf value = CoordsRelative Then
        transformCoordX = mViewport.ConvertDistanceToRelativeX(mX)
    End If
Case CoordsCounterDistance
    If value = CoordsLogical Then
        transformCoordX = XLogical
    ElseIf value = CoordsDistance Then
        transformCoordX = mViewport.ConvertCounterDistanceToDistanceX(mX)
    ElseIf value = CoordsRelative Then
        transformCoordX = mViewport.ConvertCounterDistanceToRelativeX(mX)
    End If
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function transformCoordY(value As CoordinateSystems) As Double
Const ProcName As String = "transformCoordY"
Dim failpoint As String
On Error GoTo Err

If mCoordSystemY = value Then
    transformCoordY = mY
    Exit Function
End If

Select Case mCoordSystemY
Case CoordsLogical
    If value = CoordsCounterDistance Then
        transformCoordY = mViewport.ConvertLogicalToCounterDistanceY(mY)
    ElseIf value = CoordsDistance Then
        transformCoordY = mViewport.ConvertLogicalToDistanceY(mY)
    ElseIf value = CoordsRelative Then
        transformCoordY = mViewport.ConvertLogicalToRelativeY(mY)
    End If
Case CoordsRelative
    If value = CoordsLogical Then
        transformCoordY = yLogical
    ElseIf value = CoordsCounterDistance Then
        transformCoordY = mViewport.ConvertRelativeToCounterDistanceY(mY)
    ElseIf value = CoordsDistance Then
        transformCoordY = mViewport.ConvertRelativeToDistanceY(mY)
    End If
Case CoordsDistance
    If value = CoordsLogical Then
        transformCoordY = yLogical
    ElseIf value = CoordsCounterDistance Then
        transformCoordY = mViewport.ConvertDistanceToCounterDistanceY(mY)
    ElseIf value = CoordsRelative Then
        transformCoordY = mViewport.ConvertDistanceToRelativeY(mY)
    End If
Case CoordsCounterDistance
    If value = CoordsLogical Then
        transformCoordY = yLogical
    ElseIf value = CoordsDistance Then
        transformCoordY = mViewport.ConvertCounterDistanceToDistanceY(mY)
    ElseIf value = CoordsRelative Then
        transformCoordY = mViewport.ConvertCounterDistanceToRelativeY(mY)
    End If
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function


