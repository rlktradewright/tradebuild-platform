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
Private mOffset As Size

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get CoordinateSystemX() As CoordinateSystems
CoordinateSystemX = mCoordSystemX
End Property

Public Property Get CoordinateSystemY() As CoordinateSystems
CoordinateSystemY = mCoordSystemY
End Property

Public Property Get Offset() As Size
Set Offset = mOffset
End Property

Public Property Get X() As Double
X = mX
End Property

Public Property Get Y() As Double
Y = mY
End Property

'================================================================================
' Methods
'================================================================================

Public Function ChangeCoordinateSystem( _
                ByVal pCoordSystemX As CoordinateSystems, _
                ByVal pCoordSystemY As CoordinateSystems, _
                ByVal pRegion As ChartRegion) As Point
Dim newX As Double
Dim newY As Double

Const ProcName As String = "ChangeCoordinateSystem"
On Error GoTo Err

If pCoordSystemX = Me.CoordinateSystemX And pCoordSystemY = Me.CoordinateSystemY Then
    Set ChangeCoordinateSystem = Me
    Exit Function
End If

newX = gTransformCoordX(Me.X, Me.CoordinateSystemX, pCoordSystemX, pRegion.ViewPort)
newY = gTransformCoordY(Me.Y, Me.CoordinateSystemY, pCoordSystemY, pRegion.ViewPort)

Set ChangeCoordinateSystem = gNewPoint(newX, newY, pCoordSystemX, pCoordSystemY)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Equals(ByVal aPoint As Point) As Boolean
Const ProcName As String = "Equals"

On Error GoTo Err

If aPoint Is Nothing Then Exit Function
If aPoint Is Me Then
    Equals = True
    Exit Function
End If
With aPoint
    If .CoordinateSystemX <> mCoordSystemX Then Exit Function
    If .CoordinateSystemY <> mCoordSystemY Then Exit Function
    If .X <> mX Then Exit Function
    If .Y <> mY Then Exit Function
End With
Equals = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal X As Double, _
                ByVal Y As Double, _
                ByVal coordSystemX As CoordinateSystems, _
                ByVal coordSystemY As CoordinateSystems, _
                ByVal Offset As Size)
mX = X
mY = Y
mCoordSystemX = coordSystemX
mCoordSystemY = coordSystemY
Set mOffset = Offset
End Sub

Public Function XLogical( _
                ByVal pRegion As ChartRegion) As Double
Const ProcName As String = "XLogical"

On Error GoTo Err

Select Case Me.CoordinateSystemX
Case CoordsLogical
    XLogical = Me.X
Case CoordsRelative
    XLogical = pRegion.ViewPort.ConvertRelativeToLogicalX(Me.X) + pRegion.ViewPort.Boundary.Left
Case CoordsDistance
    XLogical = pRegion.ViewPort.ConvertDistanceToLogicalX(Me.X) + pRegion.ViewPort.Boundary.Left
Case CoordsCounterDistance
    XLogical = pRegion.ViewPort.ConvertCounterDistanceToLogicalX(Me.X) + pRegion.ViewPort.Boundary.Left
End Select

If Not Me.Offset Is Nothing Then XLogical = XLogical + Me.Offset.WidthLogical(pRegion)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Property Get yLogical( _
                ByVal pRegion As ChartRegion) As Double
Const ProcName As String = "YLogical"

On Error GoTo Err

Select Case Me.CoordinateSystemY
Case CoordsLogical
    yLogical = Me.Y
Case CoordsRelative
    yLogical = pRegion.ViewPort.ConvertRelativeToLogicalY(Me.Y) + pRegion.ViewPort.Boundary.Bottom
Case CoordsDistance
    yLogical = pRegion.ViewPort.ConvertDistanceToLogicalY(Me.Y) + pRegion.ViewPort.Boundary.Bottom
Case CoordsCounterDistance
    yLogical = pRegion.ViewPort.ConvertCounterDistanceToLogicalY(Me.Y) + pRegion.ViewPort.Boundary.Bottom
End Select

If Not Me.Offset Is Nothing Then yLogical = yLogical + Me.Offset.HeightLogical(pRegion)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Helper Functions
'================================================================================


