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
                ByVal pViewport As Viewport) As Point
Dim newX As Double
Dim newY As Double

Const ProcName As String = "ChangeCoordinateSystem"
On Error GoTo Err

If pCoordSystemX = Me.CoordinateSystemX And pCoordSystemY = Me.CoordinateSystemY Then
    Set ChangeCoordinateSystem = Me
    Exit Function
End If

newX = gTransformCoordX(Me.X, Me.CoordinateSystemX, pCoordSystemX, pViewport)
newY = gTransformCoordY(Me.Y, Me.CoordinateSystemY, pCoordSystemY, pViewport)

Set ChangeCoordinateSystem = gNewPoint(newX, newY, pCoordSystemX, pCoordSystemY)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

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
                ByVal pViewport As Viewport) As Double
Const ProcName As String = "XLogical"
Dim failpoint As String
On Error GoTo Err

Select Case Me.CoordinateSystemX
Case CoordsLogical
    XLogical = Me.X
Case CoordsRelative
    XLogical = pViewport.ConvertRelativeToLogicalX(Me.X) + pViewport.Boundary.Left
Case CoordsDistance
    XLogical = pViewport.ConvertDistanceToLogicalX(Me.X) + pViewport.Boundary.Left
Case CoordsCounterDistance
    XLogical = pViewport.ConvertCounterDistanceToLogicalY(Me.X) + pViewport.Boundary.Left
End Select

If Not Me.Offset Is Nothing Then XLogical = XLogical + Me.Offset.WidthLogical(pViewport)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Property Get yLogical( _
                ByVal pViewport As Viewport) As Double
Const ProcName As String = "YLogical"
Dim failpoint As String
On Error GoTo Err

Select Case Me.CoordinateSystemY
Case CoordsLogical
    yLogical = Me.Y
Case CoordsRelative
    yLogical = pViewport.ConvertRelativeToLogicalY(Me.Y) + pViewport.Boundary.Bottom
Case CoordsDistance
    yLogical = pViewport.ConvertDistanceToLogicalY(Me.Y) + pViewport.Boundary.Bottom
Case CoordsCounterDistance
    yLogical = pViewport.ConvertCounterDistanceToLogicalY(Me.Y) + pViewport.Boundary.Bottom
End Select

If Not Me.Offset Is Nothing Then yLogical = yLogical + Me.Offset.HeightLogical(pViewport)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'================================================================================
' Helper Functions
'================================================================================


