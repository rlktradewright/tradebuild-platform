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
Private mXAssigned As Boolean
Private mY As Double
Private mYAssigned As Boolean
Private mCoordSystemX As CoordinateSystems
Private mCoordSystemXAssigned As Boolean
Private mCoordSystemY As CoordinateSystems
Private mCoordSystemYAssigned As Boolean
Private mCanvas As Canvas

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Get Canvas() As Canvas
Set Canvas = mCanvas
End Property

Friend Property Let Canvas(ByVal value As Canvas)
Set mCanvas = value
End Property

Public Property Get IsAssigned() As Boolean
IsAssigned = (mXAssigned And mYAssigned)
End Property

Friend Property Get IsAssignedX() As Boolean
IsAssignedX = mXAssigned
End Property

Friend Property Get IsAssignedY() As Boolean
IsAssignedY = mYAssigned
End Property

Friend Property Get IsCoordinateSystemAssignedX() As Boolean
IsCoordinateSystemAssignedX = mCoordSystemXAssigned
End Property

Friend Property Get IsCoordinateSystemAssignedY() As Boolean
IsCoordinateSystemAssignedY = mCoordSystemYAssigned
End Property

Public Property Get CoordinateSystemX() As CoordinateSystems
CoordinateSystemX = mCoordSystemX
End Property

Friend Property Let CoordinateSystemX(value As CoordinateSystems)
If Not mCoordSystemXAssigned Then
    mCoordSystemX = value
    mCoordSystemXAssigned = True
    Exit Property
End If

If mCoordSystemX = value Then Exit Property

Select Case mCoordSystemX
Case CoordsLogical
    If value = CoordsCounterDistance Then
        mX = mCanvas.ConvertLogicalToCounterDistanceX(mX)
    ElseIf value = CoordsDistance Then
        mX = mCanvas.ConvertLogicalToDistanceX(mX)
    ElseIf value = CoordsRelative Then
        mX = mCanvas.ConvertLogicalToRelativeX(mX)
    End If
Case CoordsRelative
    If value = CoordsLogical Then
        mX = XLogical
    ElseIf value = CoordsCounterDistance Then
        mX = mCanvas.ConvertRelativeToCounterDistanceX(mX)
    ElseIf value = CoordsDistance Then
        mX = mCanvas.ConvertRelativeToDistanceX(mX)
    End If
Case CoordsDistance
    If value = CoordsLogical Then
        mX = XLogical
    ElseIf value = CoordsCounterDistance Then
        mX = mCanvas.ConvertDistanceToCounterDistanceX(mX)
    ElseIf value = CoordsRelative Then
        mX = mCanvas.ConvertDistanceToRelativeX(mX)
    End If
Case CoordsCounterDistance
    If value = CoordsLogical Then
        mX = XLogical
    ElseIf value = CoordsDistance Then
        mX = mCanvas.ConvertCounterDistanceToDistanceX(mX)
    ElseIf value = CoordsRelative Then
        mX = mCanvas.ConvertCounterDistanceToRelativeX(mX)
    End If
End Select
mCoordSystemX = value
End Property

Public Property Get CoordinateSystemY() As CoordinateSystems
CoordinateSystemY = mCoordSystemY
End Property

Friend Property Let CoordinateSystemY(value As CoordinateSystems)
If Not mCoordSystemYAssigned Then
    mCoordSystemY = value
    mCoordSystemYAssigned = True
    Exit Property
End If

If mCoordSystemY = value Then Exit Property

Select Case mCoordSystemY
Case CoordsLogical
    If value = CoordsCounterDistance Then
        mY = mCanvas.ConvertLogicalToCounterDistanceY(mY)
    ElseIf value = CoordsDistance Then
        mY = mCanvas.ConvertLogicalToDistanceY(mY)
    ElseIf value = CoordsRelative Then
        mY = mCanvas.ConvertLogicalToRelativeY(mY)
    End If
Case CoordsRelative
    If value = CoordsLogical Then
        mY = YLogical
    ElseIf value = CoordsCounterDistance Then
        mY = mCanvas.ConvertRelativeToCounterDistanceY(mY)
    ElseIf value = CoordsDistance Then
        mY = mCanvas.ConvertRelativeToDistanceY(mY)
    End If
Case CoordsDistance
    If value = CoordsLogical Then
        mY = YLogical
    ElseIf value = CoordsCounterDistance Then
        mY = mCanvas.ConvertDistanceToCounterDistanceY(mY)
    ElseIf value = CoordsRelative Then
        mY = mCanvas.ConvertDistanceToRelativeY(mY)
    End If
Case CoordsCounterDistance
    If value = CoordsLogical Then
        mY = YLogical
    ElseIf value = CoordsDistance Then
        mY = mCanvas.ConvertCounterDistanceToDistanceY(mY)
    ElseIf value = CoordsRelative Then
        mY = mCanvas.ConvertCounterDistanceToRelativeY(mY)
    End If
End Select
mCoordSystemY = value
End Property

Public Property Get X() As Double
X = mX
End Property

Friend Property Let X(ByVal value As Double)
mX = value
mXAssigned = True
End Property

Friend Property Get XLogical() As Double
Select Case mCoordSystemX
Case CoordsLogical
    XLogical = mX
Case CoordsRelative
    XLogical = mCanvas.ConvertRelativeToLogicalX(mX)
Case CoordsDistance
    XLogical = mCanvas.ConvertDistanceToLogicalX(mX)
Case CoordsCounterDistance
    XLogical = mCanvas.ConvertCounterDistanceToLogicalY(mX)
End Select
End Property

Public Property Get Y() As Double
Y = mY
End Property

Friend Property Let Y(ByVal value As Double)
mY = value
mYAssigned = True
End Property

Friend Property Get YLogical() As Double
Select Case mCoordSystemY
Case CoordsLogical
    YLogical = mY
Case CoordsRelative
    YLogical = mCanvas.ConvertRelativeToLogicalY(mY)
Case CoordsDistance
    YLogical = mCanvas.ConvertDistanceToLogicalY(mY)
Case CoordsCounterDistance
    YLogical = mCanvas.ConvertCounterDistanceToLogicalY(mY)
End Select
End Property

'================================================================================
' Methods
'================================================================================

Public Function Equals(ByVal aPoint As Point) As Boolean
If aPoint Is Nothing Then Exit Function
If aPoint Is Me Then Exit Function
With aPoint
    If Not .Canvas Is mCanvas Then Exit Function
    If .IsCoordinateSystemAssignedX <> mCoordSystemXAssigned Then Exit Function
    If .IsCoordinateSystemAssignedY <> mCoordSystemYAssigned Then Exit Function
    If .CoordinateSystemX <> mCoordSystemX Then Exit Function
    If .CoordinateSystemY <> mCoordSystemY Then Exit Function
    If .IsAssignedX <> mXAssigned Then Exit Function
    If .IsAssignedY <> mYAssigned Then Exit Function
    If .X <> mX Then Exit Function
    If .Y <> mY Then Exit Function
End With
Equals = True
End Function

'================================================================================
' Helper Functions
'================================================================================


