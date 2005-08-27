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
Private mCanvas As canvas

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Get canvas() As canvas
Set canvas = mCanvas
End Property

Friend Property Let canvas(ByVal value As canvas)
Set mCanvas = value
End Property

Public Property Get isAssigned() As Boolean
isAssigned = (mXAssigned And mYAssigned)
End Property

Friend Property Get isAssignedX() As Boolean
isAssignedX = mXAssigned
End Property

Friend Property Get isAssignedY() As Boolean
isAssignedY = mYAssigned
End Property

Friend Property Get isCoordinateSystemAssignedX() As Boolean
isCoordinateSystemAssignedX = mCoordSystemXAssigned
End Property

Friend Property Get isCoordinateSystemAssignedY() As Boolean
isCoordinateSystemAssignedY = mCoordSystemYAssigned
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
        mX = mCanvas.convertLogicalToCounterDistanceX(mX)
    ElseIf value = CoordsDistance Then
        mX = mCanvas.convertLogicalToDistanceX(mX)
    ElseIf value = CoordsRelative Then
        mX = mCanvas.convertLogicalToRelativeX(mX)
    End If
Case CoordsRelative
    If value = CoordsLogical Then
        mX = XLogical
    ElseIf value = CoordsCounterDistance Then
        mX = mCanvas.convertRelativeToCounterDistanceX(mX)
    ElseIf value = CoordsDistance Then
        mX = mCanvas.convertRelativeToDistanceX(mX)
    End If
Case CoordsDistance
    If value = CoordsLogical Then
        mX = XLogical
    ElseIf value = CoordsCounterDistance Then
        mX = mCanvas.convertDistanceToCounterDistanceX(mX)
    ElseIf value = CoordsRelative Then
        mX = mCanvas.convertDistanceToRelativeX(mX)
    End If
Case CoordsCounterDistance
    If value = CoordsLogical Then
        mX = XLogical
    ElseIf value = CoordsDistance Then
        mX = mCanvas.convertCounterDistanceToDistanceX(mX)
    ElseIf value = CoordsRelative Then
        mX = mCanvas.convertCounterDistanceToRelativeX(mX)
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
        mY = mCanvas.convertLogicalToCounterDistanceY(mY)
    ElseIf value = CoordsDistance Then
        mY = mCanvas.convertLogicalToDistanceY(mY)
    ElseIf value = CoordsRelative Then
        mY = mCanvas.convertLogicalToRelativeY(mY)
    End If
Case CoordsRelative
    If value = CoordsLogical Then
        mY = YLogical
    ElseIf value = CoordsCounterDistance Then
        mY = mCanvas.convertRelativeToCounterDistanceY(mY)
    ElseIf value = CoordsDistance Then
        mY = mCanvas.convertRelativeToDistanceY(mY)
    End If
Case CoordsDistance
    If value = CoordsLogical Then
        mY = YLogical
    ElseIf value = CoordsCounterDistance Then
        mY = mCanvas.convertDistanceToCounterDistanceY(mY)
    ElseIf value = CoordsRelative Then
        mY = mCanvas.convertDistanceToRelativeY(mY)
    End If
Case CoordsCounterDistance
    If value = CoordsLogical Then
        mY = YLogical
    ElseIf value = CoordsDistance Then
        mY = mCanvas.convertCounterDistanceToDistanceY(mY)
    ElseIf value = CoordsRelative Then
        mY = mCanvas.convertCounterDistanceToRelativeY(mY)
    End If
End Select
mCoordSystemY = value
End Property

Public Property Get x() As Double
x = mX
End Property

Friend Property Let x(ByVal value As Double)
mX = value
mXAssigned = True
End Property

Friend Property Get XLogical() As Double
Select Case mCoordSystemX
Case CoordsLogical
    XLogical = mX
Case CoordsRelative
    XLogical = mCanvas.convertRelativeToLogicalX(mX) + mCanvas.left
Case CoordsDistance
    XLogical = mCanvas.convertDistanceToLogicalX(mX) + mCanvas.left
Case CoordsCounterDistance
    XLogical = mCanvas.convertCounterDistanceToLogicalX(mX) + mCanvas.left
End Select
End Property

Public Property Get y() As Double
y = mY
End Property

Friend Property Let y(ByVal value As Double)
mY = value
mYAssigned = True
End Property

Friend Property Get YLogical() As Double
Select Case mCoordSystemY
Case CoordsLogical
    YLogical = mY
Case CoordsRelative
    YLogical = mCanvas.convertRelativeToLogicalY(mY) + mCanvas.bottom
Case CoordsDistance
    YLogical = mCanvas.convertDistanceToLogicalY(mY) + mCanvas.bottom
Case CoordsCounterDistance
    YLogical = mCanvas.convertCounterDistanceToLogicalY(mY) + mCanvas.bottom
End Select
End Property

'================================================================================
' Methods
'================================================================================

Public Function equals(ByVal aPoint As Point) As Boolean
If aPoint Is Nothing Then Exit Function
If aPoint Is Me Then Exit Function
With aPoint
    If Not .canvas Is mCanvas Then Exit Function
    If .isCoordinateSystemAssignedX <> mCoordSystemXAssigned Then Exit Function
    If .isCoordinateSystemAssignedY <> mCoordSystemYAssigned Then Exit Function
    If .CoordinateSystemX <> mCoordSystemX Then Exit Function
    If .CoordinateSystemY <> mCoordSystemY Then Exit Function
    If .isAssignedX <> mXAssigned Then Exit Function
    If .isAssignedY <> mYAssigned Then Exit Function
    If .x <> mX Then Exit Function
    If .y <> mY Then Exit Function
End With
equals = True
End Function

'================================================================================
' Helper Functions
'================================================================================


