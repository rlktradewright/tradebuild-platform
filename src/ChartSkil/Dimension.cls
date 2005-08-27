VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dimension"
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

Public Property Get x() As Double
x = mX
End Property

Friend Property Let x(ByVal value As Double)
mX = value
mXAssigned = True
End Property

Friend Property Get XLogical() As Double
XLogical = mCanvas.convertDistanceToLogicalX(mX)
End Property

Public Property Get y() As Double
y = mY
End Property

Friend Property Let y(ByVal value As Double)
mY = value
mYAssigned = True
End Property

Friend Property Get YLogical() As Double
YLogical = mCanvas.convertDistanceToLogicalY(mY)
End Property

'================================================================================
' Methods
'================================================================================

Public Function equals(ByVal aPoint As Point) As Boolean
If aPoint Is Nothing Then Exit Function
If aPoint Is Me Then Exit Function
With aPoint
    If Not .canvas Is mCanvas Then Exit Function
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



