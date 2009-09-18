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
Private mViewport As Viewport

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

Friend Property Let Viewport(ByVal value As Viewport)
Set mViewport = value
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

Public Property Get X() As Double
X = mX
End Property

Friend Property Let X(ByVal value As Double)
mX = value
mXAssigned = True
End Property

Public Property Get XLogical() As Double
XLogical = mViewport.ConvertDistanceToLogicalX(mX) - mViewport.Boundary.Left
End Property

Public Property Get Y() As Double
Y = mY
End Property

Friend Property Let Y(ByVal value As Double)
mY = value
mYAssigned = True
End Property

Public Property Get YLogical() As Double
YLogical = mViewport.ConvertDistanceToLogicalY(mY) - mViewport.Boundary.Bottom
End Property

'================================================================================
' Methods
'================================================================================

Public Function Equals(ByVal aDim As Dimension) As Boolean
If aDim Is Nothing Then Exit Function
If aDim Is Me Then
    Equals = True
    Exit Function
End If
With aDim
    If Not .Viewport Is mViewport Then Exit Function
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



