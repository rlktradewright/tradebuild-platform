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
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private mX As Double
Private mXAssigned As Boolean
Private mY As Double
Private mYAssigned As Boolean
Private mRelative As Boolean

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get assigned() As Boolean
assigned = (mXAssigned And mYAssigned)
End Property

Public Property Get x() As Double
x = mX
End Property

Friend Property Let x(ByVal value As Double)
mX = value
mXAssigned = True
End Property

Friend Property Get Xassigned() As Boolean
Xassigned = mXAssigned
End Property

Public Property Get y() As Double
y = mY
End Property

Friend Property Let y(ByVal value As Double)
mY = value
mYAssigned = True
End Property

Friend Property Get Yassigned() As Boolean
Yassigned = mYAssigned
End Property

Public Property Get relative() As Boolean
relative = mRelative
End Property

Friend Property Let relative(ByVal value As Boolean)
mRelative = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function clone() As Point
Set clone = New Point
If mXAssigned Then clone.x = mX
If mYAssigned Then clone.y = mY
clone.relative = mRelative
End Function

Public Function equals(ByVal aPoint As Point) As Boolean
With aPoint
    If .relative <> mRelative Then Exit Function
    If .Xassigned <> mXAssigned Then Exit Function
    If .Yassigned <> Yassigned Then Exit Function
    If .x <> mX Then Exit Function
    If .y <> mY Then Exit Function
End With
equals = True
End Function

'================================================================================
' Helper Functions
'================================================================================


