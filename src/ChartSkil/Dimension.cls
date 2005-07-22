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
Private mSurface As PictureBox

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get isAssigned() As Boolean
isAssigned = (mXAssigned And mYAssigned)
End Property

Friend Property Get isAssignedX() As Boolean
isAssignedX = mXAssigned
End Property

Friend Property Get isAssignedY() As Boolean
isAssignedY = mYAssigned
End Property

Friend Property Get surface() As PictureBox
Set surface = mSurface
End Property

Friend Property Let surface(ByVal value As PictureBox)
Set mSurface = value
End Property

Public Property Get X() As Double
X = mX
End Property

Friend Property Let X(ByVal value As Double)
mX = value
mXAssigned = True
End Property

Friend Property Get XAbsolute() As Double
XAbsolute = (mX * TwipsPerCm) / mSurface.Width * mSurface.scaleWidth
End Property

Public Property Get Y() As Double
Y = mY
End Property

Friend Property Let Y(ByVal value As Double)
mY = value
mYAssigned = True
End Property

Friend Property Get YAbsolute() As Double
YAbsolute = -(mY * TwipsPerCm) / mSurface.Height * mSurface.scaleHeight
End Property

'================================================================================
' Methods
'================================================================================

Public Function clone() As Point
Set clone = New Point
clone.surface = mSurface
If mXAssigned Then clone.X = mX
If mYAssigned Then clone.Y = mY
End Function

Public Function equals(ByVal aPoint As Point) As Boolean
If aPoint Is Nothing Then Exit Function
If aPoint Is Me Then Exit Function
With aPoint
    If Not .surface Is mSurface Then Exit Function
    If .isAssignedX <> mXAssigned Then Exit Function
    If .isAssignedY <> mYAssigned Then Exit Function
    If .X <> mX Then Exit Function
    If .Y <> mY Then Exit Function
End With
equals = True
End Function

'================================================================================
' Helper Functions
'================================================================================



