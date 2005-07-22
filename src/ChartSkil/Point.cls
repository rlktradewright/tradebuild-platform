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

Public Enum PositionType
    PositionAbsolute = 0
    PositionRelative
    PositionDistance        ' Measured from left or bottom of region
    PositionCounterDistance ' Measured from right or top of region
End Enum
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
Private mPositionTypeX As PositionType
Private mPosnTypeXAssigned As Boolean
Private mPositionTypeY As PositionType
Private mPosnTypeYAssigned As Boolean
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

Friend Property Get isPosnTypeAssignedX() As Boolean
isPosnTypeAssignedX = mPosnTypeXAssigned
End Property

Friend Property Get isPosnTypeAssignedY() As Boolean
isPosnTypeAssignedY = mPosnTypeYAssigned
End Property

Public Property Get PositionTypeX() As PositionType
PositionTypeX = mPositionTypeX
End Property

Friend Property Let PositionTypeX(value As PositionType)
If Not mPosnTypeXAssigned Then
    mPositionTypeX = value
    mPosnTypeXAssigned = True
    Exit Property
End If

If mPositionTypeX = value Then Exit Property

Select Case mPositionTypeX
Case PositionAbsolute
    If value = PositionCounterDistance Then
        mX = (mSurface.scaleLeft + mSurface.scaleWidth - mX) / mSurface.scaleWidth * mSurface.Width / TwipsPerCm
    ElseIf value = PositionDistance Then
        mX = (mX - mSurface.scaleLeft) / mSurface.scaleWidth * mSurface.Width / TwipsPerCm
    ElseIf value = PositionRelative Then
        mX = (mX - mSurface.scaleLeft) / mSurface.scaleWidth * 100
    End If
Case PositionRelative
    If value = PositionAbsolute Then
        mX = XAbsolute
    ElseIf value = PositionCounterDistance Then
        mX = ((100 - mX) * mSurface.Width) / 100 / TwipsPerCm
    ElseIf value = PositionDistance Then
        mX = (mX * mSurface.Width) / 100 / TwipsPerCm
    End If
Case PositionDistance
    If value = PositionAbsolute Then
        mX = XAbsolute
    ElseIf value = PositionCounterDistance Then
        mX = ((100 - (100 * mX * TwipsPerCm / mSurface.Width)) * mSurface.Width) / 100 / TwipsPerCm
    ElseIf value = PositionRelative Then
        mX = 100 * mX * TwipsPerCm / mSurface.Width
    End If
Case PositionCounterDistance
    If value = PositionAbsolute Then
        mX = XAbsolute
    ElseIf value = PositionDistance Then
        mX = ((100 - 100 * mX * TwipsPerCm / mSurface.Width) * mSurface.Width) / 100 / TwipsPerCm
    ElseIf value = PositionRelative Then
        mX = 100 - 100 * mX * TwipsPerCm / mSurface.Width
    End If
End Select
mPositionTypeX = value
End Property

Public Property Get PositionTypeY() As PositionType
PositionTypeY = mPositionTypeY
End Property

Friend Property Let PositionTypeY(value As PositionType)
If Not mPosnTypeYAssigned Then
    mPositionTypeY = value
    mPosnTypeYAssigned = True
    Exit Property
End If

If mPositionTypeY = value Then Exit Property

Select Case mPositionTypeY
Case PositionAbsolute
    If value = PositionCounterDistance Then
        mY = -(mSurface.scaleTop - mY) / mSurface.scaleHeight * mSurface.Height / TwipsPerCm
    ElseIf value = PositionDistance Then
        mY = -(mY - (mSurface.scaleTop + mSurface.scaleHeight)) / mSurface.scaleHeight * mSurface.Height / TwipsPerCm
    ElseIf value = PositionRelative Then
        mY = -(mY - (mSurface.scaleTop + mSurface.scaleHeight)) / mSurface.scaleHeight * 100
    End If
Case PositionRelative
    If value = PositionAbsolute Then
        mY = YAbsolute
    ElseIf value = PositionCounterDistance Then
        mY = -((100 - mY) * mSurface.Height) / 100 / TwipsPerCm
    ElseIf value = PositionDistance Then
        mY = -(mY * mSurface.Height) * 100 / TwipsPerCm
    End If
Case PositionDistance
    If value = PositionAbsolute Then
        mY = YAbsolute
    ElseIf value = PositionCounterDistance Then
        mY = -((100 - (-mY * TwipsPerCm / mSurface.Height / 100)) * mSurface.Height) / 100 / TwipsPerCm
    ElseIf value = PositionRelative Then
        mY = -mY * TwipsPerCm / mSurface.Height / 100
    End If
Case PositionCounterDistance
    If value = PositionAbsolute Then
        mY = YAbsolute
    ElseIf value = PositionDistance Then
        mY = -((100 + 100 * mY * TwipsPerCm / mSurface.Height) * mSurface.Height) * 100 / TwipsPerCm
    ElseIf value = PositionRelative Then
        mY = 100 + 100 * mY * TwipsPerCm / mSurface.Height
    End If
End Select
mPositionTypeY = value
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
Select Case mPositionTypeX
Case PositionAbsolute
    XAbsolute = mX
Case PositionRelative
    XAbsolute = (mX * mSurface.scaleWidth / 100) + mSurface.scaleLeft
Case PositionDistance
    XAbsolute = (mX * TwipsPerCm) / mSurface.Width * mSurface.scaleWidth + mSurface.scaleLeft
Case PositionCounterDistance
    XAbsolute = mSurface.scaleLeft + mSurface.scaleWidth - (mX * TwipsPerCm) / mSurface.Width * mSurface.scaleWidth
End Select
End Property

Public Property Get Y() As Double
Y = mY
End Property

Friend Property Let Y(ByVal value As Double)
mY = value
mYAssigned = True
End Property

Friend Property Get YAbsolute() As Double
Select Case mPositionTypeY
Case PositionAbsolute
    YAbsolute = mY
Case PositionRelative
    YAbsolute = -(mY * mSurface.scaleHeight / 100) + mSurface.scaleTop + mSurface.scaleHeight
Case PositionDistance
    YAbsolute = -(mY * TwipsPerCm) / mSurface.Height * mSurface.scaleHeight + mSurface.scaleTop + mSurface.scaleHeight
Case PositionCounterDistance
    YAbsolute = mSurface.scaleTop + (mY * TwipsPerCm) / mSurface.Height * mSurface.scaleHeight
End Select
End Property

'================================================================================
' Methods
'================================================================================

Public Function clone() As Point
Set clone = New Point
clone.surface = mSurface
If mXAssigned Then clone.X = mX
If mYAssigned Then clone.Y = mY
If mPosnTypeXAssigned Then clone.PositionTypeX = mPositionTypeX
If mPosnTypeYAssigned Then clone.PositionTypeY = mPositionTypeY
End Function

Public Function equals(ByVal aPoint As Point) As Boolean
If aPoint Is Nothing Then Exit Function
If aPoint Is Me Then Exit Function
With aPoint
    If Not .surface Is mSurface Then Exit Function
    If .isPosnTypeAssignedX <> mPosnTypeXAssigned Then Exit Function
    If .isPosnTypeAssignedY <> mPosnTypeYAssigned Then Exit Function
    If .PositionTypeX <> mPositionTypeX Then Exit Function
    If .PositionTypeY <> mPositionTypeY Then Exit Function
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


