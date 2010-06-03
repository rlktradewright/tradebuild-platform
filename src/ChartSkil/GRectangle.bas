Attribute VB_Name = "GRectangle"
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "GRectangle"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'================================================================================
' Rectangle functions
'
' NB: these are implemented as functions rather than class methods for
' efficiency reasons, due to the very large numbers of rectangles made
' use of
'================================================================================

Public Function IntContains( _
                ByRef pInt As TInterval, _
                ByVal X As Double) As Boolean
If Not pInt.isValid Then Exit Function
IntContains = (X >= pInt.startValue) And (X <= pInt.endValue)
End Function

Public Function IntIntersection( _
                ByRef int1 As TInterval, _
                ByRef int2 As TInterval) As TInterval
Dim startValue1 As Double
Dim endValue1 As Double
Dim startValue2 As Double
Dim endValue2 As Double

If Not int1.isValid Or Not int2.isValid Then Exit Function

startValue1 = int1.startValue
endValue1 = int1.endValue
startValue2 = int2.startValue
endValue2 = int2.endValue

With IntIntersection
    If startValue1 >= startValue2 And startValue1 <= endValue2 Then
        .startValue = startValue1
        If endValue1 >= startValue2 And endValue1 <= endValue2 Then
            .endValue = endValue1
        Else
            .endValue = endValue2
        End If
        .isValid = True
        Exit Function
    End If
    If endValue1 >= startValue2 And endValue1 <= endValue2 Then
        .endValue = endValue1
        .startValue = startValue2
        .isValid = True
        Exit Function
    End If
    If startValue1 < startValue2 And endValue1 > endValue2 Then
        .startValue = startValue2
        .endValue = endValue2
        .isValid = True
        Exit Function
    End If
End With
End Function

Public Function IntOverlaps( _
                ByRef int1 As TInterval, _
                ByRef int2 As TInterval) As Boolean
                        
IntOverlaps = True

If int1.startValue >= int2.startValue And int1.startValue <= int2.endValue Then
    Exit Function
End If
If int1.endValue >= int2.startValue And int1.endValue <= int2.endValue Then
    Exit Function
End If
If int1.startValue < int2.startValue And int1.endValue > int2.endValue Then
    Exit Function
End If
IntOverlaps = False
End Function

Public Function PointAdd( _
                ByRef pPoint1 As TPoint, _
                ByRef pPoint2 As TPoint) As TPoint
PointAdd.X = pPoint1.X + pPoint2.X
PointAdd.Y = pPoint1.Y + pPoint2.Y
End Function

Public Function PointSubtract( _
                ByRef pPoint1 As TPoint, _
                ByRef pPoint2 As TPoint) As TPoint
PointSubtract.X = pPoint1.X - pPoint2.X
PointSubtract.Y = pPoint1.Y - pPoint2.Y
End Function

Public Function PointToString( _
                ByRef pPoint As TPoint) As String
PointToString = "X=" & pPoint.X & "; Y=" & pPoint.Y
End Function

Public Function RectBottomCentre( _
                ByRef pRect As TRectangle) As TPoint
RectBottomCentre.X = (pRect.Right + pRect.Left) / 2
RectBottomCentre.Y = pRect.Bottom
End Function

Public Function RectBottomLeft( _
                ByRef pRect As TRectangle) As TPoint
RectBottomLeft.X = pRect.Left
RectBottomLeft.Y = pRect.Bottom
End Function

Public Function RectBottomRight( _
                ByRef pRect As TRectangle) As TPoint
RectBottomRight.X = pRect.Right
RectBottomRight.Y = pRect.Bottom
End Function

Public Function RectCentreCentre( _
                ByRef pRect As TRectangle) As TPoint
RectCentreCentre.X = (pRect.Right + pRect.Left) / 2
RectCentreCentre.Y = (pRect.Top + pRect.Bottom) / 2
End Function

Public Function RectCentreLeft( _
                ByRef pRect As TRectangle) As TPoint
RectCentreLeft.X = pRect.Left
RectCentreLeft.Y = (pRect.Top + pRect.Bottom) / 2
End Function

Public Function RectCentreRight( _
                ByRef pRect As TRectangle) As TPoint
RectCentreRight.X = pRect.Right
RectCentreRight.Y = (pRect.Top + pRect.Bottom) / 2
End Function

Public Function RectContainsPoint( _
                ByRef pRect As TRectangle, _
                ByVal X As Double, _
                ByVal Y As Double) As Boolean
If Not pRect.isValid Then Exit Function
If X < pRect.Left Then Exit Function
If X > pRect.Right Then Exit Function
If Y < pRect.Bottom Then Exit Function
If Y > pRect.Top Then Exit Function
RectContainsPoint = True
End Function

Public Function RectContainsRect( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As Boolean
If Not rect1.isValid Then Exit Function
If Not rect2.isValid Then Exit Function
If rect2.Left < rect1.Left Then Exit Function
If rect2.Right > rect1.Right Then Exit Function
If rect2.Bottom < rect1.Bottom Then Exit Function
If rect2.Top > rect1.Top Then Exit Function
RectContainsRect = True
End Function

Public Function RectEquals( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As Boolean
With rect1
    If Not .isValid Or Not rect2.isValid Then Exit Function
    If .Bottom <> rect2.Bottom Then Exit Function
    If .Left <> rect2.Left Then Exit Function
    If .Top <> rect2.Top Then Exit Function
    If .Right <> rect2.Right Then Exit Function
End With
RectEquals = True
End Function

Public Sub RectExpand( _
                ByRef pRect As TRectangle, _
                ByVal xIncrement As Double, _
                ByVal yIncrement As Double)
With pRect
    If Not .isValid Then Exit Sub
    .Left = .Left - xIncrement
    .Right = .Right + xIncrement
    .Top = .Top + yIncrement
    .Bottom = .Bottom - yIncrement
End With
End Sub

Public Sub RectExpandBySize( _
                ByRef pRect As TRectangle, _
                ByVal pSize As Size, _
                ByVal pViewport As Viewport)
RectExpand pRect, pSize.WidthLogical(pViewport), pSize.HeightLogical(pViewport)
End Sub

Public Function RectGetXInterval( _
                ByRef pRect As TRectangle) As TInterval
With RectGetXInterval
    .startValue = pRect.Left
    .endValue = pRect.Right
    .isValid = pRect.isValid
End With
End Function

Public Function RectGetYInterval( _
                ByRef pRect As TRectangle) As TInterval
With RectGetYInterval
    .startValue = pRect.Bottom
    .endValue = pRect.Top
    .isValid = pRect.isValid
End With
End Function

Public Sub RectInitialise( _
                ByRef pRect As TRectangle)
With pRect
    .isValid = False
    .Left = PlusInfinityDouble
    .Right = MinusInfinityDouble
    .Bottom = PlusInfinityDouble
    .Top = MinusInfinityDouble
End With
End Sub

Public Function RectIntersection( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TRectangle
Dim xInt As TInterval
Dim yint As TInterval
xInt = IntIntersection(RectGetXInterval(rect1), RectGetXInterval(rect2))
yint = IntIntersection(RectGetYInterval(rect1), RectGetYInterval(rect2))
With RectIntersection
    .Left = xInt.startValue
    .Right = xInt.endValue
    .Bottom = yint.startValue
    .Top = yint.endValue
    If xInt.isValid And yint.isValid Then .isValid = True
End With
End Function


Public Function RectOverlaps( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As Boolean
RectOverlaps = IntOverlaps(RectGetXInterval(rect1), RectGetXInterval(rect2)) And _
            IntOverlaps(RectGetYInterval(rect1), RectGetYInterval(rect2))
            
End Function

Public Sub RectOffset( _
                ByRef pRect As TRectangle, _
                ByVal pdX As Double, _
                ByVal pdY As Double)

If Not pRect.isValid Then Exit Sub

With pRect
    .Left = .Left + pdX
    .Right = .Right + pdX
    .Bottom = .Bottom + pdY
    .Top = .Top + pdY
End With
End Sub

Public Sub RectOffsetBySize( _
                ByRef pRect As TRectangle, _
                ByRef pOffset As Size, _
                ByVal pViewport As Viewport)
RectOffset pRect, pOffset.WidthLogical(pViewport), pOffset.HeightLogical(pViewport)
End Sub

Public Sub RectOffsetPoint( _
                ByRef pRect As TRectangle, _
                ByRef pOffset As TPoint)
RectOffset pRect, pOffset.X, pOffset.Y
End Sub

Public Sub RectSetXInterval( _
                ByRef pRect As TRectangle, _
                ByRef interval As TInterval)
With pRect
    If interval.startValue <= interval.endValue Then
        .Left = interval.startValue
        .Right = interval.endValue
    Else
        .Left = interval.endValue
        .Right = interval.startValue
    End If
    .isValid = .isValid And interval.isValid
End With
End Sub

Public Sub RectSetYInterval( _
                ByRef pRect As TRectangle, _
                ByRef interval As TInterval)
With pRect
    If interval.startValue <= interval.endValue Then
        .Bottom = interval.startValue
        .Top = interval.endValue
    Else
        .Bottom = interval.endValue
        .Top = interval.startValue
    End If
    .isValid = interval.isValid
End With
End Sub

Public Function RectTopCentre( _
                ByRef pRect As TRectangle) As TPoint
RectTopCentre.X = (pRect.Right + pRect.Left) / 2
RectTopCentre.Y = pRect.Top
End Function

Public Function RectTopLeft( _
                ByRef pRect As TRectangle) As TPoint
RectTopLeft.X = pRect.Left
RectTopLeft.Y = pRect.Top
End Function

Public Function RectTopRight( _
                ByRef pRect As TRectangle) As TPoint
RectTopRight.X = pRect.Right
RectTopRight.Y = pRect.Top
End Function

Public Function RectToString( _
                ByRef pRect As TRectangle) As String
RectToString = IIf(pRect.isValid, "Valid: ", "Invalid: ") & "Bottom=" & pRect.Bottom & "; Left=" & pRect.Left & "; Top=" & pRect.Top & "; Right=" & pRect.Right
End Function

Public Sub RectTranslate( _
                ByRef pRect As TRectangle, _
                ByRef pDisplacement As TPoint)
pRect.Bottom = pRect.Bottom + pDisplacement.Y
pRect.Left = pRect.Left + pDisplacement.X
pRect.Top = pRect.Top + pDisplacement.Y
pRect.Right = pRect.Right + pDisplacement.X
End Sub

Public Function RectUnion( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TRectangle
If Not (rect1.isValid And rect2.isValid) Then
    If rect1.isValid Then
        RectUnion = rect1
    ElseIf rect2.isValid Then
        RectUnion = rect2
    End If
    Exit Function
End If

With RectUnion
    .isValid = False
    
    If rect1.Left < rect2.Left Then
        .Left = rect1.Left
    Else
        .Left = rect2.Left
    End If
    If rect1.Bottom < rect2.Bottom Then
        .Bottom = rect1.Bottom
    Else
        .Bottom = rect2.Bottom
    End If
    If rect1.Top > rect2.Top Then
        .Top = rect1.Top
    Else
        .Top = rect2.Top
    End If
    If rect1.Right > rect2.Right Then
        .Right = rect1.Right
    Else
        .Right = rect2.Right
    End If
    .isValid = True
End With
End Function

Public Sub RectValidate( _
                ByRef pRect As TRectangle, _
                Optional allowZeroDimensions As Boolean = False)
With pRect
    If allowZeroDimensions Then
        If .Left <= .Right And .Bottom <= .Top Then .isValid = True
    Else
        If .Left < .Right And .Bottom < .Top Then .isValid = True
    End If
End With
End Sub

Public Function RectXIntersection( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TInterval
RectXIntersection = IntIntersection(RectGetXInterval(rect1), RectGetXInterval(rect2))
End Function

Public Function RectYIntersection( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TInterval
RectYIntersection = IntIntersection(RectGetYInterval(rect1), RectGetYInterval(rect2))
End Function

Public Sub TPointMultiply( _
                ByRef pPoint As TPoint, _
                ByVal pFactor As Double)
pPoint.X = pPoint.X * pFactor
pPoint.Y = pPoint.Y * pFactor
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

