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

Private Const ProjectName                   As String = "ChartSkil26"
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
                ByVal x As Double) As Boolean
If Not pInt.isValid Then Exit Function
IntContains = (x >= pInt.startValue) And (x <= pInt.endValue)
End Function

Public Function intIntersection( _
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

With intIntersection
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

Public Function intOverlaps( _
                ByRef int1 As TInterval, _
                ByRef int2 As TInterval) As Boolean
                        
intOverlaps = True

If int1.startValue >= int2.startValue And int1.startValue <= int2.endValue Then
    Exit Function
End If
If int1.endValue >= int2.startValue And int1.endValue <= int2.endValue Then
    Exit Function
End If
If int1.startValue < int2.startValue And int1.endValue > int2.endValue Then
    Exit Function
End If
intOverlaps = False
End Function

Public Function RectContainsPoint( _
                ByRef rect As TRectangle, _
                ByVal x As Double, _
                ByVal y As Double) As Boolean
If Not rect.isValid Then Exit Function
If x < rect.Left Then Exit Function
If x > rect.right Then Exit Function
If y < rect.Bottom Then Exit Function
If y > rect.Top Then Exit Function
RectContainsPoint = True
End Function

Public Function RectContainsRect( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As Boolean
If Not rect1.isValid Then Exit Function
If Not rect2.isValid Then Exit Function
If rect2.Left < rect1.Left Then Exit Function
If rect2.right > rect1.right Then Exit Function
If rect2.Bottom < rect1.Bottom Then Exit Function
If rect2.Top > rect1.Top Then Exit Function
RectContainsRect = True
End Function

Public Function rectEquals( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As Boolean
With rect1
    If Not .isValid Or Not rect2.isValid Then Exit Function
    If .Bottom <> rect2.Bottom Then Exit Function
    If .Left <> rect2.Left Then Exit Function
    If .Top <> rect2.Top Then Exit Function
    If .right <> rect2.right Then Exit Function
End With
rectEquals = True
End Function

Public Sub rectInitialise( _
                ByRef rect As TRectangle)
With rect
    .isValid = False
    .Left = PlusInfinityDouble
    .right = MinusInfinityDouble
    .Bottom = PlusInfinityDouble
    .Top = MinusInfinityDouble
End With
End Sub

Public Function rectIntersection( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TRectangle
Dim xInt As TInterval
Dim yint As TInterval
xInt = intIntersection(rectGetXInterval(rect1), rectGetXInterval(rect2))
yint = intIntersection(rectGetYInterval(rect1), rectGetYInterval(rect2))
With rectIntersection
    .Left = xInt.startValue
    .right = xInt.endValue
    .Bottom = yint.startValue
    .Top = yint.endValue
    If xInt.isValid And yint.isValid Then .isValid = True
End With
End Function


Public Function rectOverlaps( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As Boolean
rectOverlaps = intOverlaps(rectGetXInterval(rect1), rectGetXInterval(rect2)) And _
            intOverlaps(rectGetYInterval(rect1), rectGetYInterval(rect2))
            
End Function

Public Function rectXIntersection( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TInterval
rectXIntersection = intIntersection(rectGetXInterval(rect1), rectGetXInterval(rect2))
End Function

Public Function rectYIntersection( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TInterval
rectYIntersection = intIntersection(rectGetYInterval(rect1), rectGetYInterval(rect2))
End Function

Public Function rectGetXInterval( _
                ByRef rect As TRectangle) As TInterval
With rectGetXInterval
.startValue = rect.Left
.endValue = rect.right
.isValid = rect.isValid
End With
End Function

Public Function rectGetYInterval( _
                ByRef rect As TRectangle) As TInterval
With rectGetYInterval
    .startValue = rect.Bottom
    .endValue = rect.Top
    .isValid = rect.isValid
End With
End Function

Public Sub rectSetXInterval( _
                ByRef rect As TRectangle, _
                ByRef interval As TInterval)
With rect
    .Left = interval.startValue
    .right = interval.endValue
    .isValid = interval.isValid
End With
End Sub

Public Sub rectSetYInterval( _
                ByRef rect As TRectangle, _
                ByRef interval As TInterval)
With rect
    .Bottom = interval.startValue
    .Top = interval.endValue
    .isValid = interval.isValid
End With
End Sub

Public Function rectUnion( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TRectangle
If Not (rect1.isValid And rect2.isValid) Then
    If rect1.isValid Then
        rectUnion = rect1
    ElseIf rect2.isValid Then
        rectUnion = rect2
    End If
    Exit Function
End If

With rectUnion
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
    If rect1.right > rect2.right Then
        .right = rect1.right
    Else
        .right = rect2.right
    End If
    .isValid = True
End With
End Function

Public Sub rectValidate( _
                ByRef rect As TRectangle)
With rect
    If .Left <= .right And .Bottom <= .Top Then
        .isValid = True
    Else
        .isValid = False
    End If
End With
End Sub


'@================================================================================
' Helper Functions
'@================================================================================
