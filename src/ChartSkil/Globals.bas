Attribute VB_Name = "Globals"
Option Explicit

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, _
                                            lpPoint As POINTAPI, _
                                            ByVal nCount As Long) As Long

Public Type TInterval
    isvalid         As Boolean
    startValue      As Double
    endValue        As Double
End Type

Public Const Pi As Double = 3.14159265358979

Public Const MinusInfinityDouble As Double = -1.79769313486231E+308
Public Const PlusInfinityDouble As Double = 1.79769313486231E+308

Public Const MinusInfinitySingle As Single = -3.402823E+38
Public Const PlusInfinitySingle As Single = 3.402823E+38

Public Const GridlineSpacingCm As Double = 2.5

Public gChartBackColour As Long

Public Function newPoint(ByVal x As Double, _
                        ByVal y As Double, _
                        Optional ByVal relative As Boolean = False) As Point
Set newPoint = New Point
newPoint.x = x
newPoint.y = y
newPoint.relative = relative
End Function


'================================================================================
' Rectangle functions
'
' NB: these are implemented as functions rather than class methods for
' efficiency reasons, due to the very large numbers of rectangles made
' use of
'================================================================================

Public Function intIntersection( _
                        ByRef int1 As TInterval, _
                        ByRef int2 As TInterval) As TInterval
Dim startValue1 As Double
Dim endValue1 As Double
Dim startValue2 As Double
Dim endValue2 As Double

If Not int1.isvalid Or Not int2.isvalid Then Exit Function

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
        .isvalid = True
        Exit Function
    End If
    If endValue1 >= startValue2 And endValue1 <= endValue2 Then
        .endValue = endValue1
        .startValue = startValue2
        .isvalid = True
        Exit Function
    End If
    If startValue1 < startValue2 And endValue1 > endValue2 Then
        .startValue = startValue2
        .endValue = endValue2
        .isvalid = True
        Exit Function
    End If
End With
End Function

Public Function intOverlaps(ByRef int1 As TInterval, _
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

Public Sub rectInitialise(ByRef rect As TRectangle)
With rect
    .isvalid = False
    .left = MinusInfinityDouble
    .right = PlusInfinityDouble
    .bottom = MinusInfinityDouble
    .top = PlusInfinityDouble
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
    .left = xInt.startValue
    .right = xInt.endValue
    .bottom = yint.startValue
    .top = yint.endValue
    If xInt.isvalid And yint.isvalid Then .isvalid = True
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

Public Function rectGetXInterval(ByRef rect As TRectangle) As TInterval
With rectGetXInterval
.startValue = rect.left
.endValue = rect.right
.isvalid = rect.isvalid
End With
End Function

Public Function rectGetYInterval(ByRef rect As TRectangle) As TInterval
With rectGetYInterval
    .startValue = rect.bottom
    .endValue = rect.top
    .isvalid = rect.isvalid
End With
End Function

Public Sub rectSetXInterval(ByRef rect As TRectangle, ByRef interval As TInterval)
With rect
    .left = interval.startValue
    .right = interval.endValue
    .isvalid = interval.isvalid
End With
End Sub

Public Sub rectSetYInterval(ByRef rect As TRectangle, ByRef interval As TInterval)
With rect
    .bottom = interval.startValue
    .top = interval.endValue
    .isvalid = interval.isvalid
End With
End Sub

Public Function rectUnion(ByRef rect1 As TRectangle, ByRef rect2 As TRectangle) As TRectangle
If Not (rect1.isvalid And rect2.isvalid) Then
    If rect1.isvalid Then
        rectUnion = rect1
    ElseIf rect2.isvalid Then
        rectUnion = rect2
    End If
    Exit Function
End If

With rectUnion
    .isvalid = False
    
    If rect1.left < rect2.left Then
        .left = rect1.left
    Else
        .left = rect2.left
    End If
    If rect1.bottom < rect2.bottom Then
        .bottom = rect1.bottom
    Else
        .bottom = rect2.bottom
    End If
    If rect1.top > rect2.top Then
        .top = rect1.top
    Else
        .top = rect2.top
    End If
    If rect1.right > rect2.right Then
        .right = rect1.right
    Else
        .right = rect2.right
    End If
    .isvalid = True
End With
End Function

