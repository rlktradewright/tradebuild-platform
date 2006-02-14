Attribute VB_Name = "Globals"
Option Explicit

Public Type Int64
    lowpart As Long
    highpart As Long
End Type

Public Type TInterval
    isValid         As Boolean
    startValue      As Double
    endValue        As Double
End Type

Public Const Pi As Double = 3.14159265358979

Public Const MinusInfinityDouble As Double = -1.79769313486231E+308
Public Const PlusInfinityDouble As Double = 1.79769313486231E+308

Public Const MinusInfinitySingle As Single = -3.402823E+38
Public Const PlusInfinitySingle As Single = 3.402823E+38

Public Const GridlineSpacingCm As Double = 2.5

Public Const TwipsPerCm As Long = 1440 / 2.54

Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Int64) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Int64) As Long

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

Public Function rectEquals( _
                        ByRef rect1 As TRectangle, _
                        ByRef rect2 As TRectangle) As Boolean
With rect1
    If Not .isValid Or Not rect2.isValid Then Exit Function
    If .bottom <> rect2.bottom Then Exit Function
    If .left <> rect2.left Then Exit Function
    If .top <> rect2.top Then Exit Function
    If .right <> rect2.right Then Exit Function
End With
rectEquals = True
End Function

Public Sub rectInitialise(ByRef rect As TRectangle)
With rect
    .isValid = False
    .left = PlusInfinityDouble
    .right = MinusInfinityDouble
    .bottom = PlusInfinityDouble
    .top = MinusInfinityDouble
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

Public Function rectGetXInterval(ByRef rect As TRectangle) As TInterval
With rectGetXInterval
.startValue = rect.left
.endValue = rect.right
.isValid = rect.isValid
End With
End Function

Public Function rectGetYInterval(ByRef rect As TRectangle) As TInterval
With rectGetYInterval
    .startValue = rect.bottom
    .endValue = rect.top
    .isValid = rect.isValid
End With
End Function

Public Sub rectSetXInterval(ByRef rect As TRectangle, ByRef interval As TInterval)
With rect
    .left = interval.startValue
    .right = interval.endValue
    .isValid = interval.isValid
End With
End Sub

Public Sub rectSetYInterval(ByRef rect As TRectangle, ByRef interval As TInterval)
With rect
    .bottom = interval.startValue
    .top = interval.endValue
    .isValid = interval.isValid
End With
End Sub

Public Function rectUnion(ByRef rect1 As TRectangle, ByRef rect2 As TRectangle) As TRectangle
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
    .isValid = True
End With
End Function

Public Sub rectValidate(ByRef rect As TRectangle)
With rect
    If .left <= .right And .bottom <= .top Then
        .isValid = True
    Else
        .isValid = False
    End If
End With
End Sub

'================================================================================
' Int64 functions
'================================================================================

Public Function Int64AddInt64( _
        ByRef num1 As Int64, _
        ByRef num2 As Int64) As Int64
Dim result As Int64
result = Int64SubtractInt64(num1, Int64ReverseSign(num2))
Int64AddInt64 = result
End Function

Public Function Int64SubtractInt64( _
        ByRef num1 As Int64, _
        ByRef num2 As Int64) As Int64
        
Dim result As Int64
Dim sign1 As Long
Dim sign2 As Long
Dim carry As Long
Dim sign As Long

sign1 = IIf((num1.lowpart And &H80000000), 1, 0)
sign2 = IIf((num2.lowpart And &H80000000), 1, 0)
result.lowpart = (num1.lowpart And &H7FFFFFFF) - (num2.lowpart And &H7FFFFFFF)
carry = IIf(result.lowpart And &H80000000, 1, 0)
sign = sign1 - (sign2 + carry)
If (sign And 1) Then
    result.lowpart = (result.lowpart Or &H80000000)
Else
    result.lowpart = (result.lowpart And &H7FFFFFFF)
End If
If (sign And 2) Then
    result.highpart = num1.highpart - (num2.highpart) - 1
Else
    result.highpart = num1.highpart - num2.highpart
End If
Int64SubtractInt64 = result
End Function

Public Function Int64ToDouble(ByRef num As Int64) As Double
Dim result As Double
result = CDbl(num.highpart) * 4294967296# + _
                IIf((num.lowpart And &H80000000) = &H80000000, 2147483648#, 0) + _
                CDbl(num.lowpart And &H7FFFFFFF)
Int64ToDouble = result
End Function

Public Function Int64ReverseSign(ByRef num As Int64) As Int64
Dim result As Int64
result.highpart = Not num.highpart
result.lowpart = (Not num.lowpart) + 1
Int64ReverseSign = result
End Function


