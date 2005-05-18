VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements IRectangle

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private mLeft As Double
Private mRight As Double
Private mTop As Double
Private mBottom As Double

Private mXInterval As zInterval
Private mYInterval As zInterval

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mXInterval = New zInterval
Set mYInterval = New zInterval
mLeft = MinusInfinityDouble
mRight = PlusInfinityDouble
mBottom = MinusInfinityDouble
mTop = PlusInfinityDouble
End Sub

'================================================================================
' IRectangle Members
'================================================================================

Private Property Get IRectangle_bottom() As Double
IRectangle_bottom = bottom
End Property

Private Function IRectangle_intersection(ByVal aRectangle As ChartSkilTypes.IRectangle) As ChartSkilTypes.IRectangle
Set IRectangle_intersection = intersection(aRectangle)
End Function

Private Property Get IRectangle_left() As Double
IRectangle_left = left
End Property

Private Function IRectangle_overlaps(ByVal aRectangle As ChartSkilTypes.IRectangle) As Boolean
IRectangle_overlaps = overlaps(aRectangle)
End Function

Private Property Get IRectangle_right() As Double
IRectangle_right = right
End Property

Private Property Get IRectangle_top() As Double
IRectangle_top = top
End Property

Private Function IRectangle_XIntersection(ByVal aRectangle As ChartSkilTypes.IRectangle) As ChartSkilTypes.IInterval
Set IRectangle_XIntersection = XIntersection(aRectangle)
End Function

Private Property Get IRectangle_XInterval() As ChartSkilTypes.IInterval
Set IRectangle_XInterval = XInterval
End Property

Private Function IRectangle_YIntersection(ByVal aRectangle As ChartSkilTypes.IRectangle) As ChartSkilTypes.IInterval
Set IRectangle_YIntersection = YIntersection(aRectangle)
End Function

Private Property Get IRectangle_YInterval() As ChartSkilTypes.IInterval
Set IRectangle_YInterval = YInterval
End Property

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get bottom() As Double
bottom = mBottom
End Property

Public Property Let bottom(ByVal value As Double)
mBottom = value
mYInterval.startValue = value
End Property

Public Property Get left() As Double
left = mLeft
End Property

Public Property Let left(ByVal value As Double)
mLeft = value
mXInterval.startValue = value
End Property

Public Property Get right() As Double
right = mRight
End Property

Public Property Let right(ByVal value As Double)
mRight = value
mXInterval.endValue = value
End Property

Public Property Get top() As Double
top = mTop
End Property

Public Property Let top(ByVal value As Double)
mTop = value
mYInterval.endValue = value
End Property

Public Property Get XInterval() As zInterval
Set XInterval = mXInterval
End Property

Public Property Get YInterval() As zInterval
Set YInterval = mYInterval
End Property

'================================================================================
' Methods
'================================================================================

Public Function clone() As zRectangle
Set clone = New zRectangle
clone.bottom = mBottom
clone.left = mLeft
clone.right = mRight
clone.top = mTop
End Function

Public Function intersection(ByVal rect As zRectangle) As zRectangle
Dim xInt As zInterval
Dim yint As zInterval
If overlaps(rect) Then
    Set intersection = New zRectangle
    Set xInt = XInterval.intersection(rect.XInterval)
    Set yint = YInterval.intersection(rect.YInterval)
    intersection.left = xInt.startValue
    intersection.right = xInt.endValue
    intersection.bottom = yint.startValue
    intersection.top = yint.endValue
End If
End Function

Public Function overlaps(ByVal value As zRectangle) As Boolean
overlaps = XInterval.overlaps(value.XInterval) And _
            YInterval.overlaps(value.YInterval)
            
'overlaps = True
'If crosses(value.left, value.right, value.bottom, Me.left, Me.bottom, Me.top) Then Exit Function
'If crosses(value.left, value.right, value.top, Me.left, Me.bottom, Me.top) Then Exit Function
'If crosses(value.left, value.right, value.bottom, Me.right, Me.bottom, Me.top) Then Exit Function
'If crosses(value.left, value.right, value.top, Me.right, Me.bottom, Me.top) Then Exit Function
'If crosses(Me.left, Me.right, Me.bottom, value.left, value.bottom, value.top) Then Exit Function
'If crosses(Me.left, Me.right, Me.top, value.left, value.bottom, value.top) Then Exit Function
'If crosses(Me.left, Me.right, Me.bottom, value.right, value.bottom, value.top) Then Exit Function
'If crosses(Me.left, Me.right, Me.top, value.right, value.bottom, value.top) Then Exit Function
'If Me.left >= value.left And Me.left <= value.right And _
'    Me.right >= value.left And Me.left <= value.right And _
'    Me.bottom >= value.bottom And Me.bottom <= value.top And _
'    Me.top >= value.bottom And Me.bottom <= value.top _
'Then Exit Function
'If value.left >= Me.left And value.left <= Me.right And _
'    value.right >= Me.left And value.left <= Me.right And _
'    value.bottom >= Me.bottom And value.bottom <= Me.top And _
'    value.top >= Me.bottom And value.bottom <= Me.top _
'Then Exit Function
'
'overlaps = False
End Function

Public Function XIntersection(ByVal rect As zRectangle) As zInterval
Set XIntersection = XInterval.intersection(rect.XInterval)
End Function

Public Function YIntersection(ByVal rect As zRectangle) As zInterval
Set YIntersection = YInterval.intersection(rect.YInterval)
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Function crosses(ByVal x1Left As Double, _
                        ByVal x1Right As Double, _
                        ByVal y1 As Double, _
                        ByVal x2 As Double, _
                        ByVal y2Bottom As Double, _
                        ByVal y2Top As Double) As Boolean
                        
If (x2 >= x1Left And x2 <= x1Right) And _
    (y1 >= y2Bottom And y1 <= y2Top) _
Then
    crosses = True
End If
End Function


Private Function inside(ByVal x As Double, _
                        ByVal Y As Double, _
                        ByVal rect As zRectangle) As Boolean
If x >= rect.left And x <= rect.right And Y >= rect.bottom And Y <= rect.top Then
    inside = True
Else
    inside = False
End If
End Function

