VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zInterval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements IInterval

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private mStartValue As Double
Private mEndValue As Double

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mStartValue = MinusInfinityDouble
mEndValue = PlusInfinityDouble
End Sub

'================================================================================
' IInterval Members
'================================================================================

Private Property Get IInterval_endValue() As Double
IInterval_endValue = endValue
End Property

Private Function IInterval_intersection(ByVal anInterval As ChartSkilTypes.IInterval) As ChartSkilTypes.IInterval
Set IInterval_intersection = intersection(anInterval)
End Function

Private Function IInterval_overlaps(ByVal anInterval As ChartSkilTypes.IInterval) As Boolean
IInterval_overlaps = overlaps(anInterval)
End Function

Private Property Get IInterval_startValue() As Double
IInterval_startValue = startValue
End Property
'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get endValue() As Double
endValue = mEndValue
End Property

Public Property Let endValue(ByVal value As Double)
mEndValue = value
End Property

Public Property Get startValue() As Double
startValue = mStartValue
End Property

Public Property Let startValue(ByVal value As Double)
mStartValue = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function intersection(ByVal pInterval As zInterval) As zInterval
Dim startValue As Double
Dim endValue As Double

startValue = pInterval.startValue
endValue = pInterval.endValue
If mStartValue >= startValue And mStartValue <= endValue Then
    Set intersection = New zInterval
    intersection.startValue = mStartValue
    If mEndValue >= startValue And mEndValue <= endValue Then
        intersection.endValue = mEndValue
    Else
        intersection.endValue = endValue
    End If
    Exit Function
End If
If mEndValue >= startValue And mEndValue <= endValue Then
    Set intersection = New zInterval
    intersection.endValue = mEndValue
    intersection.startValue = startValue
    Exit Function
End If
If mStartValue < startValue And mEndValue > endValue Then
    Set intersection = New zInterval
    intersection.startValue = startValue
    intersection.endValue = endValue
End If
End Function

Public Function overlaps(ByVal pInterval As zInterval) As Boolean
Dim startValue As Double
Dim endValue As Double

overlaps = True

startValue = pInterval.startValue
endValue = pInterval.endValue

If mStartValue >= startValue And mStartValue <= endValue Then
    Exit Function
End If
If mEndValue >= startValue And mEndValue <= endValue Then
    Exit Function
End If
If mStartValue < startValue And mEndValue > endValue Then
    Exit Function
End If
overlaps = False
End Function

'================================================================================
' Helper Functions
'================================================================================


