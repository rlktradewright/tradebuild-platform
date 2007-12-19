VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
' @remarks
' @see
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

Private Const MaxDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023

'@================================================================================
' Member variables
'@================================================================================

Private mDataValues()           As Double
Private mDataValuesMaxIndex     As Long
Private mminimumDifference      As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mDataValues(1000) As Double
mDataValuesMaxIndex = -1
mminimumDifference = MaxDouble
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get minimumDifference() As Double
minimumDifference = mminimumDifference
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub addDataValue( _
                ByVal value As Double)
mDataValuesMaxIndex = mDataValuesMaxIndex + 1
If mDataValuesMaxIndex > UBound(mDataValues) Then
    ReDim Preserve mDataValues(UBound(mDataValues) + 1000) As Double
End If
mDataValues(mDataValuesMaxIndex) = value
End Sub

Public Sub analyze()
Dim i As Long
Dim diff As Double

SortDoubles mDataValues, False, 0, mDataValuesMaxIndex

For i = 0 To mDataValuesMaxIndex - 1
    diff = mDataValues(i + 1) - mDataValues(i)
    If diff <> 0 And diff < mminimumDifference Then mminimumDifference = diff
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


