Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const IncreasedValueColor As Long = &HB7E43
Public Const DecreasedValueColor As Long = &H4444EB

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Global object references
'================================================================================

'================================================================================
' External function declarations
'================================================================================

Public Declare Sub InitCommonControls Lib "comctl32" ()

'================================================================================
' Variables
'================================================================================

'================================================================================
' Procedures
'================================================================================

Public Function gIsInteger( _
                ByVal value As String, _
                Optional ByVal minValue As Long = 0, _
                Optional ByVal maxValue As Long = &H7FFFFFFF) As Boolean
Dim quantity As Long

On Error GoTo err

If IsNumeric(value) Then
    quantity = CLng(value)
    If CDbl(value) - quantity = 0 Then
        If quantity >= minValue And quantity <= maxValue Then
            gIsInteger = True
        End If
    End If
End If
                
Exit Function

err:
If err.Number <> TradeBuild.ErrorCodes.ErrOverflow Then err.Raise err.Number
End Function


