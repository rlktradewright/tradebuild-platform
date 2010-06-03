VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PriceFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IPriceFormatter

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

Private Const ModuleName                            As String = "PriceFormatter"

'@================================================================================
' Member variables
'@================================================================================

Private mYScaleQuantum                              As Double

Private mIntegerYScale                              As Boolean

Private mYScaleFormatStr                            As String
Private mYScaleSubFormatStr                         As String

Private mYScaleGridSpacing                        As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IPriceFormatter Interface Members
'@================================================================================

Private Function IPriceFormatter_FormatPrice( _
                ByVal value As Double) As String
Const ProcName As String = "IPriceFormatter_FormatPrice"
On Error GoTo Err

If (value - Int(value / mYScaleQuantum) * mYScaleQuantum) = 0 Then
    IPriceFormatter_FormatPrice = FormatPrice(value, mYScaleFormatStr)
Else
    IPriceFormatter_FormatPrice = FormatPrice(value, mYScaleSubFormatStr)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Let IPriceFormatter_IntegerYScale(ByVal value As Boolean)
mIntegerYScale = value
generateYScaleFormatStr
End Property

Private Property Let IPriceFormatter_YScaleGridSpacing(ByVal value As Double)
mYScaleGridSpacing = value
generateYScaleSubFormatStr
End Property

Private Property Let IPriceFormatter_YScaleQuantum(ByVal value As Double)
mYScaleQuantum = value
generateYScaleFormatStr
generateYScaleSubFormatStr
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Function FormatPrice( _
                ByVal value As Double, _
                ByRef pFormatString As String) As String
Dim fract As Single
Dim numerator As Long

Const ProcName As String = "ormatPrice"
Dim failpoint As String
On Error GoTo Err

If mYScaleQuantum = 0.03125 Then
    fract = value - Int(value)
    numerator = Int(fract / 0.03125)
    FormatPrice = Int(value) & "'" & Format(numerator, "00")
ElseIf mYScaleQuantum = 0.015625 Or mYScaleQuantum = 0.0078125 Then
    fract = value - Int(value)
    numerator = Int(fract / 0.0078125)
    FormatPrice = Int(value) & "'" & Format(numerator \ 4, "00")
    Select Case numerator Mod 4
    Case 0
        FormatPrice = FormatPrice & "0"
    Case 1
        FormatPrice = FormatPrice & "2"
    Case 2
        FormatPrice = FormatPrice & "5"
    Case 3
        FormatPrice = FormatPrice & "7"
    End Select
Else
    FormatPrice = Format(value, pFormatString)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Private Sub generateYScaleFormatStr()
Const ProcName As String = "generateYScaleFormatStr"
On Error GoTo Err

If mIntegerYScale Then
    mYScaleFormatStr = "0"
Else
    Dim minTickString As String
    minTickString = Format(mYScaleQuantum, "0.##############")

    If Len(minTickString) - 2 = 0 Then
        mYScaleFormatStr = "0"
    Else
        mYScaleFormatStr = "0." & String$(Len(minTickString) - 2, "0")
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub generateYScaleSubFormatStr()
Const ProcName As String = "generateYScaleSubFormatStr"
On Error GoTo Err

If mYScaleGridSpacing <= 0 Then Exit Sub

If mIntegerYScale Then
    mYScaleSubFormatStr = "0"
Else
    If mYScaleGridSpacing < 1 Then
        mYScaleSubFormatStr = "0." & String$(Abs(Int(Log(mYScaleGridSpacing / Log10))) + 2, "0")
    Else
        mYScaleSubFormatStr = "0.00"
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub
