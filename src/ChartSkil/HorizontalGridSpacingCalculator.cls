VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HorizontalGridSpacingCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "HorizontalGridSpacingCalculator"

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

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcGridSpacings( _
                ByVal pHeight As Double, _
                ByVal pHeightCm As Double)
Dim numGridDivisions As Double
Dim gridSpacing As Double
Dim logSpacing As Double
Dim exp As Long
Dim mantissa As Double

Const ProcName As String = "calcGridSpacings"
Dim failpoint As String
On Error GoTo Err

numGridDivisions = pHeightCm / mGridlineSpacingY
gridSpacing = pHeight / numGridDivisions

logSpacing = Log(gridSpacing) / Log(10#)

exp = Int(logSpacing)
mantissa = gridSpacing / 10 ^ exp

If (mYScaleQuantum = 0.03125 Or _
    mYScaleQuantum = 0.015625 Or _
    mYScaleQuantum = 0.0078125) And _
    exp < 0 _
Then
    ' Tick size is 1/32 or a fraction thereof, which influences where we put
    ' the horizontal grid lines
    If exp >= -1 Then
        If mantissa >= 5 Then
            mScaleGridSpacingY = 16 * 0.03125
        ElseIf mantissa >= 2.5 Then
            mScaleGridSpacingY = 8 * 0.03125
        ElseIf mantissa >= 1.25 Then
            mScaleGridSpacingY = 4 * 0.03125
        ElseIf mantissa >= 0.625 Then
            mScaleGridSpacingY = 2 * 0.03125
        Else
            mScaleGridSpacingY = 0.03125
        End If
    Else
        mScaleGridSpacingY = 0.03125
    End If
Else
    If mantissa >= 5 Then
        mScaleGridSpacingY = 5 * 10 ^ exp
    ElseIf mantissa >= 2.5 Then
        mScaleGridSpacingY = 2.5 * 10 ^ exp
    ElseIf mantissa >= 2# Then
        mScaleGridSpacingY = 2 * 10 ^ exp
    Else
        mScaleGridSpacingY = 1 * 10 ^ exp
    End If
End If

If mScaleGridSpacingY < mYScaleQuantum Then mScaleGridSpacingY = mYScaleQuantum

If mIntegerYScale Then
    mYScaleSubFormatStr = "0"
Else
    If mScaleGridSpacingY < 1 Then
        mYScaleSubFormatStr = "0." & String(Abs(exp) + 2, "0")
    Else
        mYScaleSubFormatStr = "0.00"
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub


