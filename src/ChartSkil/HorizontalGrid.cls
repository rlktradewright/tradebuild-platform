VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HorizontalGrid"
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

Private Const ModuleName                            As String = "HorizontalGrid"

'@================================================================================
' Member variables
'@================================================================================

Private mYGridTexts As TextSeries
Private mYGridLines As LineSeries

Private mScaleGridSpacingY As Double
Private mYScaleFormatStr As String
Private mYScaleSubFormatStr As String

' indicates whether the grid is currently
' hidden
Private mHideGrid As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

Friend Property Get ScaleGridSpacingY() As Single
ScaleGridSpacingY = mScaleGridSpacingY
End Property

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcGridSpacings()
Dim numGridDivisions As Double
Dim gridSpacing As Double
Dim log10 As Double
Dim logSpacing As Double
Dim exp As Long
Dim mantissa As Double

numGridDivisions = mCanvas.HeightCm / GridlineSpacingY
gridSpacing = mCanvas.Height / numGridDivisions

If HasGrid And Not mHideGrid Then
    log10 = Log(10#)
    logSpacing = Log(gridSpacing) / log10
    
    exp = Fix(logSpacing)
    mantissa = Abs(logSpacing) - Abs(exp)
    
    If logSpacing < 0 Then
        exp = exp - 1
        mantissa = 1 - mantissa
    End If
    
    If YScaleQuantum = 0.03125 And _
        exp < 0 _
    Then
        ' Tick size is 1/32, which influences where we put the horizontal grid lines
        If exp >= -1 Then
            If mantissa >= Log(5) / log10 Then
                mScaleGridSpacingY = 16 * 0.03125
            ElseIf mantissa >= Log(2.5) / log10 Then
                mScaleGridSpacingY = 8 * 0.03125
            ElseIf mantissa >= Log(1.25) / log10 Then
                mScaleGridSpacingY = 4 * 0.03125
            ElseIf mantissa >= Log(0.625) / log10 Then
                mScaleGridSpacingY = 2 * 0.03125
            Else
                mScaleGridSpacingY = 0.03125
            End If
        Else
            mScaleGridSpacingY = 0.03125
        End If
    Else
        If mantissa >= Log(5) / log10 Then
            mScaleGridSpacingY = 5 * 10 ^ exp
        ElseIf mantissa >= Log(2.5) / log10 Then
            mScaleGridSpacingY = 2.5 * 10 ^ exp
        ElseIf mantissa >= Log(2#) / log10 Then
            mScaleGridSpacingY = 2# * 10 ^ exp
        Else
            mScaleGridSpacingY = 10 ^ exp
        End If
    End If
    If IntegerYScale Then
            mYScaleFormatStr = "0"
            mYScaleSubFormatStr = "0"
    Else
        If mScaleGridSpacingY < 1 Then
            mYScaleFormatStr = "0." & String(Abs(exp) + 1, "0")
            mYScaleSubFormatStr = "0." & String(Abs(exp) + 2, "0")
        Else
            mYScaleFormatStr = "0.0"
            mYScaleSubFormatStr = "0.00"
        End If
    End If
End If

End Sub

