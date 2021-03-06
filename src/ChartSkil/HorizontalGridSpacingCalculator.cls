VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HorizontalGridSpacer"
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

Private Const ModuleName                            As String = "HorizontalGridSpacer"

'@================================================================================
' Member variables
'@================================================================================

Private mYGridlineSpacing                           As Double
Private mYScaleQuantum                              As Double

Private mYScaleGridSpacing                        As Double

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

Friend Property Let YGridlineSpacing(ByVal Value As Double)
Assert Value <> 0, "Value is 0"
mYGridlineSpacing = Value
End Property

Friend Property Get YScaleGridSpacing() As Single
YScaleGridSpacing = mYScaleGridSpacing
End Property

Friend Property Let YScaleQuantum(ByVal Value As Double)
mYScaleQuantum = Value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Calculate( _
                ByVal pHeight As Double, _
                ByVal pHeightCm As Double) As Double
Const ProcName As String = "Calculate"
On Error GoTo Err

Dim numGridDivisions As Long
numGridDivisions = Round(pHeightCm / mYGridlineSpacing, 0)
If numGridDivisions = 0 Then numGridDivisions = 1

Dim gridSpacing As Double
gridSpacing = pHeight / numGridDivisions

Dim logSpacing As Double
logSpacing = Log(gridSpacing) / Log10

Dim exp As Long
exp = Int(logSpacing)

Dim mantissa As Double
mantissa = gridSpacing / 10 ^ exp

If (mYScaleQuantum = 0.03125 Or _
    mYScaleQuantum = 0.015625 Or _
    mYScaleQuantum = 0.0078125) And _
    exp < 0 _
Then
    ' Tick size is 1/32 or a fraction thereof, which influences where we put
    ' the horizontal grid lines
    If exp >= -1 Then
        If mantissa >= 5.25 Then
            mYScaleGridSpacing = 32 * 0.03125
        ElseIf mantissa >= 2.75 Then
            mYScaleGridSpacing = 16 * 0.03125
        ElseIf mantissa >= 1.5 Then
            mYScaleGridSpacing = 8 * 0.03125
        ElseIf mantissa >= 0.65 Then
            mYScaleGridSpacing = 4 * 0.03125
        ElseIf mantissa >= 0.325 Then
            mYScaleGridSpacing = 2 * 0.03125
        Else
            mYScaleGridSpacing = 0.03125
        End If
    Else
        mYScaleGridSpacing = 0.03125
    End If
Else
    If mantissa > 5.25 Then
        mYScaleGridSpacing = 1 * 10 ^ (exp + 1)
    ElseIf mantissa >= 2.75 Then
        mYScaleGridSpacing = 5 * 10 ^ exp
    ElseIf mantissa >= 2.25 And 10 ^ exp <> mYScaleQuantum Then
        mYScaleGridSpacing = 2.5 * 10 ^ exp
    ElseIf mantissa >= 1.25 Then
        mYScaleGridSpacing = 2 * 10 ^ exp
    Else
        mYScaleGridSpacing = 1 * 10 ^ exp
    End If
End If

If mYScaleGridSpacing < mYScaleQuantum Then mYScaleGridSpacing = mYScaleQuantum

Calculate = mYScaleGridSpacing

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


