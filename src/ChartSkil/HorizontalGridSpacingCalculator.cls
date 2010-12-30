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

Private mGridlineSpacingY                           As Double
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

Friend Property Let GridlineSpacingY(ByVal Value As Double)
mGridlineSpacingY = Value
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

logSpacing = Log(gridSpacing) / Log10

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
            mYScaleGridSpacing = 16 * 0.03125
        ElseIf mantissa >= 2.5 Then
            mYScaleGridSpacing = 8 * 0.03125
        ElseIf mantissa >= 1.25 Then
            mYScaleGridSpacing = 4 * 0.03125
        ElseIf mantissa >= 0.625 Then
            mYScaleGridSpacing = 2 * 0.03125
        Else
            mYScaleGridSpacing = 0.03125
        End If
    Else
        mYScaleGridSpacing = 0.03125
    End If
Else
    If mantissa >= 5 Then
        mYScaleGridSpacing = 5 * 10 ^ exp
    ElseIf mantissa >= 2.5 Then
        mYScaleGridSpacing = 2.5 * 10 ^ exp
    ElseIf mantissa >= 2# Then
        mYScaleGridSpacing = 2 * 10 ^ exp
    Else
        mYScaleGridSpacing = 1 * 10 ^ exp
    End If
End If

If mYScaleGridSpacing < mYScaleQuantum Then mYScaleGridSpacing = mYScaleQuantum

Calculate = mYScaleGridSpacing

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function

'@================================================================================
' Helper Functions
'@================================================================================


