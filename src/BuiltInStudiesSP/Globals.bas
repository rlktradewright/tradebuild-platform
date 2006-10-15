Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const DummyHigh As Double = -(2 - 2 ^ -52) * 2 ^ 1023
Public Const DummyLow As Double = (2 - 2 ^ -52) * 2 ^ 1023

Public Const DefaultStudyValueName As String = "$default"

' study name constants

Public Const AtrName As String = "Average True Range"
Public Const AtrShortName As String = "ATR"

Public Const BbName As String = "Bollinger Bands"
Public Const BbShortName As String = "BB"

Public Const DoncName As String = "Donchian Channels"
Public Const DoncShortName As String = "Donc"

Public Const EmaName As String = "Exponential Moving Average"
Public Const EmaShortName As String = "EMA"

Public Const MacdName As String = "MACD"
Public Const MacdShortName As String = "MACD"

Public Const SdName As String = "Standard Deviation"
Public Const SdShortName As String = "SD"

Public Const SStochName As String = "Slow Stochastic"
Public Const SStochShortName As String = "SStoch"

Public Const SmaName As String = "Simple Moving Average"
Public Const SmaShortName As String = "SMA"

Public Const StochName As String = "Stochastic"
Public Const StochShortName As String = "Stoch"

Public Const SwingName As String = "Swing"
Public Const SwingShortName As String = "Swing"

' generic study parameter names - these are parameter names that are common
' to many studies
Public Const ParamMovingAverageType As String = "Mov avg type"
Public Const ParamPeriods As String = "Periods"

'================================================================================
' Enums
'================================================================================

Public Enum TaskDiscriminators
    TaskAddStudy
    TaskAddStudyValueListener
End Enum
'================================================================================
' Types
'================================================================================

'================================================================================
' Global object references
'================================================================================

'================================================================================
' External function declarations
'================================================================================

'================================================================================
' Variables
'================================================================================

'================================================================================
' Procedures
'================================================================================

Public Function gCreateMA( _
                ByVal maType As String) As MovingAverageStudy
Select Case UCase$(maType)
Case UCase$(SmaShortName)
    Set gCreateMA = New SMA
Case UCase$(EmaShortName)
    Set gCreateMA = New EMA
End Select
End Function

Public Function gParamsToString( _
                ByVal params As IParameters, _
                ByVal studyDef As TradeBuildSP.IStudyDefinition) As String
Dim paramDefs As TradeBuildSP.IStudyParameterDefinitions
Dim paramDef As TradeBuildSP.IStudyParameterDefinition
Dim i As Long

On Error Resume Next
Set paramDefs = studyDef.StudyParameterDefinitions
For i = 1 To paramDefs.Count
    Set paramDef = paramDefs.Item(i)
    If Len(gParamsToString) = 0 Then
        gParamsToString = params.getParameterValue(paramDef.name)
    Else
        gParamsToString = gParamsToString & "," & params.getParameterValue(paramDef.name)
    End If
Next
End Function

'================================================================================
' Helper Function
'================================================================================





