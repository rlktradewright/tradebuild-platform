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

Public Const SmaName As String = "Simple Moving Average"
Public Const SmaShortName As String = "SMA"

Public Const SwingName As String = "Swing"
Public Const SwingShortName As String = "Swing"

' generic study parameter names - these are parameter names that are common
' to many studies
Public Const ParamMovingAverageType As String = "Mov avg type"
Public Const ParamPeriods As String = "Periods"

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
                ByVal params As IParameters) As String
Dim i As Long
Dim param As IParameter

Set param = params.getFirstParameter
If Not param Is Nothing Then gParamsToString = param.value

Set param = params.getNextParameter
Do
    If Not param Is Nothing Then
        gParamsToString = gParamsToString & "," & param.value
    Else
        Exit Do
    End If
    Set param = params.getNextParameter
Loop

End Function

'================================================================================
' Helper Function
'================================================================================





