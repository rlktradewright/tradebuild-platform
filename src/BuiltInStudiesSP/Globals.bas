Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const MaxDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023
Public Const MinDouble As Double = -(2 - 2 ^ -52) * 2 ^ 1023

Public Const DummyHigh As Double = MinDouble
Public Const DummyLow As Double = MaxDouble

Public Const DefaultStudyValueName As String = "$default"

' study name constants

Public Const AccDistName As String = "Accumulation/Distribution"
Public Const AccDistShortName As String = "AccDist"

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

Public Const PsName As String = "Parabolic Stop"
Public Const PsShortName As String = "PS"

Public Const RsiName As String = "Relative Strength Index"
Public Const RsiShortName As String = "RSI"

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

'Private mCommonServiceConsumer As ICommonServiceConsumer

'================================================================================
' Properties
'================================================================================

'Public Property Let commonServiceConsumer( _
'                ByVal value As ICommonServiceConsumer)
'Set mCommonServiceConsumer = value
'End Property

'================================================================================
' Methods
'================================================================================

Public Function gCreateMA( _
                ByVal maType As String, _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal periods As Long, _
                ByVal numberOfValuesToCache As Long) As IMovingAverageStudy
Dim lparams As IParameters
Dim lStudy As IStudy
Dim valueNames(0) As String

valueNames(0) = "in"

Select Case UCase$(maType)
Case UCase$(EmaShortName)
    Dim lEMA As EMA
    Set lEMA = New EMA
    Set lStudy = lEMA
    Set lparams = GEMA.defaultParameters
    lparams.setParameterValue ParamPeriods, periods
    lStudy.initialise commonServiceConsumer, _
                    studyServiceConsumer, _
                    commonServiceConsumer.GenerateGUIDString, _
                    lparams, _
                    numberOfValuesToCache, _
                    valueNames, _
                    Nothing
    Set gCreateMA = lEMA
Case Else
    Dim lSMA As SMA
    Set lSMA = New SMA
    Set lStudy = lSMA
    Set lparams = GSMA.defaultParameters
    lparams.setParameterValue ParamPeriods, periods
    lStudy.initialise commonServiceConsumer, _
                    studyServiceConsumer, _
                    commonServiceConsumer.GenerateGUIDString, _
                    lparams, _
                    numberOfValuesToCache, _
                    valueNames, _
                    Nothing
    Set gCreateMA = lSMA
End Select
End Function

'================================================================================
' Helper Function
'================================================================================





