Attribute VB_Name = "Globals"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

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

Public Const ConstMomentumBarsName As String = "Constant momentum bars"
Public Const ConstMomentumBarsShortName As String = "CM Bars"

Public Const ConstTimeBarsName As String = "Constant time bars"
Public Const ConstTimeBarsShortName As String = "Bars"

Public Const ConstVolBarsName As String = "Constant volume bars"
Public Const ConstVolBarsShortName As String = "CV Bars"

Public Const DoncName As String = "Donchian Channels"
Public Const DoncShortName As String = "Donc"

Public Const EmaName As String = "Exponential Moving Average"
Public Const EmaShortName As String = "EMA"

Public Const FIName As String = "Force Index"
Public Const FIShortName As String = "FI"

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

' sub-value names for study values in bar mode
Public Const BarValueOpen As String = "Open"
Public Const BarValueHigh As String = "High"
Public Const BarValueLow As String = "Low"
Public Const BarValueClose As String = "Close"
Public Const BarValueVolume As String = "Volume"
Public Const BarValueTickVolume As String = "Tick Volume"
Public Const BarValueHL2 As String = "(H+L)/2"
Public Const BarValueHLC3 As String = "(H+L+C)/3"
Public Const BarValueOHLC4 As String = "(O+H+L+C)/4"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Global object references
'@================================================================================

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Variables
'@================================================================================

Public gLibraryManager As StudyLibraryManager

'@================================================================================
' Properties
'@================================================================================

'Public Property Let commonServiceConsumer( _
'                ByVal value As ICommonServiceConsumer)
'Set mCommonServiceConsumer = value
'End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gCreateMA( _
                ByVal maType As String, _
                ByVal periods As Long, _
                ByVal numberOfValuesToCache As Long) As Study
Dim lparams As Parameters
Dim lStudy As Study
Dim valueNames(0) As String

valueNames(0) = "in"

Select Case UCase$(maType)
Case UCase$(EmaShortName)
    Dim lEMA As EMA
    Set lEMA = New EMA
    Set lStudy = lEMA
    Set lparams = GEMA.defaultParameters
    lparams.setParameterValue ParamPeriods, periods
    lStudy.initialise GenerateGUIDString, _
                    lparams, _
                    numberOfValuesToCache, _
                    valueNames, _
                    Nothing, _
                    Nothing
                    
    Set gCreateMA = lEMA
Case Else
    Dim lSMA As SMA
    Set lSMA = New SMA
    Set lStudy = lSMA
    Set lparams = GSMA.defaultParameters
    lparams.setParameterValue ParamPeriods, periods
    lStudy.initialise GenerateGUIDString, _
                    lparams, _
                    numberOfValuesToCache, _
                    valueNames, _
                    Nothing, _
                    Nothing
    Set gCreateMA = lSMA
End Select
End Function

Public Function gMaTypes() As Variant()
Dim ar(1) As Variant
ar(0) = EmaShortName
ar(1) = SmaShortName
gMaTypes = ar
End Function

'@================================================================================
' Helper Function
'@================================================================================





