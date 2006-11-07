Attribute VB_Name = "GMACD"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const MACDInputValue As String = "Input"

Public Const MACDParamLongPeriods As String = "Long periods"
Public Const MACDParamMAType As String = ParamMovingAverageType
Public Const MACDParamShortPeriods As String = "Short periods"
Public Const MACDParamSmoothingPeriods As String = "Smoothing periods"

Public Const MACDValueMACD As String = "MACD"
Public Const MACDValueMACDHist As String = "MACD hist"
Public Const MACDValueMACDLowerBalance As String = "MACD lower balance"
Public Const MACDValueMACDSignal As String = "MACD signal"
Public Const MACDValueStrength As String = "Strength"
Public Const MACDValueStrengthCount As String = "Strength count"
Public Const MACDValueMACDUpperBalance As String = "MACD upper balance"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Global object references
'================================================================================

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mDefaultParameters As IParameters
Private mStudyDefinition As IStudyDefinition

'================================================================================
' External function declarations
'================================================================================

'================================================================================
' Variables
'================================================================================

'================================================================================
' Procedures
'================================================================================

Public Property Let commonServiceConsumer( _
                ByVal value As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property


Public Property Let defaultParameters(ByVal value As IParameters)
' create a clone of the default parameters supplied by the caller
Set mDefaultParameters = value.Clone
End Property

Public Property Get defaultParameters() As IParameters
If mDefaultParameters Is Nothing Then
    Set mDefaultParameters = mCommonServiceConsumer.NewParameters
    mDefaultParameters.setParameterValue MACDParamShortPeriods, 12
    mDefaultParameters.setParameterValue MACDParamLongPeriods, 26
    mDefaultParameters.setParameterValue MACDParamSmoothingPeriods, 9
    mDefaultParameters.setParameterValue MACDParamMAType, EmaShortName
End If

' now create a clone of the default parameters for the caller
Set defaultParameters = mDefaultParameters.Clone
End Property

Public Property Get studyDefinition() As TradeBuildSP.IStudyDefinition
Dim inputDef As IStudyInputDefinition
Dim valueDef As IStudyValueDefinition
Dim paramDef As IStudyParameterDefinition

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = mCommonServiceConsumer.NewStudyDefinition
    mStudyDefinition.name = MacdName
    mStudyDefinition.shortName = MacdShortName
    mStudyDefinition.Description = "MACD (Moving Average Convergence/Divergence) " & _
                        "calculates the difference between two moving averages of " & _
                        "different periods. A further moving average is applied " & _
                        "to this difference to give a signal line. Finally the " & _
                        "difference between the MACD and the signal value gives " & _
                        "another indicator that is usually plotted as a " & _
                        "histogram."
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(MACDInputValue)
    inputDef.name = MACDInputValue
    inputDef.inputType = InputTypeDouble
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MACDValueMACD)
    valueDef.name = MACDValueMACD
    valueDef.Description = "The MACD value"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MACDValueMACDSignal)
    valueDef.name = MACDValueMACDSignal
    valueDef.Description = "The MACD signal value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MACDValueMACDHist)
    valueDef.name = MACDValueMACDHist
    valueDef.Description = "The MACD histogram value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MACDValueStrengthCount)
    valueDef.name = MACDValueStrengthCount
    valueDef.Description = "The number of consecutive bars for which the current " & _
                            "strength value has not changed"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MACDValueStrength)
    valueDef.name = MACDValueStrength
    valueDef.Description = "An indication of the strength of the current move"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MACDValueMACDUpperBalance)
    valueDef.name = MACDValueMACDUpperBalance
    valueDef.Description = "The price above which is confirmed strength"
    valueDef.defaultRegion = DefaultRegionPrice
    valueDef.valueType = ValueTypeDouble
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MACDValueMACDLowerBalance)
    valueDef.name = MACDValueMACDLowerBalance
    valueDef.Description = "The price below which is confirmed weakness"
    valueDef.defaultRegion = DefaultRegionPrice
    valueDef.valueType = ValueTypeDouble
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(MACDParamShortPeriods)
    paramDef.name = MACDParamShortPeriods
    paramDef.Description = "The number of periods in the shorter moving average"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(MACDParamLongPeriods)
    paramDef.name = MACDParamLongPeriods
    paramDef.Description = "The number of periods in the longer moving average"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(MACDParamSmoothingPeriods)
    paramDef.name = MACDParamSmoothingPeriods
    paramDef.Description = "The number of periods for smoothing the MACD to " & _
                            "produce the MACD signal value"
    paramDef.parameterType = ParameterTypeInteger
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(MACDParamMAType)
    paramDef.name = MACDParamMAType
    paramDef.Description = "The type of moving averages to be used"
    paramDef.parameterType = ParameterTypeString
    
End If

Set studyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================





