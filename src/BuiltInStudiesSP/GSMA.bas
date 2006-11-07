Attribute VB_Name = "GSMA"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const SMAInputValue As String = "Input"

Public Const SMAParamPeriods As String = ParamPeriods
Public Const SMAParamSlopeThreshold As String = "Slope threshold"

Public Const SMAValueSMA As String = "MA"

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
    mDefaultParameters.setParameterValue SMAParamPeriods, 21
    mDefaultParameters.setParameterValue SMAParamSlopeThreshold, "0.0"
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
    mStudyDefinition.name = SmaName
    mStudyDefinition.shortName = SmaShortName
    mStudyDefinition.Description = "A simple moving average"
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionPrice
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(SMAInputValue)
    inputDef.name = SMAInputValue
    inputDef.inputType = InputTypeDouble
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SMAValueSMA)
    valueDef.name = SMAValueSMA
    valueDef.Description = "The moving average value"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SMAParamPeriods)
    paramDef.name = SMAParamPeriods
    paramDef.Description = "The number of periods used to calculate the moving average"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SMAParamSlopeThreshold)
    paramDef.name = SMAParamSlopeThreshold
    paramDef.Description = "The smallest slope value that is not to be considered flat"
    paramDef.parameterType = ParameterTypeDouble
    
End If

Set studyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================



