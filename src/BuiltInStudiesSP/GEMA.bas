Attribute VB_Name = "GEMA"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const EmaInputValue As String = "Input"

Public Const EMAParamPeriods As String = ParamPeriods
Public Const EMAParamSlopeThreshold As String = "Slope threshold"

Public Const EMAValueEMA As String = "MA"

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
    mDefaultParameters.setParameterValue EMAParamPeriods, 21
    mDefaultParameters.setParameterValue EMAParamSlopeThreshold, "0.0"
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
    mStudyDefinition.name = EmaName
    mStudyDefinition.shortName = EmaShortName
    mStudyDefinition.Description = "An exponential moving average"
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionPrice
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(EmaInputValue)
    inputDef.name = EmaInputValue
    inputDef.inputType = InputTypeDouble
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(EMAValueEMA)
    valueDef.name = EMAValueEMA
    valueDef.Description = "The moving average value"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(EMAParamPeriods)
    paramDef.name = EMAParamPeriods
    paramDef.Description = "The number of periods used to calculate the moving average"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(EMAParamSlopeThreshold)
    paramDef.name = EMAParamSlopeThreshold
    paramDef.Description = "The smallest slope value that is not to be considered flat"
    paramDef.parameterType = ParameterTypeDouble
    
End If

Set studyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================





