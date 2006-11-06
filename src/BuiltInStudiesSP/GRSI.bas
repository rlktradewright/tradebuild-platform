Attribute VB_Name = "GRSI"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const RsiInputValue As String = "Input"

Public Const RsiParamPeriods As String = ParamPeriods
Public Const RsiParamMovingAverageType As String = ParamMovingAverageType

Public Const RsiValueRsi As String = "Rsi"

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

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mDefaultParameters As IParameters
Private mStudyDefinition As IStudyDefinition

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
    mDefaultParameters.setParameterValue RsiParamPeriods, 14
    mDefaultParameters.setParameterValue RsiParamMovingAverageType, "SMA"
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
    mStudyDefinition.name = RsiName
    mStudyDefinition.shortName = RsiShortName
    mStudyDefinition.Description = "Relative Strength Indicator shows strength or " & _
                                    "weakness based on the gains and losses made " & _
                                    "during the specified number of periods"
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(RsiInputValue)
    inputDef.name = RsiInputValue
    inputDef.inputType = InputTypeDouble
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(RsiValueRsi)
    valueDef.name = RsiValueRsi
    valueDef.Description = "The Relative Strength Index value"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(RsiParamPeriods)
    paramDef.name = RsiParamPeriods
    paramDef.Description = "The number of periods used to calculate the RSI"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(RsiParamMovingAverageType)
    paramDef.name = RsiParamMovingAverageType
    paramDef.Description = "The type of moving average used to smooth the RSI"
    paramDef.parameterType = ParameterTypeInteger

End If

Set studyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================









