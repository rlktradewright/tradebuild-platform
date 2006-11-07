Attribute VB_Name = "GSlowStochastic"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const SStochInputValue As String = "Input"

Public Const SStochParamKPeriods As String = "%K periods"
Public Const SStochParamKDPeriods As String = "%KD periods"
Public Const SStochParamDPeriods As String = "%D periods"

Public Const SStochValueK As String = "%K"
Public Const SStochValueD As String = "%D"

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
    mDefaultParameters.setParameterValue SStochParamKPeriods, 5
    mDefaultParameters.setParameterValue SStochParamKDPeriods, 3
    mDefaultParameters.setParameterValue SStochParamDPeriods, 3
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
    mStudyDefinition.name = SStochName
    mStudyDefinition.shortName = SStochShortName
    mStudyDefinition.Description = "Slow stochastic compares the latest price to the " & _
                                "recent trading range, and smoothes the result, " & _
                                "giving a value called %K. " & _
                                "It also has another value, called %D, which is " & _
                                "calculated by smoothing %K."
                                
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(SStochInputValue)
    inputDef.name = SStochInputValue
    inputDef.inputType = InputTypeDouble
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SStochValueK)
    valueDef.name = SStochValueK
    valueDef.Description = "The slow stochastic value (%K)"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    valueDef.minimumValue = 0#
    valueDef.maximumValue = 100#
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SStochValueD)
    valueDef.name = SStochValueD
    valueDef.Description = "The result of smoothing %K, also known as the signal line"
    valueDef.isDefault = False
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    valueDef.minimumValue = 0#
    valueDef.maximumValue = 100#
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SStochParamKPeriods)
    paramDef.name = SStochParamKPeriods
    paramDef.Description = "The number of periods used to determine the recent " & _
                            "trading range"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SStochParamKDPeriods)
    paramDef.name = SStochParamKDPeriods
    paramDef.name = SStochParamKDPeriods
    paramDef.Description = "The number of periods of smoothing used in " & _
                            "calculating %K"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SStochParamDPeriods)
    paramDef.name = SStochParamDPeriods
    paramDef.name = SStochParamDPeriods
    paramDef.Description = "The number of periods used to smooth the %K value " & _
                            "to obtain %D"
    paramDef.parameterType = ParameterTypeInteger

End If

Set studyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================







