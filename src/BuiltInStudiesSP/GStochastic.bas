Attribute VB_Name = "GStochastic"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const StochInputValue As String = "Input"

Public Const StochParamKPeriods As String = "%K periods"
Public Const StochParamDPeriods As String = "%D periods"

Public Const StochValueK As String = "%K"
Public Const StochValueD As String = "%D"

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
    mDefaultParameters.setParameterValue StochParamKPeriods, 5
    mDefaultParameters.setParameterValue StochParamDPeriods, 3
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
    mStudyDefinition.name = StochName
    mStudyDefinition.shortName = StochShortName
    mStudyDefinition.Description = "Stochastic compares the latest price to the " & _
                                "recent trading range, giving a value called %K. " & _
                                "It also has another value, called %D, which is " & _
                                "calculated by smoothing %K."
                                
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(StochInputValue)
    inputDef.name = StochInputValue
    inputDef.inputType = InputTypeDouble
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(StochValueK)
    valueDef.name = StochValueK
    valueDef.Description = "The stochastic value (%K)"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    valueDef.minimumValue = 0#
    valueDef.maximumValue = 100#
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(StochValueD)
    valueDef.name = StochValueD
    valueDef.Description = "The result of smoothing %K, also known as the signal line"
    valueDef.isDefault = False
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    valueDef.minimumValue = 0#
    valueDef.maximumValue = 100#
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(StochParamKPeriods)
    paramDef.name = StochParamKPeriods
    paramDef.Description = "The number of periods used to determine the recent " & _
                            "trading range"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(StochParamDPeriods)
    paramDef.name = StochParamDPeriods
    paramDef.Description = "The number of periods used to smooth the %K value " & _
                            "to obtain %D"
    paramDef.parameterType = ParameterTypeInteger

End If

Set studyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================





