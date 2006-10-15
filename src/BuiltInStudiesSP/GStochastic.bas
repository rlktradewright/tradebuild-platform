Attribute VB_Name = "GStochastic"
Option Explicit

'================================================================================
' Constants
'================================================================================

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
Dim valueDef As IStudyValueDefinition
Dim paramDef As IStudyParameterDefinition

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = mCommonServiceConsumer.NewStudyDefinition
    mStudyDefinition.name = StochName
    mStudyDefinition.Description = "Stochastic compares the latest price to the " & _
                                "recent trading range, giving a value called %K. " & _
                                "It also has another value, called %D, which is " & _
                                "calculated by smoothing %K."
                                
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set valueDef = mCommonServiceConsumer.NewStudyValueDefinition
    valueDef.name = StochValueK
    valueDef.Description = "The stochastic value (%K)"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valuetype = ValueTypeDouble
    valueDef.minimumValue = 0#
    valueDef.maximumValue = 100#
    mStudyDefinition.StudyValueDefinitions.Add valueDef
    
    Set valueDef = mCommonServiceConsumer.NewStudyValueDefinition
    valueDef.name = StochValueD
    valueDef.Description = "The result of smoothing %K, also known as the signal line"
    valueDef.isDefault = False
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valuetype = ValueTypeDouble
    valueDef.minimumValue = 0#
    valueDef.maximumValue = 100#
    mStudyDefinition.StudyValueDefinitions.Add valueDef
    
    Set paramDef = mCommonServiceConsumer.NewStudyParameterDefinition
    paramDef.name = StochParamKPeriods
    paramDef.Description = "The number of periods used to determine the recent " & _
                            "trading range"
    paramDef.parameterType = ParameterTypeInteger
    mStudyDefinition.StudyParameterDefinitions.Add paramDef

    Set paramDef = mCommonServiceConsumer.NewStudyParameterDefinition
    paramDef.name = StochParamDPeriods
    paramDef.Description = "The number of periods used to smooth the %K value " & _
                            "to obtain %D"
    paramDef.parameterType = ParameterTypeInteger
    mStudyDefinition.StudyParameterDefinitions.Add paramDef

End If

Set studyDefinition = mStudyDefinition
End Property

'================================================================================
' Helper Function
'================================================================================





