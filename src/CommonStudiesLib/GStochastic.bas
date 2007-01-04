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


Private mDefaultParameters As Parameters
Private mStudyDefinition As StudyDefinition

'================================================================================
' External function declarations
'================================================================================

'================================================================================
' Variables
'================================================================================

'================================================================================
' Procedures
'================================================================================


Public Property Let defaultParameters(ByVal value As Parameters)
' create a clone of the default parameters supplied by the caller
Set mDefaultParameters = value.Clone
End Property

Public Property Get defaultParameters() As Parameters
If mDefaultParameters Is Nothing Then
    Set mDefaultParameters = New Parameters
    mDefaultParameters.setParameterValue StochParamKPeriods, 5
    mDefaultParameters.setParameterValue StochParamDPeriods, 3
End If

' now create a clone of the default parameters for the caller
Set defaultParameters = mDefaultParameters.Clone
End Property

Public Property Get StudyDefinition() As StudyDefinition
Dim inputDef As StudyInputDefinition
Dim valueDef As StudyValueDefinition
Dim paramDef As StudyParameterDefinition

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = New StudyDefinition
    mStudyDefinition.name = StochName
    mStudyDefinition.shortName = StochShortName
    mStudyDefinition.Description = "Stochastic compares the latest price to the " & _
                                "recent trading range, giving a value called %K. " & _
                                "It also has another value, called %D, which is " & _
                                "calculated by smoothing %K."
                                
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(StochInputValue)
    inputDef.inputType = InputTypeReal
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(StochValueK)
    valueDef.Description = "The stochastic value (%K)"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    valueDef.minimumValue = -5#
    valueDef.maximumValue = 105#
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(StochValueD)
    valueDef.Description = "The result of smoothing %K, also known as the signal line"
    valueDef.isDefault = False
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    valueDef.minimumValue = -5#
    valueDef.maximumValue = 105#
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(StochParamKPeriods)
    paramDef.Description = "The number of periods used to determine the recent " & _
                            "trading range"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(StochParamDPeriods)
    paramDef.Description = "The number of periods used to smooth the %K value " & _
                            "to obtain %D"
    paramDef.parameterType = ParameterTypeInteger

End If

Set StudyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================





