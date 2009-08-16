Attribute VB_Name = "GSlowStochastic"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Public Const SStochInputValue As String = "Input"

Public Const SStochParamKPeriods As String = "%K periods"
Public Const SStochParamKDPeriods As String = "%KD periods"
Public Const SStochParamDPeriods As String = "%D periods"

Public Const SStochValueK As String = "%K"
Public Const SStochValueD As String = "%D"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Global object references
'@================================================================================


Private mDefaultParameters As Parameters
Private mStudyDefinition As StudyDefinition

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Variables
'@================================================================================

'@================================================================================
' Procedures
'@================================================================================


Public Property Let defaultParameters(ByVal value As Parameters)
' create a clone of the default parameters supplied by the caller
Set mDefaultParameters = value.Clone
End Property

Public Property Get defaultParameters() As Parameters
If mDefaultParameters Is Nothing Then
    Set mDefaultParameters = New Parameters
    mDefaultParameters.setParameterValue SStochParamKPeriods, 5
    mDefaultParameters.setParameterValue SStochParamKDPeriods, 3
    mDefaultParameters.setParameterValue SStochParamDPeriods, 3
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
    mStudyDefinition.name = SStochName
    mStudyDefinition.shortName = SStochShortName
    mStudyDefinition.Description = "Slow stochastic compares the latest price to the " & _
                                "recent trading range, and smoothes the result, " & _
                                "giving a value called %K. " & _
                                "It also has another value, called %D, which is " & _
                                "calculated by smoothing %K."
                                
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(SStochInputValue)
    inputDef.inputType = InputTypeReal
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SStochValueK)
    valueDef.Description = "The slow stochastic value (%K)"
    valueDef.IncludeInChart = True
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    valueDef.minimumValue = -5#
    valueDef.valueStyle = gCreateDataPointStyle(vbBlue)
    valueDef.maximumValue = 105#
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SStochValueD)
    valueDef.Description = "The result of smoothing %K, also known as the signal line (%D)"
    valueDef.IncludeInChart = True
    valueDef.isDefault = False
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    valueDef.minimumValue = -5#
    valueDef.valueStyle = gCreateDataPointStyle(vbRed)
    valueDef.maximumValue = 105#
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SStochParamKPeriods)
    paramDef.Description = "The number of periods used to determine the recent " & _
                            "trading range"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SStochParamKDPeriods)
    paramDef.Description = "The number of periods of smoothing used in " & _
                            "calculating %K"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SStochParamDPeriods)
    paramDef.Description = "The number of periods used to smooth the %K value " & _
                            "to obtain %D"
    paramDef.parameterType = ParameterTypeInteger

End If

Set StudyDefinition = mStudyDefinition.Clone
End Property

'@================================================================================
' Helper Function
'@================================================================================







