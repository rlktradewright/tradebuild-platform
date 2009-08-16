Attribute VB_Name = "GSMA"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Public Const SMAInputValue As String = "Input"

Public Const SMAParamPeriods As String = ParamPeriods
Public Const SMAParamSlopeThreshold As String = "Slope threshold"

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
    mDefaultParameters.setParameterValue SMAParamPeriods, 21
    mDefaultParameters.setParameterValue SMAParamSlopeThreshold, "0.0"
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
    mStudyDefinition.name = SmaName
    mStudyDefinition.shortName = SmaShortName
    mStudyDefinition.Description = "A simple moving average"
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionNone
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(SMAInputValue)
    inputDef.inputType = InputTypeReal
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MovingAverageStudyValueName)
    valueDef.Description = "The moving average value"
    valueDef.IncludeInChart = True
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueStyle = gCreateDataPointStyle(&H1D9311)
    valueDef.valueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SMAParamPeriods)
    paramDef.Description = "The number of periods used to calculate the moving average"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SMAParamSlopeThreshold)
    paramDef.Description = "The smallest slope value that is not to be considered flat"
    paramDef.parameterType = ParameterTypeReal
    
End If

Set StudyDefinition = mStudyDefinition.Clone
End Property

'@================================================================================
' Helper Function
'@================================================================================



