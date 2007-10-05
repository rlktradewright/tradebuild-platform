Attribute VB_Name = "GStandardDeviation"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Public Const SDInputValue As String = "Input"

Public Const SDParamPeriods As String = ParamPeriods

Public Const SDValueStandardDeviation As String = "Standard Deviation"

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
    mDefaultParameters.setParameterValue SDParamPeriods, 20
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
    mStudyDefinition.name = SdName
    mStudyDefinition.shortName = SdShortName
    mStudyDefinition.Description = "Standard Deviation " & _
                        "calculates the standard deviation of the n most " & _
                        "recent values, where n is given by the Periods parameter."
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(SDInputValue)
    inputDef.inputType = InputTypeReal
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SDValueStandardDeviation)
    valueDef.Description = "The standard deviation value"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SDParamPeriods)
    paramDef.Description = "The number of standard deviations used to calculate the " & _
                            "standard deviation"
    paramDef.parameterType = ParameterTypeInteger

End If

Set StudyDefinition = mStudyDefinition.Clone
End Property

'@================================================================================
' Helper Function
'@================================================================================









