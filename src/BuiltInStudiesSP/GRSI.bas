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


Private mDefaultParameters As Parameters
Private mStudyDefinition As StudyDefinition

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
    mDefaultParameters.setParameterValue RsiParamPeriods, 14
    mDefaultParameters.setParameterValue RsiParamMovingAverageType, "SMA"
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
    mStudyDefinition.name = RsiName
    mStudyDefinition.shortName = RsiShortName
    mStudyDefinition.Description = "Relative Strength Indicator shows strength or " & _
                                    "weakness based on the gains and losses made " & _
                                    "during the specified number of periods"
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(RsiInputValue)
    inputDef.inputType = InputTypeReal
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(RsiValueRsi)
    valueDef.Description = "The Relative Strength Index value"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.maximumValue = 105
    valueDef.minimumValue = -5
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(RsiParamPeriods)
    paramDef.Description = "The number of periods used to calculate the RSI"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(RsiParamMovingAverageType)
    paramDef.Description = "The type of moving average used to smooth the RSI"
    paramDef.parameterType = ParameterTypeString

End If

Set StudyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================









