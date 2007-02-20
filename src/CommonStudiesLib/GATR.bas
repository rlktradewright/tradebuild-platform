Attribute VB_Name = "GATR"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const AtrInputPrice As String = "Price"

Public Const AtrParamMAType As String = ParamMovingAverageType
Public Const AtrParamPeriods As String = ParamPeriods

Public Const AtrValueATR As String = "ATR"

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
    mDefaultParameters.setParameterValue AtrParamPeriods, 27
    mDefaultParameters.setParameterValue AtrParamMAType, EmaShortName
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
    mStudyDefinition.name = AtrName
    mStudyDefinition.shortName = AtrShortName
    mStudyDefinition.Description = "Average True Range " & _
                        "calculates the moving average of the 'true ranges' of bars " & _
                        "over the specified number of periods. " & _
                        "The true range of a bar is calculated by substituting the " & _
                        "previous close for the bar's low (if lower), or for the high (if higher)."
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(AtrInputPrice)
    inputDef.inputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(AtrValueATR)
    valueDef.Description = "The Average True Range value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(AtrParamPeriods)
    paramDef.Description = "The number of periods used to calculate the Average True Range"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(AtrParamMAType)
    paramDef.Description = "The type of moving average to be used"
    paramDef.parameterType = ParameterTypeString
    paramDef.permittedValues = gMaTypes
    
End If

Set StudyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================









