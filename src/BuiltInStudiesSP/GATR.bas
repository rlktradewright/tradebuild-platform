Attribute VB_Name = "GATR"
Option Explicit

'================================================================================
' Constants
'================================================================================

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
    mDefaultParameters.setParameterValue AtrParamPeriods, 27
    mDefaultParameters.setParameterValue AtrParamMAType, EmaShortName
End If

' now create a clone of the default parameters for the caller
Set defaultParameters = mDefaultParameters.Clone
End Property

Public Property Get studyDefinition() As TradeBuildSP.IStudyDefinition
Dim valueDef As IStudyValueDefinition
Dim paramDef As IStudyParameterDefinition

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = mCommonServiceConsumer.NewStudyDefinition
    mStudyDefinition.name = AtrName
    mStudyDefinition.Description = "Average True Range " & _
                        "calculates the moving average of the 'true ranges' of bars " & _
                        "over the specified number of periods. " & _
                        "The true range of a bar is calculated by substituting the " & _
                        "previous close for the bar's low (if lower), or for the high (if higher)."
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set valueDef = mCommonServiceConsumer.NewStudyValueDefinition
    valueDef.name = AtrValueATR
    valueDef.Description = "The Average True Range value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valuetype = ValueTypeDouble
    mStudyDefinition.StudyValueDefinitions.Add valueDef
    
    Set paramDef = mCommonServiceConsumer.NewStudyParameterDefinition
    paramDef.name = AtrParamPeriods
    paramDef.Description = "The number of periods used to calculate the Average True Range"
    paramDef.parameterType = ParameterTypeInteger
    mStudyDefinition.StudyParameterDefinitions.Add paramDef

    Set paramDef = mCommonServiceConsumer.NewStudyParameterDefinition
    paramDef.name = AtrParamMAType
    paramDef.Description = "The type of moving average to be used"
    paramDef.parameterType = ParameterTypeString
    mStudyDefinition.StudyParameterDefinitions.Add paramDef
    
End If

Set studyDefinition = mStudyDefinition
End Property

'================================================================================
' Helper Function
'================================================================================









