Attribute VB_Name = "GStandardDeviation"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const SDInputValue As String = "Input"

Public Const SDParamPeriods As String = ParamPeriods

Public Const SDValueStandardDeviation As String = "Standard Deviation"

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
    mDefaultParameters.setParameterValue SDParamPeriods, 20
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
    mStudyDefinition.name = SdName
    mStudyDefinition.shortName = SdShortName
    mStudyDefinition.Description = "Standard Deviation " & _
                        "calculates the standard deviation of the n most " & _
                        "recent values, where n is given by the Periods parameter."
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(SDInputValue)
    inputDef.name = SDInputValue
    inputDef.inputType = InputTypeDouble
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SDValueStandardDeviation)
    valueDef.name = SDValueStandardDeviation
    valueDef.Description = "The standard deviation value"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SDParamPeriods)
    paramDef.name = SDParamPeriods
    paramDef.Description = "The number of standard deviations used to calculate the " & _
                            "standard deviation"
    paramDef.parameterType = ParameterTypeInteger

End If

Set studyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================









