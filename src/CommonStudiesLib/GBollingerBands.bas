Attribute VB_Name = "GBollingerBands"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Public Const BBInputPrice As String = "Price"

Public Const BBParamCentreBandWidth As String = "Centre band width"
Public Const BBParamDeviations As String = "Standard deviations"
Public Const BBParamEdgeBandWidth As String = "Edge band width"
Public Const BBParamMAType As String = ParamMovingAverageType
Public Const BBParamPeriods As String = ParamPeriods
Public Const BBParamSlopeThreshold As String = "Slope threshold"

Public Const BBValueBottom As String = "Bottom"
Public Const BBValueCentre As String = "Centre"
Public Const BBValueSpread As String = "Spread"
Public Const BBValueTop As String = "Top"

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
    mDefaultParameters.setParameterValue BBParamPeriods, 20
    mDefaultParameters.setParameterValue BBParamDeviations, 2
    mDefaultParameters.setParameterValue BBParamMAType, SmaShortName
    mDefaultParameters.setParameterValue BBParamCentreBandWidth, "0.0"
    mDefaultParameters.setParameterValue BBParamEdgeBandWidth, "0.0"
    mDefaultParameters.setParameterValue BBParamSlopeThreshold, "0.0"
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
    mStudyDefinition.name = BbName
    mStudyDefinition.shortName = BbShortName
    mStudyDefinition.Description = "Bollinger Bands " & _
                        "calculates upper and lower values that are a specified " & _
                        "number of standard deviations from a moving average. " & _
                        "When volatility increases, the bands widen, and they " & _
                        "narrow when volatility decreases."
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionNone
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(BBInputPrice)
    inputDef.inputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BBValueTop)
    valueDef.Description = "The top Bollinger band value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BBValueBottom)
    valueDef.Description = "The bottom Bollinger band value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BBValueCentre)
    valueDef.Description = "The MA value between the top and bottom bands"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BBValueSpread)
    valueDef.Description = "The difference between the top and bottom " & _
                            "band values"
    valueDef.defaultRegion = DefaultRegionCustom
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(BBParamPeriods)
    paramDef.Description = "The number of periods in the moving average"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(BBParamDeviations)
    paramDef.Description = "The number of standard deviations used to calculate the " & _
                            "values of the top and bottom bands"
    paramDef.parameterType = ParameterTypeReal

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(BBParamMAType)
    paramDef.Description = "The type of moving average to be used"
    paramDef.parameterType = ParameterTypeString
    paramDef.permittedValues = gMaTypes
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(BBParamCentreBandWidth)
    paramDef.Description = "The width of the central region"
    paramDef.parameterType = ParameterTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(BBParamEdgeBandWidth)
    paramDef.Description = "The width of the edge region"
    paramDef.parameterType = ParameterTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(BBParamSlopeThreshold)
    paramDef.Description = "The smallest slope value that is not to be considered flat"
    paramDef.parameterType = ParameterTypeReal
    
End If

Set StudyDefinition = mStudyDefinition.Clone
End Property

'@================================================================================
' Helper Function
'@================================================================================







