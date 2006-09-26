Attribute VB_Name = "GBollingerBands"
Option Explicit

'================================================================================
' Constants
'================================================================================

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

Public Property Get studyDefinition() As TradeBuildSP.IStudyDefinition
Dim valueDef As IStudyValueDefinition
Dim paramDef As IStudyParameterDefinition

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = mCommonServiceConsumer.NewStudyDefinition
    mStudyDefinition.name = BbName
    mStudyDefinition.Description = "Bollinger Bands " & _
                        "calculates upper and lower values that are a specified " & _
                        "number of standard deviations from a moving average. " & _
                        "When volatility increases, the bands widen, and they " & _
                        "narrow when volatility decreases."
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionPrice
    
    Set valueDef = mCommonServiceConsumer.NewStudyValueDefinition
    valueDef.name = BBValueTop
    valueDef.Description = "The top Bollinger band value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valuetype = ValueTypeDouble
    mStudyDefinition.StudyValueDefinitions.Add valueDef
    
    Set valueDef = mCommonServiceConsumer.NewStudyValueDefinition
    valueDef.name = BBValueBottom
    valueDef.Description = "The bottom Bollinger band value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valuetype = ValueTypeDouble
    mStudyDefinition.StudyValueDefinitions.Add valueDef
    
    Set valueDef = mCommonServiceConsumer.NewStudyValueDefinition
    valueDef.name = BBValueCentre
    valueDef.Description = "The MA value between the top and bottom bands"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valuetype = ValueTypeDouble
    mStudyDefinition.StudyValueDefinitions.Add valueDef
    
    Set valueDef = mCommonServiceConsumer.NewStudyValueDefinition
    valueDef.name = BBValueSpread
    valueDef.Description = "The difference between the top and bottom " & _
                            "band values"
    valueDef.defaultRegion = DefaultRegionCustom
    valueDef.valuetype = ValueTypeDouble
    mStudyDefinition.StudyValueDefinitions.Add valueDef
    
    Set paramDef = mCommonServiceConsumer.NewStudyParameterDefinition
    paramDef.name = BBParamPeriods
    paramDef.Description = "The number of periods in the moving average"
    paramDef.parameterType = ParameterTypeInteger
    mStudyDefinition.StudyParameterDefinitions.Add paramDef

    Set paramDef = mCommonServiceConsumer.NewStudyParameterDefinition
    paramDef.name = BBParamDeviations
    paramDef.Description = "The number of standard deviations used to calculate the " & _
                            "values of the top and bottom bands"
    paramDef.parameterType = ParameterTypeDouble
    mStudyDefinition.StudyParameterDefinitions.Add paramDef

    Set paramDef = mCommonServiceConsumer.NewStudyParameterDefinition
    paramDef.name = BBParamMAType
    paramDef.Description = "The type of moving average to be used"
    paramDef.parameterType = ParameterTypeString
    mStudyDefinition.StudyParameterDefinitions.Add paramDef
    
    Set paramDef = mCommonServiceConsumer.NewStudyParameterDefinition
    paramDef.name = BBParamCentreBandWidth
    paramDef.Description = "The width of the central region"
    paramDef.parameterType = ParameterTypeDouble
    mStudyDefinition.StudyParameterDefinitions.Add paramDef
    
    Set paramDef = mCommonServiceConsumer.NewStudyParameterDefinition
    paramDef.name = BBParamEdgeBandWidth
    paramDef.Description = "The width of the edge region"
    paramDef.parameterType = ParameterTypeDouble
    mStudyDefinition.StudyParameterDefinitions.Add paramDef
    
    Set paramDef = mCommonServiceConsumer.NewStudyParameterDefinition
    paramDef.name = BBParamSlopeThreshold
    paramDef.Description = "The smallest slope value that is not to be considered flat"
    paramDef.parameterType = ParameterTypeDouble
    mStudyDefinition.StudyParameterDefinitions.Add paramDef
    
End If

Set studyDefinition = mStudyDefinition
End Property

'================================================================================
' Helper Function
'================================================================================







