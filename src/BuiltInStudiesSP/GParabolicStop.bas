Attribute VB_Name = "GParabolicStop"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const PsParamStartFactor As String = "Start factor"
Public Const PsParamIncrement As String = "Increment"
Public Const PsParamMaxFactor As String = "Max factor"

Public Const PsValuePs As String = "PS"

Public Const PsDefaultStartFactor As Double = 0.02
Public Const PsDefaultIncrement As Double = 0.02
Public Const PsDefaultMaxFactor As Double = 0.2

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
    mDefaultParameters.setParameterValue PsParamStartFactor, PsDefaultStartFactor
    mDefaultParameters.setParameterValue PsParamIncrement, PsDefaultIncrement
    mDefaultParameters.setParameterValue PsParamMaxFactor, PsDefaultMaxFactor
End If

' now create a clone of the default parameters for the caller
Set defaultParameters = mDefaultParameters.Clone
End Property

Public Property Get studyDefinition() As TradeBuildSP.IStudyDefinition
Dim valueDef As IStudyValueDefinition
Dim paramDef As IStudyParameterDefinition

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = mCommonServiceConsumer.NewStudyDefinition
    mStudyDefinition.name = PsName
    mStudyDefinition.Description = "Parabolic Stop calculates a value that can be used " & _
                                    "as a stop loss for trades. When the market is " & _
                                    "rising, the value increases with each period. When " & _
                                    "the market is falling, the value decreases with " & _
                                    "each period."
                                    
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionPrice
    
    Set valueDef = mCommonServiceConsumer.NewStudyValueDefinition
    valueDef.name = PsValuePs
    valueDef.Description = "The parabolic stop value"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valuetype = ValueTypeDouble
    mStudyDefinition.StudyValueDefinitions.Add valueDef
    
    Set paramDef = mCommonServiceConsumer.NewStudyParameterDefinition
    paramDef.name = PsParamStartFactor
    paramDef.Description = "The initial value of the acceleration factor that governs " & _
                            "the increase in the speed with which the parabolic stop " & _
                            "rises or falls"
    paramDef.parameterType = ParameterTypeDouble
    mStudyDefinition.StudyParameterDefinitions.Add paramDef

    Set paramDef = mCommonServiceConsumer.NewStudyParameterDefinition
    paramDef.name = PsParamIncrement
    paramDef.Description = "The amount by which the acceleration factor is increased " & _
                            "at each period"
    paramDef.parameterType = ParameterTypeDouble
    mStudyDefinition.StudyParameterDefinitions.Add paramDef

    Set paramDef = mCommonServiceConsumer.NewStudyParameterDefinition
    paramDef.name = PsParamMaxFactor
    paramDef.Description = "The maximum value of the acceleration factor that governs " & _
                            " how fast the parabolic stop rises or falls"
    paramDef.parameterType = ParameterTypeDouble
    mStudyDefinition.StudyParameterDefinitions.Add paramDef

End If

Set studyDefinition = mStudyDefinition
End Property

'================================================================================
' Helper Function
'================================================================================





