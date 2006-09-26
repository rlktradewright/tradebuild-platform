Attribute VB_Name = "GDonchianChannels"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const DoncParamPeriods As String = ParamPeriods

Public Const DoncValueLower As String = "Lower"
Public Const DoncValueUpper As String = "Upper"

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
    mDefaultParameters.setParameterValue DoncParamPeriods, 13
End If

' now create a clone of the default parameters for the caller
Set defaultParameters = mDefaultParameters.Clone
End Property

Public Property Get studyDefinition() As TradeBuildSP.IStudyDefinition
Dim valueDef As IStudyValueDefinition
Dim paramDef As IStudyParameterDefinition

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = mCommonServiceConsumer.NewStudyDefinition
    mStudyDefinition.name = DoncName
    mStudyDefinition.Description = "Donchian channels show the highest high and the " & _
                                    "lowest low during the specified preceeding number " & _
                                    "of periods"
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionPrice
    
    Set valueDef = mCommonServiceConsumer.NewStudyValueDefinition
    valueDef.name = DoncValueLower
    valueDef.Description = "The lower channel value"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valuetype = ValueTypeDouble
    mStudyDefinition.StudyValueDefinitions.Add valueDef
    
    Set valueDef = mCommonServiceConsumer.NewStudyValueDefinition
    valueDef.name = DoncValueUpper
    valueDef.Description = "The upper channel value"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valuetype = ValueTypeDouble
    mStudyDefinition.StudyValueDefinitions.Add valueDef
    
    Set paramDef = mCommonServiceConsumer.NewStudyParameterDefinition
    paramDef.name = DoncParamPeriods
    paramDef.Description = "The number of periods used to calculate the channel values"
    paramDef.parameterType = ParameterTypeInteger
    mStudyDefinition.StudyParameterDefinitions.Add paramDef

End If

Set studyDefinition = mStudyDefinition
End Property

'================================================================================
' Helper Function
'================================================================================





