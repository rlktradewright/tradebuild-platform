Attribute VB_Name = "GDonchianChannels"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const DoncInputPrice As String = "Price"

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
    mDefaultParameters.setParameterValue DoncParamPeriods, 13
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
    mStudyDefinition.name = DoncName
    mStudyDefinition.shortName = DoncShortName
    mStudyDefinition.Description = "Donchian channels show the highest high and the " & _
                                    "lowest low during the specified preceeding number " & _
                                    "of periods"
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionNone
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(DoncInputPrice)
    inputDef.inputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(DoncValueLower)
    valueDef.Description = "The lower channel value"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(DoncValueUpper)
    valueDef.Description = "The upper channel value"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(DoncParamPeriods)
    paramDef.Description = "The number of periods used to calculate the channel values"
    paramDef.parameterType = ParameterTypeInteger

End If

Set StudyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================





