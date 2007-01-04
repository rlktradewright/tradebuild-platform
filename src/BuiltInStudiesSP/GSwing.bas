Attribute VB_Name = "GSwing"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const SwingInputValue As String = "Input"

Public Const SwingParamIncludeImplicitSwingPoints As String = "Include implicit swing points"
Public Const SwingParamMinimumSwingTicks As String = "Minimum swing (ticks)"

Public Const SwingValueSwingHighPoint As String = "Swing high"
'Public Const SwingValueSwingHighType As String = "Swing high type"
Public Const SwingValueSwingLowPoint As String = "Swing low"
'Public Const SwingValueSwingLowType As String = "Swing low type"
Public Const SwingValueSwingPoint As String = "Swing point"

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
    mDefaultParameters.setParameterValue SwingParamMinimumSwingTicks, "10"
    mDefaultParameters.setParameterValue SwingParamIncludeImplicitSwingPoints, "Yes"
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
    mStudyDefinition.name = SwingName
    mStudyDefinition.shortName = SwingShortName
    mStudyDefinition.Description = "Determines the significant swing points of " & _
                                    "the underlying. For a move to be considered a swing, " & _
                                    "it must move at least the distance specified in the " & _
                                    "Minimum swing (ticks) parameter."
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionNone
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(SwingInputValue)
    inputDef.inputType = InputTypeReal
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SwingValueSwingPoint)
    valueDef.Description = "Swing point values"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeLine
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SwingValueSwingHighPoint)
    valueDef.Description = "Swing high point values"
    valueDef.isDefault = False
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeLine
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SwingValueSwingLowPoint)
    valueDef.Description = "Swing low point values"
    valueDef.isDefault = False
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeLine
    valueDef.valueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SwingParamMinimumSwingTicks)
    paramDef.Description = "The minimum number of ticks bar clearance from a high/low to " & _
                            "establish a new swing"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SwingParamIncludeImplicitSwingPoints)
    paramDef.Description = "Indicates whether to include implied swing points"
    paramDef.parameterType = ParameterTypeReal
    
End If

Set StudyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================





