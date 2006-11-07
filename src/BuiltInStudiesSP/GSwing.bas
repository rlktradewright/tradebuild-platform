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
    mDefaultParameters.setParameterValue SwingParamMinimumSwingTicks, "10"
    mDefaultParameters.setParameterValue SwingParamIncludeImplicitSwingPoints, "Yes"
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
    mStudyDefinition.name = SwingName
    mStudyDefinition.shortName = SwingShortName
    mStudyDefinition.Description = "Determines the significant swing points of " & _
                                    "the underlying. For a move to be considered a swing, " & _
                                    "it must move at least the distance specified in the " & _
                                    "Minimum swing (ticks) parameter."
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionPrice
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(SwingInputValue)
    inputDef.name = SwingInputValue
    inputDef.inputType = InputTypeDouble
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SwingValueSwingPoint)
    valueDef.name = SwingValueSwingPoint
    valueDef.Description = "Swing point values"
    valueDef.isDefault = True
    valueDef.multipleValuesPerBar = True    ' a bar can be both a swing point high and low
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SwingValueSwingHighPoint)
    valueDef.name = SwingValueSwingHighPoint
    valueDef.Description = "Swing high point values"
    valueDef.isDefault = False
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SwingValueSwingLowPoint)
    valueDef.name = SwingValueSwingLowPoint
    valueDef.Description = "Swing low point values"
    valueDef.isDefault = False
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SwingParamMinimumSwingTicks)
    paramDef.name = SwingParamMinimumSwingTicks
    paramDef.Description = "The minimum number of ticks bar clearance from a high/low to " & _
                            "establish a new swing"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SwingParamIncludeImplicitSwingPoints)
    paramDef.name = SwingParamIncludeImplicitSwingPoints
    paramDef.Description = "Indicates whether to include implied swing points"
    paramDef.parameterType = ParameterTypeDouble
    
End If

Set studyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================





