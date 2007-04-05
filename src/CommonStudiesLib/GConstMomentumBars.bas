Attribute VB_Name = "Module1"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const ConstMomentumBarsInputPrice As String = "Price"
Public Const ConstMomentumBarsInputTotalVolume As String = "Total volume"
Public Const ConstMomentumBarsInputTickVolume As String = "Tick volume"

Public Const ConstMomentumBarsParamTicksPerBar As String = "Ticks move per bar"

Public Const ConstMomentumBarsValueBar As String = "Bar"
Public Const ConstMomentumBarsValueTotalVolume As String = "Total volume"

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
    mDefaultParameters.setParameterValue ConstMomentumBarsParamTicksPerBar, 10
End If

' now create a clone of the default parameters for the caller
Set defaultParameters = mDefaultParameters.Clone
End Property

Public Property Get StudyDefinition() As StudyDefinition
Dim inputDef As StudyInputDefinition
Dim valueDef As StudyValueDefinition
Dim paramDef As StudyParameterDefinition
ReDim ar(6) As Variant

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = New StudyDefinition
    mStudyDefinition.name = ConstMomentumBarsName
    mStudyDefinition.needsBars = False
    mStudyDefinition.shortName = ConstMomentumBarsShortName
    mStudyDefinition.Description = "Constant Momentum bars " & _
                        "divide price movement into periods (bars) of equal price movement. " & _
                        "For each period the open, high, low and close price values " & _
                        "are determined."
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstMomentumBarsInputPrice)
    inputDef.inputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstMomentumBarsInputTotalVolume)
    inputDef.inputType = InputTypeInteger
    inputDef.Description = "Accumulated volume"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstMomentumBarsInputTickVolume)
    inputDef.inputType = InputTypeInteger
    inputDef.Description = "Tick volume"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(ConstMomentumBarsValueBar)
    valueDef.Description = "The constant momentum bars"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeBar
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueOpen)
    valueDef.Description = "Bar open value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueHigh)
    valueDef.Description = "Bar high value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueLow)
    valueDef.Description = "Bar low value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueClose)
    valueDef.Description = "Bar close value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.isDefault = True
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueVolume)
    valueDef.Description = "Bar volume"
    valueDef.defaultRegion = DefaultRegionCustom
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(ConstMomentumBarsValueTotalVolume)
    valueDef.Description = "Accumulated volume"
    valueDef.defaultRegion = DefaultRegionCustom
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueTickVolume)
    valueDef.Description = "Bar tick volume"
    valueDef.defaultRegion = DefaultRegionCustom
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueHL2)
    valueDef.Description = "Bar H+L/2 value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueHLC3)
    valueDef.Description = "Bar H+L+C/3 value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueOHLC4)
    valueDef.Description = "Bar O+H+L+C/4 value"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(ConstMomentumBarsParamTicksPerBar)
    paramDef.Description = "The number of ticks ovement in each constant momentum bar"
    paramDef.parameterType = ParameterTypeInteger

End If

Set StudyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================
















