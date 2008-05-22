Attribute VB_Name = "GConstTimeBars"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Public Const ConstTimeBarsInputOpenInterest As String = "Open interest"
Public Const ConstTimeBarsInputPrice As String = "Price"
Public Const ConstTimeBarsInputTotalVolume As String = "Total volume"
Public Const ConstTimeBarsInputTickVolume As String = "Tick volume"

Public Const ConstTimeBarsParamBarLength As String = "Bar length"
Public Const ConstTimeBarsParamTimeUnits As String = "Time units"

Public Const ConstTimeBarsValueBar As String = "Bar"

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
    mDefaultParameters.setParameterValue ConstTimeBarsParamBarLength, 5
    mDefaultParameters.setParameterValue ConstTimeBarsParamTimeUnits, _
                                        TimePeriodUnitsToString(TimePeriodMinute)
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
    mStudyDefinition.name = ConstTimeBarsName
    mStudyDefinition.needsBars = False
    mStudyDefinition.shortName = ConstTimeBarsShortName
    mStudyDefinition.Description = "Constant time bars " & _
                        "divide price movement into periods (bars) of equal time. " & _
                        "For each period the open, high, low and close price values " & _
                        "are determined."
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstTimeBarsInputPrice)
    inputDef.inputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstTimeBarsInputTotalVolume)
    inputDef.inputType = InputTypeInteger
    inputDef.Description = "Accumulated volume"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstTimeBarsInputTickVolume)
    inputDef.inputType = InputTypeInteger
    inputDef.Description = "Tick volume"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstTimeBarsInputOpenInterest)
    inputDef.inputType = InputTypeInteger
    inputDef.Description = "Open interest"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(ConstTimeBarsValueBar)
    valueDef.Description = "The constant time bars"
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
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(ConstTimeBarsParamBarLength)
    paramDef.Description = "The number of time units in each constant time bar"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(ConstTimeBarsParamTimeUnits)
    paramDef.Description = "The time units that the constant time bars are measured in"
    paramDef.parameterType = ParameterTypeString
    ar(0) = TimePeriodUnitsToString(TimePeriodSecond)
    ar(1) = TimePeriodUnitsToString(TimePeriodMinute)
    ar(2) = TimePeriodUnitsToString(TimePeriodHour)
    ar(3) = TimePeriodUnitsToString(TimePeriodDay)
    ar(4) = TimePeriodUnitsToString(TimePeriodWeek)
    ar(5) = TimePeriodUnitsToString(TimePeriodMonth)
    ar(6) = TimePeriodUnitsToString(TimePeriodYear)
    paramDef.permittedValues = ar
    
End If

Set StudyDefinition = mStudyDefinition.Clone
End Property

'@================================================================================
' Helper Function
'@================================================================================












