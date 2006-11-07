Attribute VB_Name = "GBars"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const BarsParamPeriodLength As String = "Period length"
Public Const BarsParamPeriodUnits As String = "Period units"

Public Const BarsValueClose As String = "Close"
Public Const BarsValueSize As String = "Size"
Public Const BarsValueTickVolume As String = "Tick volume"
Public Const BarsValueVolume As String = "Volume"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Global object references
'================================================================================

'================================================================================
' External function declarations
'================================================================================

'================================================================================
' Variables
'================================================================================

'================================================================================
' Procedures
'================================================================================

Public Property Get studyDefinition() As studyDefinition
Dim paramDef As StudyParameterDefinition
Dim valueDef As StudyValueDefinition

Set studyDefinition = New studyDefinition

studyDefinition.Description = "Formats the price stream into Open/High/Low/Close bars of an appropriate length."
studyDefinition.name = "Bars"
studyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionPrice

Set paramDef = studyDefinition.studyParameterDefinitions.add(BarsParamPeriodLength)
paramDef.name = BarsParamPeriodLength
paramDef.Description = "Length of one bar"
paramDef.parameterType = StudyParameterTypes.ParameterTypeInteger

Set paramDef = studyDefinition.studyParameterDefinitions.add(BarsParamPeriodUnits)
paramDef.name = BarsParamPeriodUnits
paramDef.Description = "The units in which Period length is measured."
paramDef.parameterType = StudyParameterTypes.ParameterTypeString

Set valueDef = studyDefinition.studyValueDefinitions.add(BarsValueClose)
valueDef.name = BarsValueClose
valueDef.Description = "The latest underlying value"
valueDef.isDefault = True
valueDef.valueType = StudyValueTypes.ValueTypeDouble

'Set valueDef = studyDefinition.studyValueDefinitions.add(BarsValueSize)
'valueDef.name = BarsValueSize
'valueDef.Description = "The size associated with the latest underlying value (where relevant)"
'valueDef.valueType = StudyValueTypes.ValueTypeInteger

Set valueDef = studyDefinition.studyValueDefinitions.add(BarsValueTickVolume)
valueDef.name = BarsValueTickVolume
valueDef.Description = "The number of ticks in the current bar for the underlying value"
valueDef.valueType = StudyValueTypes.ValueTypeInteger

Set valueDef = studyDefinition.studyValueDefinitions.add(BarsValueVolume)
valueDef.name = BarsValueVolume
valueDef.Description = "The cumulative size associated with the latest underlying value (where relevant)"
valueDef.valueType = StudyValueTypes.ValueTypeInteger

End Sub

'================================================================================
' Helper Function
'================================================================================





