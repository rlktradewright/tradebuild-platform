Attribute VB_Name = "GForceIndex"
Option Explicit



'@================================================================================
' Constants
'@================================================================================

Public Const FiInputPrice As String = "Price"
Public Const FiInputPriceUcase As String = "PRICE"

Public Const FiInputVolume As String = "Volume"
Public Const FiInputVolumeUcase As String = "VOLUME"

Public Const FiParamShortPeriods As String = "Short EMA periods"
Public Const FiParamLongPeriods As String = "Long EMA periods"

Public Const FiValueForceIndex As String = "FI"
Public Const FiValueForceIndexShort As String = "FI (short)"
Public Const FiValueForceIndexLong As String = "FI (long)"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Global object references
'@================================================================================

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Variables
'@================================================================================

Private mDefaultParameters As Parameters
Private mStudyDefinition As StudyDefinition

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
    mDefaultParameters.setParameterValue FiParamShortPeriods, 2
    mDefaultParameters.setParameterValue FiParamLongPeriods, 13
End If

' now return a clone of the default parameters for the caller, to
' prevent the caller changing ours
Set defaultParameters = mDefaultParameters.Clone
End Property

Public Property Get StudyDefinition() As StudyDefinition
Dim inputDef As StudyInputDefinition
Dim valueDef As StudyValueDefinition
Dim paramDef As StudyParameterDefinition

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = New StudyDefinition
    mStudyDefinition.name = FIName
    mStudyDefinition.shortName = FIShortName
    mStudyDefinition.Description = "Force Index combines price and volume to " & _
                                    "give a measure of bullish or bearish " & _
                                    "force in the market"
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(FiInputPrice)
    inputDef.inputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(FiInputVolume)
    inputDef.inputType = InputTypeInteger
    inputDef.Description = "Volume"
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(FiParamShortPeriods)
    paramDef.Description = "The number of periods used for the short EMA"
    paramDef.parameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(FiParamLongPeriods)
    paramDef.Description = "The number of periods used for the long EMA"
    paramDef.parameterType = ParameterTypeInteger

    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(FiValueForceIndex)
    valueDef.Description = "The Force Index value"
    valueDef.IncludeInChart = True
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueStyle = gCreateDataPointStyle
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(FiValueForceIndexShort)
    valueDef.Description = "The Force Index value smoothed by the short EMA"
    valueDef.IncludeInChart = True
    valueDef.isDefault = False
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueStyle = gCreateDataPointStyle(vbRed)
    valueDef.valueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(FiValueForceIndexLong)
    valueDef.Description = "The Force Index value smoothed by the long EMA"
    valueDef.IncludeInChart = True
    valueDef.isDefault = False
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueStyle = gCreateDataPointStyle(vbBlue)
    valueDef.valueType = ValueTypeReal
    
    
End If

' return a clone to prevent the application changing our definition
Set StudyDefinition = mStudyDefinition.Clone
End Property

'@================================================================================
' Helper Function
'@================================================================================





