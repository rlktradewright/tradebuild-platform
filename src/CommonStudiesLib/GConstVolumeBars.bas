Attribute VB_Name = "GConstVolumeBars"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Public Const ConstVolBarsInputOpenInterest As String = "Open interest"
Public Const ConstVolBarsInputOpenInterestUcase As String = "OPEN INTEREST"

Public Const ConstVolBarsInputPrice As String = "Price"
Public Const ConstVolBarsInputPriceUcase As String = "PRICE"

Public Const ConstVolBarsInputTotalVolume As String = "Total volume"
Public Const ConstVolBarsInputTotalVolumeUcase As String = "TOTAL VOLUME"

Public Const ConstVolBarsInputTickVolume As String = "Tick volume"
Public Const ConstVolBarsInputTickVolumeUcase As String = "TICK VOLUME"

Public Const ConstVolBarsParamVolPerBar As String = "Volume per bar"

Public Const ConstVolBarsValueBar As String = "Bar"

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
    mDefaultParameters.setParameterValue ConstVolBarsParamVolPerBar, 1000
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
    mStudyDefinition.name = ConstVolBarsName
    mStudyDefinition.needsBars = False
    mStudyDefinition.shortName = ConstVolBarsShortName
    mStudyDefinition.Description = "Constant volume bars " & _
                        "divide price movement into periods (bars) of equal volume. " & _
                        "For each period the open, high, low and close price values " & _
                        "are determined."
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstVolBarsInputPrice)
    inputDef.inputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstVolBarsInputTotalVolume)
    inputDef.inputType = InputTypeInteger
    inputDef.Description = "Accumulated volume"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstVolBarsInputTickVolume)
    inputDef.inputType = InputTypeInteger
    inputDef.Description = "Tick volume"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstVolBarsInputOpenInterest)
    inputDef.inputType = InputTypeInteger
    inputDef.Description = "Open interest"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(ConstVolBarsValueBar)
    valueDef.Description = "The constant volume bars"
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.IncludeInChart = True
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
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueOpenInterest)
    valueDef.Description = "Bar open interest"
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
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(ConstVolBarsParamVolPerBar)
    paramDef.Description = "The volume in each constant volume bar"
    paramDef.parameterType = ParameterTypeInteger

End If

Set StudyDefinition = mStudyDefinition.Clone
End Property

'@================================================================================
' Helper Function
'@================================================================================














