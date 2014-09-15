Attribute VB_Name = "GConstTickVolumeBars"
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GConstTickVolumeBars"

Public Const ConstTickVolumeBarsInputOpenInterest As String = "Open interest"
Public Const ConstTickVolumeBarsInputOpenInterestUcase As String = "OPEN INTEREST"

Public Const ConstTickVolumeBarsInputPrice As String = "Price"
Public Const ConstTickVolumeBarsInputPriceUcase As String = "PRICE"

Public Const ConstTickVolumeBarsInputTotalVolume As String = "Total volume"
Public Const ConstTickVolumeBarsInputTotalVolumeUcase As String = "TOTAL VOLUME"

Public Const ConstTickVolumeBarsInputTickVolume As String = "Tick volume"
Public Const ConstTickVolumeBarsInputTickVolumeUcase As String = "TICK VOLUME"

Public Const ConstTickVolumeBarsValueBar As String = "Bar"

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


Public Property Let defaultParameters(ByVal Value As Parameters)
' create a clone of the default parameters supplied by the caller
Const ProcName As String = "defaultParameters"
On Error GoTo Err

Set mDefaultParameters = Value.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get defaultParameters() As Parameters
Const ProcName As String = "defaultParameters"
On Error GoTo Err

If mDefaultParameters Is Nothing Then
    Set mDefaultParameters = New Parameters
    mDefaultParameters.SetParameterValue ConstTickVolumeBarsParamTicksPerBar, 100
End If

' now create a clone of the default parameters for the caller
Set defaultParameters = mDefaultParameters.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StudyDefinition() As StudyDefinition
Dim inputDef As StudyInputDefinition
Dim valueDef As StudyValueDefinition
Dim paramDef As StudyParameterDefinition
Const ProcName As String = "StudyDefinition"
On Error GoTo Err

ReDim ar(6) As Variant

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = New StudyDefinition
    mStudyDefinition.name = ConstTickVolumeBarsStudyName
    mStudyDefinition.NeedsBars = False
    mStudyDefinition.ShortName = ConstTickVolumeBarsStudyShortName
    mStudyDefinition.Description = "Constant Tick Volume bars " & _
                        "divide price movement into periods (bars) with equal numbers of trades. " & _
                        "For each period the open, high, low and close price values " & _
                        "are determined."
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionCustom
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstTickVolumeBarsInputPrice)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstTickVolumeBarsInputTotalVolume)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Accumulated volume"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstTickVolumeBarsInputTickVolume)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Tick volume"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstTickVolumeBarsInputOpenInterest)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Open interest"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(ConstTickVolumeBarsValueBar)
    valueDef.Description = "The constant tick volume bars"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.IncludeInChart = True
    valueDef.ValueMode = ValueModeBar
    valueDef.ValueStyle = gCreateBarStyle
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueOpen)
    valueDef.Description = "Bar open Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H8000&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueHigh)
    valueDef.Description = "Bar high Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbBlue)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueLow)
    valueDef.Description = "Bar low Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbRed)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueClose)
    valueDef.Description = "Bar close Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.IsDefault = True
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H80&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueVolume)
    valueDef.Description = "Bar volume"
    valueDef.DefaultRegion = StudyValueDefaultRegionCustom
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(Color:=&H808080, DisplayMode:=DataPointDisplayModeHistogram)
    valueDef.ValueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueTickVolume)
    valueDef.Description = "Bar tick volume"
    valueDef.DefaultRegion = StudyValueDefaultRegionCustom
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(Color:=&H800000, DisplayMode:=DataPointDisplayModeHistogram)
    valueDef.ValueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueOpenInterest)
    valueDef.Description = "Bar open interest"
    valueDef.DefaultRegion = StudyValueDefaultRegionCustom
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(Color:=&H80&, DisplayMode:=DataPointDisplayModeHistogram)
    valueDef.ValueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueHL2)
    valueDef.Description = "Bar H+L/2 Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HFF&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueHLC3)
    valueDef.Description = "Bar H+L+C/3 Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HFF00&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueOHLC4)
    valueDef.Description = "Bar O+H+L+C/4 Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HFF0000)
    valueDef.ValueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(ConstTickVolumeBarsParamTicksPerBar)
    paramDef.Description = "The number of trades in each constant tick volume bar"
    paramDef.ParameterType = ParameterTypeInteger

End If

Set StudyDefinition = mStudyDefinition.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Helper Function
'@================================================================================


















