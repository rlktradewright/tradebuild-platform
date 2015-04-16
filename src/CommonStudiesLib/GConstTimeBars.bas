Attribute VB_Name = "GConstTimeBars"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GConstTimeBars"

Public Const ConstTimeBarsInputOpenInterest As String = "Open interest"
Public Const ConstTimeBarsInputOpenInterestUcase As String = "OPEN INTEREST"

Public Const ConstTimeBarsInputPrice As String = "Price"
Public Const ConstTimeBarsInputPriceUcase As String = "PRICE"

Public Const ConstTimeBarsInputTotalVolume As String = "Total volume"
Public Const ConstTimeBarsInputTotalVolumeUcase As String = "TOTAL VOLUME"

Public Const ConstTimeBarsInputTickVolume As String = "Tick volume"
Public Const ConstTimeBarsInputTickVolumeUcase As String = "TICK VOLUME"

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
    mDefaultParameters.SetParameterValue ConstTimeBarsParamBarLength, 5
    mDefaultParameters.SetParameterValue ConstTimeBarsParamTimeUnits, _
                                        TimePeriodUnitsToString(TimePeriodMinute)
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
    mStudyDefinition.name = ConstTimeBarsStudyName
    mStudyDefinition.NeedsBars = False
    mStudyDefinition.ShortName = ConstTimeBarsStudyShortName
    mStudyDefinition.Description = "Constant time bars " & _
                        "divide price movement into periods (bars) of equal time. " & _
                        "For each period the open, high, low and close price values " & _
                        "are determined."
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionCustom
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstTimeBarsInputPrice)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstTimeBarsInputTotalVolume)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Accumulated volume"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstTimeBarsInputTickVolume)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Tick volume"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstTimeBarsInputOpenInterest)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Open interest"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(ConstTimeBarsValueBar)
    valueDef.Description = "The constant time bars"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.IncludeInChart = True
    valueDef.ValueMode = ValueModeBar
    valueDef.ValueStyle = gCreateBarStyle
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueOpen)
    valueDef.Description = "Bar open Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H8000&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueHigh)
    valueDef.Description = "Bar high Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbBlue)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueLow)
    valueDef.Description = "Bar low Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbRed)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueClose)
    valueDef.Description = "Bar close Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.IsDefault = True
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H80&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueVolume)
    valueDef.Description = "Bar volume"
    valueDef.DefaultRegion = StudyValueDefaultRegionCustom
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(Color:=&H808080, DisplayMode:=DataPointDisplayModeHistogram)
    valueDef.ValueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueTickVolume)
    valueDef.Description = "Bar tick volume"
    valueDef.DefaultRegion = StudyValueDefaultRegionCustom
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(Color:=&H800000, DisplayMode:=DataPointDisplayModeHistogram)
    valueDef.ValueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueOpenInterest)
    valueDef.Description = "Bar open interest"
    valueDef.DefaultRegion = StudyValueDefaultRegionCustom
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(Color:=&H80&, DisplayMode:=DataPointDisplayModeHistogram)
    valueDef.ValueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueHL2)
    valueDef.Description = "Bar H+L/2 Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HFF&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueHLC3)
    valueDef.Description = "Bar H+L+C/3 Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HFF00&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueOHLC4)
    valueDef.Description = "Bar O+H+L+C/4 Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HFF0000)
    valueDef.ValueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(ConstTimeBarsParamBarLength)
    paramDef.Description = "The number of time units in each constant time bar"
    paramDef.ParameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(ConstTimeBarsParamTimeUnits)
    paramDef.Description = "The time units that the constant time bars are measured in"
    paramDef.ParameterType = ParameterTypeString
    ar(0) = TimePeriodUnitsToString(TimePeriodSecond)
    ar(1) = TimePeriodUnitsToString(TimePeriodMinute)
    ar(2) = TimePeriodUnitsToString(TimePeriodHour)
    ar(3) = TimePeriodUnitsToString(TimePeriodDay)
    ar(4) = TimePeriodUnitsToString(TimePeriodWeek)
    ar(5) = TimePeriodUnitsToString(TimePeriodMonth)
    ar(6) = TimePeriodUnitsToString(TimePeriodYear)
    paramDef.PermittedValues = ar
    
End If

Set StudyDefinition = mStudyDefinition.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Helper Function
'@================================================================================












