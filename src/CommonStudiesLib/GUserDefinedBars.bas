Attribute VB_Name = "GUserDefinedBars"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GUserDefinedBars"

Public Const UserDefinedBarsInputValue  As String = "Value"
Public Const UserDefinedBarsInputValueUCase  As String = "VALUE"

Public Const UserDefinedBarsInputBarNumber  As String = "Bar number"
Public Const UserDefinedBarsInputBarNumberUCase  As String = "BAR NUMBER"

Public Const UserDefinedBarsValueBar As String = "Bar"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Global object references
'@================================================================================


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
Const ProcName As String = "defaultParameters"
On Error GoTo Err

Assert False, "Study has no parameters", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get defaultParameters() As Parameters
Const ProcName As String = "defaultParameters"
On Error GoTo Err

Set defaultParameters = New Parameters

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
    mStudyDefinition.name = ConstTimeBarsName
    mStudyDefinition.NeedsBars = False
    mStudyDefinition.ShortName = ConstTimeBarsShortName
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














