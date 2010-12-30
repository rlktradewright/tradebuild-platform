Attribute VB_Name = "GConstVolumeBars"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GConstVolumeBars"

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
Const ProcName As String = "defaultParameters"
On Error GoTo Err

Set mDefaultParameters = value.Clone

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get defaultParameters() As Parameters
Const ProcName As String = "defaultParameters"
On Error GoTo Err

If mDefaultParameters Is Nothing Then
    Set mDefaultParameters = New Parameters
    mDefaultParameters.SetParameterValue ConstVolBarsParamVolPerBar, 1000
End If

' now create a clone of the default parameters for the caller
Set defaultParameters = mDefaultParameters.Clone

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
    mStudyDefinition.name = ConstVolBarsName
    mStudyDefinition.NeedsBars = False
    mStudyDefinition.ShortName = ConstVolBarsShortName
    mStudyDefinition.Description = "Constant volume bars " & _
                        "divide price movement into periods (bars) of equal volume. " & _
                        "For each period the open, high, low and close price values " & _
                        "are determined."
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionCustom
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstVolBarsInputPrice)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstVolBarsInputTotalVolume)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Accumulated volume"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstVolBarsInputTickVolume)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Tick volume"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstVolBarsInputOpenInterest)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Open interest"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(ConstVolBarsValueBar)
    valueDef.Description = "The constant volume bars"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.IncludeInChart = True
    valueDef.ValueMode = ValueModeBar
    valueDef.ValueStyle = gCreateBarStyle
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueOpen)
    valueDef.Description = "Bar open value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H8000&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueHigh)
    valueDef.Description = "Bar high value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbBlue)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueLow)
    valueDef.Description = "Bar low value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbRed)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueClose)
    valueDef.Description = "Bar close value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.IsDefault = True
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H80&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueVolume)
    valueDef.Description = "Bar volume"
    valueDef.DefaultRegion = StudyValueDefaultRegionCustom
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(DisplayMode:=DataPointDisplayModeHistogram)
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
    valueDef.Description = "Bar H+L/2 value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HFF&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueHLC3)
    valueDef.Description = "Bar H+L+C/3 value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HFF00&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueOHLC4)
    valueDef.Description = "Bar O+H+L+C/4 value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HFF0000)
    valueDef.ValueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(ConstVolBarsParamVolPerBar)
    paramDef.Description = "The volume in each constant volume bar"
    paramDef.ParameterType = ParameterTypeInteger

End If

Set StudyDefinition = mStudyDefinition.Clone

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Helper Function
'@================================================================================














