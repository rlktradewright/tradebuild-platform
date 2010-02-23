Attribute VB_Name = "GConstMomentumBars"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GConstMomentumBars"

Public Const ConstMomentumBarsInputOpenInterest As String = "Open interest"
Public Const ConstMomentumBarsInputOpenInterestUcase As String = "OPEN INTEREST"

Public Const ConstMomentumBarsInputPrice As String = "Price"
Public Const ConstMomentumBarsInputPriceUcase As String = "PRICE"

Public Const ConstMomentumBarsInputTotalVolume As String = "Total volume"
Public Const ConstMomentumBarsInputTotalVolumeUcase As String = "TOTAL VOLUME"

Public Const ConstMomentumBarsInputTickVolume As String = "Tick volume"
Public Const ConstMomentumBarsInputTickVolumeUcase As String = "TICK VOLUME"

Public Const ConstMomentumBarsParamTicksPerBar As String = "Ticks move per bar"

Public Const ConstMomentumBarsValueBar As String = "Bar"

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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get defaultParameters() As Parameters
Const ProcName As String = "defaultParameters"
On Error GoTo Err

If mDefaultParameters Is Nothing Then
    Set mDefaultParameters = New Parameters
    mDefaultParameters.SetParameterValue ConstMomentumBarsParamTicksPerBar, 10
End If

' now create a clone of the default parameters for the caller
Set defaultParameters = mDefaultParameters.Clone

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
    mStudyDefinition.name = ConstMomentumBarsName
    mStudyDefinition.NeedsBars = False
    mStudyDefinition.ShortName = ConstMomentumBarsShortName
    mStudyDefinition.Description = "Constant Momentum bars " & _
                        "divide price movement into periods (bars) of equal price movement. " & _
                        "For each period the open, high, low and close price values " & _
                        "are determined."
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstMomentumBarsInputPrice)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstMomentumBarsInputTotalVolume)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Accumulated volume"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstMomentumBarsInputTickVolume)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Tick volume"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(ConstMomentumBarsInputOpenInterest)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Open interest"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(ConstMomentumBarsValueBar)
    valueDef.Description = "The constant momentum bars"
    valueDef.DefaultRegion = DefaultRegionNone
    valueDef.IncludeInChart = True
    valueDef.ValueMode = ValueModeBar
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueOpen)
    valueDef.Description = "Bar open value"
    valueDef.DefaultRegion = DefaultRegionNone
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueHigh)
    valueDef.Description = "Bar high value"
    valueDef.DefaultRegion = DefaultRegionNone
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueLow)
    valueDef.Description = "Bar low value"
    valueDef.DefaultRegion = DefaultRegionNone
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueClose)
    valueDef.Description = "Bar close value"
    valueDef.DefaultRegion = DefaultRegionNone
    valueDef.IsDefault = True
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueVolume)
    valueDef.Description = "Bar volume"
    valueDef.DefaultRegion = DefaultRegionCustom
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueTickVolume)
    valueDef.Description = "Bar tick volume"
    valueDef.DefaultRegion = DefaultRegionCustom
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueOpenInterest)
    valueDef.Description = "Bar open interest"
    valueDef.DefaultRegion = DefaultRegionCustom
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueHL2)
    valueDef.Description = "Bar H+L/2 value"
    valueDef.DefaultRegion = DefaultRegionNone
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueHLC3)
    valueDef.Description = "Bar H+L+C/3 value"
    valueDef.DefaultRegion = DefaultRegionNone
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarValueOHLC4)
    valueDef.Description = "Bar O+H+L+C/4 value"
    valueDef.DefaultRegion = DefaultRegionNone
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(ConstMomentumBarsParamTicksPerBar)
    paramDef.Description = "The number of ticks movement in each constant momentum bar"
    paramDef.ParameterType = ParameterTypeInteger

End If

Set StudyDefinition = mStudyDefinition.Clone

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Helper Function
'@================================================================================
















