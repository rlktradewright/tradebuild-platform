Attribute VB_Name = "GMACD"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GMACD"

Public Const MACDInputValue As String = "Input"

Public Const MACDParamLongPeriods As String = "Long periods"
Public Const MACDParamMAType As String = ParamMovingAverageType
Public Const MACDParamShortPeriods As String = "Short periods"
Public Const MACDParamSmoothingPeriods As String = "Smoothing periods"

Public Const MACDValueMACD As String = "MACD"
Public Const MACDValueMACDHist As String = "MACD hist"
Public Const MACDValueMACDLowerBalance As String = "MACD lower balance"
Public Const MACDValueMACDSignal As String = "MACD signal"
Public Const MACDValueStrength As String = "Strength"
Public Const MACDValueStrengthCount As String = "Strength count"
Public Const MACDValueMACDUpperBalance As String = "MACD upper balance"

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
    mDefaultParameters.SetParameterValue MACDParamShortPeriods, 12
    mDefaultParameters.SetParameterValue MACDParamLongPeriods, 26
    mDefaultParameters.SetParameterValue MACDParamSmoothingPeriods, 9
    mDefaultParameters.SetParameterValue MACDParamMAType, EmaShortName
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

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = New StudyDefinition
    mStudyDefinition.name = MacdName
    mStudyDefinition.ShortName = MacdShortName
    mStudyDefinition.Description = "MACD (Moving Average Convergence/Divergence) " & _
                        "calculates the difference between two moving averages of " & _
                        "different periods. A further moving average is applied " & _
                        "to this difference to give a signal line. Finally the " & _
                        "difference between the MACD and the signal Value gives " & _
                        "another indicator that is usually plotted as a " & _
                        "histogram."
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionCustom
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(MACDInputValue)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MACDValueMACD)
    valueDef.Description = "The MACD value"
    valueDef.IsDefault = True
    valueDef.IncludeInChart = True
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H5C6FED, Layer:=LayerDataPoints + 3)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MACDValueMACDSignal)
    valueDef.Description = "The MACD signal value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.IncludeInChart = True
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HDEB15F, Layer:=LayerDataPoints + 4)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MACDValueMACDHist)
    valueDef.Description = "The MACD histogram value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.IncludeInChart = True
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(DisplayMode:=DataPointDisplayModeHistogram, DownColor:=&H43FC2, Layer:=LayerDataPoints + 2, UpColor:=&H1D9311)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MACDValueStrengthCount)
    valueDef.Description = "The number of consecutive bars for which the current " & _
                            "strength Value has not changed"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H808080, DisplayMode:=DataPointDisplayModeHistogram, Layer:=LayerDataPoints + 1)
    valueDef.ValueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MACDValueStrength)
    valueDef.Description = "An indication of the strength of the current move"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(DisplayMode:=DataPointDisplayModeHistogram, DownColor:=&H43FC2, Layer:=LayerDataPoints, UpColor:=&H1D9311)
    valueDef.ValueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MACDValueMACDUpperBalance)
    valueDef.Description = "The price above which is confirmed strength"
    valueDef.DefaultRegion = StudyValueDefaultRegionUnderlying
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HC47E44, Layer:=LayerDataPoints + 20)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MACDValueMACDLowerBalance)
    valueDef.Description = "The price below which is confirmed weakness"
    valueDef.DefaultRegion = StudyValueDefaultRegionUnderlying
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H4144C7, Layer:=LayerDataPoints + 20)
    valueDef.ValueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(MACDParamShortPeriods)
    paramDef.Description = "The number of periods in the shorter moving average"
    paramDef.ParameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(MACDParamLongPeriods)
    paramDef.Description = "The number of periods in the longer moving average"
    paramDef.ParameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(MACDParamSmoothingPeriods)
    paramDef.Description = "The number of periods for smoothing the MACD to " & _
                            "produce the MACD signal value"
    paramDef.ParameterType = ParameterTypeInteger
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(MACDParamMAType)
    paramDef.Description = "The type of moving averages to be used"
    paramDef.ParameterType = ParameterTypeString
    paramDef.PermittedValues = gMaTypes
    
End If

Set StudyDefinition = mStudyDefinition.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Helper Function
'@================================================================================





