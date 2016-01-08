Attribute VB_Name = "GForceIndex"
Option Explicit



'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GForceIndex"

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
    mDefaultParameters.SetParameterValue FiParamShortPeriods, 2
    mDefaultParameters.SetParameterValue FiParamLongPeriods, 13
End If

' now return a clone of the default parameters for the caller, to
' prevent the caller changing ours
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
    mStudyDefinition.name = FIName
    mStudyDefinition.ShortName = FIShortName
    mStudyDefinition.Description = "Force Index combines price and volume to " & _
                                    "give a measure of bullish or bearish " & _
                                    "force in the market"
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionCustom
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(FiInputPrice)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(FiInputVolume)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Volume"
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(FiParamShortPeriods)
    paramDef.Description = "The number of periods used for the short EMA"
    paramDef.ParameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(FiParamLongPeriods)
    paramDef.Description = "The number of periods used for the long EMA"
    paramDef.ParameterType = ParameterTypeInteger

    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(FiValueForceIndex)
    valueDef.Description = "The Force Index value"
    valueDef.IncludeInChart = True
    valueDef.IsDefault = True
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueStyle = gCreateDataPointStyle(&HACD2B1, Layer:=LayerDataPoints + 2)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(FiValueForceIndexShort)
    valueDef.Description = "The Force Index Value smoothed by the short EMA"
    valueDef.IncludeInChart = True
    valueDef.IsDefault = False
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueStyle = gCreateDataPointStyle(&H3264CD, Layer:=LayerDataPoints + 1)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(FiValueForceIndexLong)
    valueDef.Description = "The Force Index Value smoothed by the long EMA"
    valueDef.IncludeInChart = True
    valueDef.IsDefault = False
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueStyle = gCreateDataPointStyle(&HDCA58D, Layer:=LayerDataPoints)
    valueDef.ValueType = ValueTypeReal
    
    
End If

' return a clone to prevent the application changing our definition
Set StudyDefinition = mStudyDefinition.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Helper Function
'@================================================================================





