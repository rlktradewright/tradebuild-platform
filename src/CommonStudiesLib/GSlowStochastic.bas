Attribute VB_Name = "GSlowStochastic"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GSlowStochastic"

Public Const SStochInputValue As String = "Input"

Public Const SStochParamKPeriods As String = "%K periods"
Public Const SStochParamKDPeriods As String = "%KD periods"
Public Const SStochParamDPeriods As String = "%D periods"

Public Const SStochValueK As String = "%K"
Public Const SStochValueD As String = "%D"

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
    mDefaultParameters.SetParameterValue SStochParamKPeriods, 5
    mDefaultParameters.SetParameterValue SStochParamKDPeriods, 3
    mDefaultParameters.SetParameterValue SStochParamDPeriods, 3
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
    mStudyDefinition.name = SStochName
    mStudyDefinition.ShortName = SStochShortName
    mStudyDefinition.Description = "Slow stochastic compares the latest price to the " & _
                                "recent trading range, and smoothes the result, " & _
                                "giving a Value called %K. " & _
                                "It also has another Value, called %D, which is " & _
                                "calculated by smoothing %K."
                                
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionCustom
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(SStochInputValue)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SStochValueK)
    valueDef.Description = "The slow stochastic Value (%K)"
    valueDef.IncludeInChart = True
    valueDef.IsDefault = True
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H3F5BC0, Layer:=LayerDataPoints + 1)
    valueDef.ValueType = ValueTypeReal
    valueDef.MinimumValue = -5#
    valueDef.MaximumValue = 105#
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SStochValueD)
    valueDef.Description = "The result of smoothing %K, also known as the signal line (%D)"
    valueDef.IncludeInChart = True
    valueDef.IsDefault = False
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HC6A566, Layer:=LayerDataPoints)
    valueDef.ValueType = ValueTypeReal
    valueDef.MinimumValue = -5#
    valueDef.MaximumValue = 105#
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SStochParamKPeriods)
    paramDef.Description = "The number of periods used to determine the recent " & _
                            "trading range"
    paramDef.ParameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SStochParamKDPeriods)
    paramDef.Description = "The number of periods of smoothing used in " & _
                            "calculating %K"
    paramDef.ParameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SStochParamDPeriods)
    paramDef.Description = "The number of periods used to smooth the %K Value " & _
                            "to obtain %D"
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







