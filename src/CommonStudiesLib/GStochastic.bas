Attribute VB_Name = "GStochastic"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GStochastic"

Public Const StochInputValue As String = "Input"

Public Const StochParamKPeriods As String = "%K periods"
Public Const StochParamDPeriods As String = "%D periods"

Public Const StochValueK As String = "%K"
Public Const StochValueD As String = "%D"

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
    mDefaultParameters.SetParameterValue StochParamKPeriods, 5
    mDefaultParameters.SetParameterValue StochParamDPeriods, 3
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
    mStudyDefinition.name = StochName
    mStudyDefinition.ShortName = StochShortName
    mStudyDefinition.Description = "Stochastic compares the latest price to the " & _
                                "recent trading range, giving a Value called %K. " & _
                                "It also has another Value, called %D, which is " & _
                                "calculated by smoothing %K."
                                
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionCustom
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(StochInputValue)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Input Value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(StochValueK)
    valueDef.Description = "The stochastic Value (%K)"
    valueDef.IncludeInChart = True
    valueDef.IsDefault = True
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbBlue)
    valueDef.ValueType = ValueTypeReal
    valueDef.MinimumValue = -5#
    valueDef.MaximumValue = 105#
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(StochValueD)
    valueDef.Description = "The result of smoothing %K, also known as the signal line (%D)"
    valueDef.IncludeInChart = True
    valueDef.IsDefault = False
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbRed)
    valueDef.ValueType = ValueTypeReal
    valueDef.MinimumValue = -5#
    valueDef.MaximumValue = 105#
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(StochParamKPeriods)
    paramDef.Description = "The number of periods used to determine the recent " & _
                            "trading range"
    paramDef.ParameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(StochParamDPeriods)
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





