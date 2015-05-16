Attribute VB_Name = "GATR"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GATR"

Public Const AtrInputPrice As String = "Price"

Public Const AtrParamMAType As String = ParamMovingAverageType
Public Const AtrParamPeriods As String = ParamPeriods

Public Const AtrValueATR As String = "ATR"

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
    mDefaultParameters.SetParameterValue AtrParamPeriods, 27
    mDefaultParameters.SetParameterValue AtrParamMAType, EmaShortName
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
    mStudyDefinition.name = AtrName
    mStudyDefinition.ShortName = AtrShortName
    mStudyDefinition.Description = "Average True Range " & _
                        "calculates the moving average of the 'true ranges' of bars " & _
                        "over the specified number of periods. " & _
                        "The true range of a bar is calculated by substituting the " & _
                        "previous close for the bar's low (if lower), or for the high (if higher)."
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionCustom
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(AtrInputPrice)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(AtrValueATR)
    valueDef.Description = "The Average True Range Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.IncludeInChart = True
    valueDef.IsDefault = True
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbGreen)
    valueDef.ValueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(AtrParamPeriods)
    paramDef.Description = "The number of periods used to calculate the Average True Range"
    paramDef.ParameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(AtrParamMAType)
    paramDef.Description = "The type of moving average to be used"
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









