Attribute VB_Name = "GEMA"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GEMA"

Public Const EmaInputValue As String = "Input"

Public Const EMAParamPeriods As String = ParamPeriods
Public Const EMAParamSlopeThreshold As String = "Slope threshold"

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
    mDefaultParameters.SetParameterValue EMAParamPeriods, 21
    mDefaultParameters.SetParameterValue EMAParamSlopeThreshold, "0.0"
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

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = New StudyDefinition
    mStudyDefinition.name = EmaName
    mStudyDefinition.ShortName = EmaShortName
    mStudyDefinition.Description = "An exponential moving average"
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionUnderlying
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(EmaInputValue)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MovingAverageStudyValueName)
    valueDef.Description = "The moving average value"
    valueDef.IncludeInChart = True
    valueDef.IsDefault = True
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H1D9311)
    valueDef.ValueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(EMAParamPeriods)
    paramDef.Description = "The number of periods used to calculate the moving average"
    paramDef.ParameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(EMAParamSlopeThreshold)
    paramDef.Description = "The smallest slope value that is not to be considered flat"
    paramDef.ParameterType = ParameterTypeReal
    
End If

Set StudyDefinition = mStudyDefinition.Clone

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Helper Function
'@================================================================================





