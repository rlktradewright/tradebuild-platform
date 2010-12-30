Attribute VB_Name = "GParabolicStop"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GParabolicStop"

Public Const PsInputPrice As String = "Price"

Public Const PsParamStartFactor As String = "Start factor"
Public Const PsParamIncrement As String = "Increment"
Public Const PsParamMaxFactor As String = "Max factor"

Public Const PsValuePs As String = "PS"

Public Const PsDefaultStartFactor As Double = 0.02
Public Const PsDefaultIncrement As Double = 0.02
Public Const PsDefaultMaxFactor As Double = 0.2

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
    mDefaultParameters.SetParameterValue PsParamStartFactor, PsDefaultStartFactor
    mDefaultParameters.SetParameterValue PsParamIncrement, PsDefaultIncrement
    mDefaultParameters.SetParameterValue PsParamMaxFactor, PsDefaultMaxFactor
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
    mStudyDefinition.name = PsName
    mStudyDefinition.ShortName = PsShortName
    mStudyDefinition.Description = "Parabolic Stop calculates a value that can be used " & _
                                    "as a stop loss for trades. When the market is " & _
                                    "rising, the value increases with each period. When " & _
                                    "the market is falling, the value decreases with " & _
                                    "each period."
                                    
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionUnderlying
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(PsInputPrice)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(PsValuePs)
    valueDef.Description = "The parabolic stop value"
    valueDef.IncludeInChart = True
    valueDef.IsDefault = True
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H618A55, DataPointDisplayModePoint, Linethickness:=5, PointStyle:=PointSquare)
    valueDef.ValueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(PsParamStartFactor)
    paramDef.Description = "The initial value of the acceleration factor that governs " & _
                            "the increase in the speed with which the parabolic stop " & _
                            "rises or falls"
    paramDef.ParameterType = ParameterTypeReal

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(PsParamIncrement)
    paramDef.Description = "The amount by which the acceleration factor is increased " & _
                            "at each period"
    paramDef.ParameterType = ParameterTypeReal

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(PsParamMaxFactor)
    paramDef.Description = "The maximum value of the acceleration factor that governs " & _
                            " how fast the parabolic stop rises or falls"
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





