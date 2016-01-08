Attribute VB_Name = "GDonchianChannels"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GDonchianChannels"

Public Const DoncInputPrice As String = "Price"

Public Const DoncParamPeriods As String = ParamPeriods

Public Const DoncValueLower As String = "Lower"
Public Const DoncValueUpper As String = "Upper"

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
    mDefaultParameters.SetParameterValue DoncParamPeriods, 13
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
    mStudyDefinition.name = DoncName
    mStudyDefinition.ShortName = DoncShortName
    mStudyDefinition.Description = "Donchian channels show the highest high and the " & _
                                    "lowest low during the specified preceeding number " & _
                                    "of periods"
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionUnderlying
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(DoncInputPrice)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(DoncValueLower)
    valueDef.Description = "The lower channel value"
    valueDef.IncludeInChart = True
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H8083FF, Layer:=LayerDataPoints + 5)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(DoncValueUpper)
    valueDef.Description = "The upper channel value"
    valueDef.IncludeInChart = True
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HD1EB43, Layer:=LayerDataPoints + 5)
    valueDef.ValueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(DoncParamPeriods)
    paramDef.Description = "The number of periods used to calculate the channel values"
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





