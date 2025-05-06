Attribute VB_Name = "GAccDist"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GAccDist"

Public Const AccDistInputOpen As String = "Open"
Public Const AccDistInputOpenUcase As String = "OPEN"

Public Const AccDistInputHigh As String = "High"
Public Const AccDistInputHighUcase As String = "HIGH"

Public Const AccDistInputLow As String = "Low"
Public Const AccDistInputLowUcase As String = "LOW"

Public Const AccDistInputClose As String = "Close"
Public Const AccDistInputCloseUcase As String = "CLOSE"

Public Const AccDistInputVolume As String = "Volume"
Public Const AccDistInputVolumeUcase As String = "VOLUME"

Public Const AccDistValueAccDist As String = "AccDist"

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

Const ProcName As String = "StudyDefinition"
On Error GoTo Err

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = New StudyDefinition
    mStudyDefinition.name = AccDistName
    mStudyDefinition.ShortName = AccDistShortName
    mStudyDefinition.Description = "Accumulation/Distribution tracks buying and selling " & _
                                "by combining price movements and volume"
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionCustom
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(AccDistInputOpen)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Open"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(AccDistInputHigh)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "High"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(AccDistInputLow)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Low"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(AccDistInputClose)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Close"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(AccDistInputVolume)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Volume"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(AccDistValueAccDist)
    valueDef.Description = "The Accumulation/Distribution value"
    valueDef.IncludeInChart = True
    valueDef.IsDefault = True
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbBlue, Layer:=LayerDataPoints)
    valueDef.ValueType = ValueTypeReal
    
End If

Set StudyDefinition = mStudyDefinition.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Helper Function
'@================================================================================





