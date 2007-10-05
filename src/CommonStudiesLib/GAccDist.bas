Attribute VB_Name = "GAccDist"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Public Const AccDistInputPrice As String = "Price"
Public Const AccDistInputVolume As String = "Volume"

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


Public Property Let defaultParameters(ByVal value As Parameters)
' create a clone of the default parameters supplied by the caller
Set mDefaultParameters = value.Clone
End Property

Public Property Get defaultParameters() As Parameters
If mDefaultParameters Is Nothing Then
    Set mDefaultParameters = New Parameters
End If

' now create a clone of the default parameters for the caller
Set defaultParameters = mDefaultParameters.Clone
End Property

Public Property Get StudyDefinition() As StudyDefinition
Dim inputDef As StudyInputDefinition
Dim valueDef As StudyValueDefinition

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = New StudyDefinition
    mStudyDefinition.name = AccDistName
    mStudyDefinition.shortName = AccDistShortName
    mStudyDefinition.Description = "Accumulation/Distribution tracks buying and selling " & _
                                "by combining price movements and volume"
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(AccDistInputPrice)
    inputDef.inputType = InputTypeReal
    inputDef.Description = "Price"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(AccDistInputVolume)
    inputDef.inputType = InputTypeInteger
    inputDef.Description = "Volume"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(AccDistValueAccDist)
    valueDef.Description = "The Accumulation/Distribution value"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueMode = ValueModeNone
    valueDef.valueType = ValueTypeReal
    
End If

Set StudyDefinition = mStudyDefinition.Clone
End Property

'@================================================================================
' Helper Function
'@================================================================================





