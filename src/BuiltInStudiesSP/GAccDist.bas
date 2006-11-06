Attribute VB_Name = "GAccDist"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const AccDistInputPrice As String = "Price"
Public Const AccDistInputVolume As String = "Volume"

Public Const AccDistValueAccDist As String = "AccDist"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Global object references
'================================================================================

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mDefaultParameters As IParameters
Private mStudyDefinition As IStudyDefinition

'================================================================================
' External function declarations
'================================================================================

'================================================================================
' Variables
'================================================================================

'================================================================================
' Procedures
'================================================================================

Public Property Let commonServiceConsumer( _
                ByVal value As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property


Public Property Let defaultParameters(ByVal value As IParameters)
' create a clone of the default parameters supplied by the caller
Set mDefaultParameters = value.Clone
End Property

Public Property Get defaultParameters() As IParameters
If mDefaultParameters Is Nothing Then
    Set mDefaultParameters = mCommonServiceConsumer.NewParameters
End If

' now create a clone of the default parameters for the caller
Set defaultParameters = mDefaultParameters.Clone
End Property

Public Property Get studyDefinition() As TradeBuildSP.IStudyDefinition
Dim inputDef As IStudyInputDefinition
Dim valueDef As IStudyValueDefinition

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = mCommonServiceConsumer.NewStudyDefinition
    mStudyDefinition.name = AccDistName
    mStudyDefinition.shortName = AccDistShortName
    mStudyDefinition.Description = "Accumulation/Distribution tracks buying and selling " & _
                                "by combining price movements and volume"
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(AccDistInputPrice)
    inputDef.name = AccDistInputPrice
    inputDef.inputType = InputTypeDouble
    inputDef.Description = "Price"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(AccDistInputVolume)
    inputDef.name = AccDistInputVolume
    inputDef.inputType = InputTypeInteger
    inputDef.Description = "Volume"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(AccDistValueAccDist)
    valueDef.name = AccDistValueAccDist
    valueDef.Description = "The Accumulation/Distribution value"
    valueDef.isDefault = True
    valueDef.defaultRegion = DefaultRegionNone
    valueDef.valueType = ValueTypeDouble
    
End If

Set studyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================





