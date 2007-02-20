VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'Implements StudyDefinition

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mName As String
Private mShortName As String
Private mDescription As String

Private mNeedsBars As Boolean

Private mDefaultRegion As StudyDefaultRegions

Private mStudyInputDefinitions As studyInputDefinitions
Private mStudyValueDefinitions As studyValueDefinitions
Private mStudyParameterDefinitions As studyParameterDefinitions

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mStudyInputDefinitions = New studyInputDefinitions
Set mStudyValueDefinitions = New studyValueDefinitions
Set mStudyParameterDefinitions = New studyParameterDefinitions
mNeedsBars = True
End Sub

'================================================================================
' IStudyDefinition Interface Members
'================================================================================

'Private Function IStudyDefinition_Clone() As StudyDefinition
'Set IStudyDefinition_Clone = clone
'End Function
'
'Private Property Let IStudyDefinition_defaultRegion( _
'                ByVal RHS As TradeBuildSP.StudyDefaultRegions)
'mDefaultRegion = RHS
'End Property
'
'Private Property Get IStudyDefinition_defaultRegion() As TradeBuildSP.StudyDefaultRegions
'IStudyDefinition_defaultRegion = mDefaultRegion
'End Property
'
'Private Property Let IStudyDefinition_Description(ByVal RHS As String)
'mDescription = RHS
'End Property
'
'Private Property Get IStudyDefinition_Description() As String
'IStudyDefinition_Description = mDescription
'End Property
'
'Private Property Let IStudyDefinition_name(ByVal RHS As String)
'mName = RHS
'End Property
'
'Private Property Get IStudyDefinition_name() As String
'IStudyDefinition_name = mName
'End Property
'
'Private Property Let IStudyDefinition_shortName(ByVal RHS As String)
'mShortName = RHS
'End Property
'
'Private Property Get IStudyDefinition_shortName() As String
'IStudyDefinition_shortName = mShortName
'End Property
'
'Private Property Get IStudyDefinition_StudyInputDefinitions() As StudyInputDefinitions
'Set IStudyDefinition_StudyInputDefinitions = mStudyInputDefinitions
'End Property
'
'Private Property Get IStudyDefinition_StudyParameterDefinitions() As StudyParameterDefinitions
'Set IStudyDefinition_StudyParameterDefinitions = mStudyParameterDefinitions
'End Property
'
'Private Property Get IStudyDefinition_StudyValueDefinitions() As StudyValueDefinitions
'Set IStudyDefinition_StudyValueDefinitions = mStudyValueDefinitions
'End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let defaultRegion(ByVal value As StudyDefaultRegions)
mDefaultRegion = value
End Property

Public Property Get defaultRegion() As StudyDefaultRegions
defaultRegion = mDefaultRegion
End Property

Public Property Let description(ByVal value As String)
mDescription = value
End Property

Public Property Get description() As String
description = mDescription
End Property

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

Public Property Let needsBars(ByVal value As Boolean)
mNeedsBars = value
End Property

Public Property Get needsBars() As Boolean
needsBars = mNeedsBars
End Property

Public Property Let shortName(ByVal value As String)
mShortName = value
End Property

Public Property Get shortName() As String
shortName = mShortName
End Property

Public Property Get studyInputDefinitions() As studyInputDefinitions
Set studyInputDefinitions = mStudyInputDefinitions
End Property

Public Property Get studyParameterDefinitions() As studyParameterDefinitions
Set studyParameterDefinitions = mStudyParameterDefinitions
End Property

Public Property Get studyValueDefinitions() As studyValueDefinitions
Set studyValueDefinitions = mStudyValueDefinitions
End Property

'================================================================================
' Methods
'================================================================================

Public Function clone() As studyDefinition
Dim newSIDs As studyInputDefinitions
Dim newSID As StudyInputDefinition
Dim sid As StudyInputDefinition
Dim newSPDs As studyParameterDefinitions
Dim newSPD As StudyParameterDefinition
Dim spd As StudyParameterDefinition
Dim newSVDs As studyValueDefinitions
Dim newSVD As StudyValueDefinition
Dim svd As StudyValueDefinition

Set clone = New studyDefinition
clone.defaultRegion = mDefaultRegion
clone.description = mDescription
clone.name = mName
clone.needsBars = mNeedsBars
clone.shortName = mShortName

Set newSIDs = clone.studyInputDefinitions
For Each sid In mStudyInputDefinitions
    Set newSID = newSIDs.add(sid.name)
    newSID.description = sid.description
    newSID.inputType = sid.inputType
Next

Set newSPDs = clone.studyParameterDefinitions
For Each spd In mStudyParameterDefinitions
    Set newSPD = newSPDs.add(spd.name)
    newSPD.description = spd.description
    newSPD.maximumValue = spd.maximumValue
    newSPD.minimumValue = spd.minimumValue
    newSPD.parameterType = spd.parameterType
    newSPD.permittedValues = spd.permittedValues
Next

Set newSVDs = clone.studyValueDefinitions
For Each svd In mStudyValueDefinitions
    Set newSVD = newSVDs.add(svd.name)
    newSVD.description = svd.description
    newSVD.isDefault = svd.isDefault
    newSVD.maximumValue = svd.maximumValue
    newSVD.minimumValue = svd.minimumValue
    newSVD.valueMode = svd.valueMode
    newSVD.valueType = svd.valueType
Next

End Function

'================================================================================
' Helper Functions
'================================================================================

