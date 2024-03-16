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

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "StudyDefinition"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mName As String
Private mShortName As String
Private mDescription As String

Private mNeedsBars As Boolean

Private mDefaultRegion As StudyDefaultRegions

Private mStudyInputDefinitions As StudyInputDefinitions
Private mStudyValueDefinitions As StudyValueDefinitions
Private mStudyParameterDefinitions As StudyParameterDefinitions

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStudyInputDefinitions = New StudyInputDefinitions
Set mStudyValueDefinitions = New StudyValueDefinitions
Set mStudyParameterDefinitions = New StudyParameterDefinitions
mNeedsBars = True
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let DefaultRegion(ByVal Value As StudyDefaultRegions)
mDefaultRegion = Value
End Property

Public Property Get DefaultRegion() As StudyDefaultRegions
DefaultRegion = mDefaultRegion
End Property

Public Property Get DefaultValueName() As String
Dim svd As StudyValueDefinition
Const ProcName As String = "DefaultValueName"
On Error GoTo Err

For Each svd In mStudyValueDefinitions
    If svd.IsDefault Then
        DefaultValueName = svd.Name
        Exit Property
    End If
Next

Exit Property

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Description(ByVal Value As String)
mDescription = Value
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Public Property Let Name(ByVal Value As String)
mName = Value
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let NeedsBars(ByVal Value As Boolean)
mNeedsBars = Value
End Property

Public Property Get NeedsBars() As Boolean
NeedsBars = mNeedsBars
End Property

Public Property Let ShortName(ByVal Value As String)
mShortName = Value
End Property

Public Property Get ShortName() As String
ShortName = mShortName
End Property

Public Property Get StudyInputDefinitions() As StudyInputDefinitions
Set StudyInputDefinitions = mStudyInputDefinitions
End Property

Public Property Get StudyParameterDefinitions() As StudyParameterDefinitions
Set StudyParameterDefinitions = mStudyParameterDefinitions
End Property

Public Property Get StudyValueDefinitions() As StudyValueDefinitions
Set StudyValueDefinitions = mStudyValueDefinitions
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Clone() As StudyDefinition
Dim newSIDs As StudyInputDefinitions
Dim newSID As StudyInputDefinition
Dim sid As StudyInputDefinition
Dim newSPDs As StudyParameterDefinitions
Dim newSPD As StudyParameterDefinition
Dim spd As StudyParameterDefinition
Dim newSVDs As StudyValueDefinitions
Dim newSVD As StudyValueDefinition
Dim svd As StudyValueDefinition

Const ProcName As String = "Clone"
Dim failpoint As Long

On Error GoTo Err

Set Clone = New StudyDefinition
Clone.DefaultRegion = mDefaultRegion
Clone.Description = mDescription
Clone.Name = mName
Clone.NeedsBars = mNeedsBars
Clone.ShortName = mShortName

Set newSIDs = Clone.StudyInputDefinitions
For Each sid In mStudyInputDefinitions
    Set newSID = newSIDs.Add(sid.Name)
    newSID.Description = sid.Description
    newSID.InputType = sid.InputType
Next

Set newSPDs = Clone.StudyParameterDefinitions
For Each spd In mStudyParameterDefinitions
    Set newSPD = newSPDs.Add(spd.Name)
    newSPD.Description = spd.Description
    newSPD.MaximumValue = spd.MaximumValue
    newSPD.MinimumValue = spd.MinimumValue
    newSPD.ParameterType = spd.ParameterType
    newSPD.PermittedValues = spd.PermittedValues
Next

Set newSVDs = Clone.StudyValueDefinitions
For Each svd In mStudyValueDefinitions
    Set newSVD = newSVDs.Add(svd.Name)
    newSVD.DefaultRegion = svd.DefaultRegion
    newSVD.Description = svd.Description
    newSVD.IncludeInChart = svd.IncludeInChart
    newSVD.IsDefault = svd.IsDefault
    newSVD.MaximumValue = svd.MaximumValue
    newSVD.MinimumValue = svd.MinimumValue
    newSVD.ValueMode = svd.ValueMode
    newSVD.ValueStyle = svd.ValueStyle
    newSVD.ValueType = svd.ValueType
Next

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

