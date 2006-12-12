VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InputStudy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuildSP.IStudy

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

Private mIdent As String
Private mName As String

Private mStudyDef As studyDefinition

Private mStudies As IStudies
Private mListeners As IStudyValueListeners

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
'mIdent = gGenerateGUIDString
Set mStudies = New Studies
Set mListeners = New StudyValueListeners

Set mStudyDef = New studyDefinition
mStudyDef.defaultRegion = StudyDefaultRegions.DefaultRegionNone

End Sub

'================================================================================
' IStudy Interface Members
'================================================================================

Private Function IStudy_addStudy( _
                ByVal study As TradeBuildSP.IStudy, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As Tasks.TaskCompletion
mStudies.add study, valueNames
End Function

Private Function IStudy_addStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As Tasks.TaskCompletion
mListeners.add listener, valueName
End Function

Private Property Get IStudy_baseStudy() As TradeBuildSP.IStudy
Set IStudy_baseStudy = Me
End Property

Private Function IStudy_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As TradeBuildSP.StudyValue

End Function

Private Property Get IStudy_id() As String
IStudy_id = mIdent
End Property

Private Sub IStudy_initialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                valueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)

End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = mName
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = mName
End Property

Private Sub IStudy_Notify( _
                ByRef ev As TradeBuildSP.StudyValueEvent)
Dim evOut As StudyValueEvent
evOut = ev
evOut.barStartTime = 0  ' a higher level study must generate this
mStudies.notify evOut
mListeners.notify evOut
End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long

End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal valueName As String = "") As Long

End Function

Private Property Get IStudy_parameters() As TradeBuildSP.IParameters

End Property

Private Sub IStudy_removeStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
mListeners.remove listener
End Sub

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition

Set IStudy_studyDefinition = mStudyDef.clone

End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub addInput( _
                ByVal name As String, _
                ByVal description As String, _
                ByVal valueType As StudiesSupp.StudyInputTypes, _
                ByVal isDefault As Boolean)
Dim valueDef As StudyValueDefinition

    
Set valueDef = mStudyDef.studyValueDefinitions.add(name)
valueDef.description = description
valueDef.isDefault = isDefault
valueDef.valueMode = ValueModeNone
Select Case valueType
Case InputTypeNone
    valueDef.valueType = ValueTypeNone
Case InputTypeInteger
    valueDef.valueType = ValueTypeInteger
Case InputTypeReal
    valueDef.valueType = ValueTypeReal
Case InputTypeString
    valueDef.valueType = ValueTypeString
Case InputTypeDate
    valueDef.valueType = ValueTypeDate
Case InputTypeBoolean
    valueDef.valueType = ValueTypeBoolean
End Select
    
End Sub

Friend Sub initialise( _
                ByVal name As String)
mName = name
mIdent = name
mStudyDef.name = mName
End Sub
'================================================================================
' Helper Functions
'================================================================================



