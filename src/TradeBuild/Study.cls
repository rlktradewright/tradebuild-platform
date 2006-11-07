VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Study"
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
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

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

Private mStudy As TradeBuildSP.IStudy

Private mStudyValueListeners As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mStudyValueListeners = New Collection
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'Public Property Let defaultParameters(ByVal value As TradeBuild.parameters)
'mStudy.defaultParameters = value
'End Property
'
'Public Property Get defaultParameters() As TradeBuild.parameters
'Set defaultParameters = mStudy.defaultParameters
'End Property

Public Property Get id() As String
id = mStudy.id
End Property

Public Property Get instanceName() As String
instanceName = mStudy.instanceName
End Property

Public Property Get instancePath() As String
instancePath = mStudy.instancePath
End Property

Public Property Get numberOfCachedValues() As Long
numberOfCachedValues = mStudy.numberOfCachedValues
End Property

Public Property Get parameters() As TradeBuild.parameters
Set parameters = mStudy.parameters
End Property

Friend Property Get study() As TradeBuildSP.IStudy
Set study = mStudy
End Property

Friend Property Let study(ByVal value As TradeBuildSP.IStudy)
Set mStudy = value
End Property

Public Property Get studyDefinition() As TradeBuild.studyDefinition
Set studyDefinition = mStudy.studyDefinition
End Property

'================================================================================
' Methods
'================================================================================

Public Function addStudyValueListener( _
                ByVal listener As TradeBuild.studyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskCompletion
Dim listenerProxy As StudyValueListenerProxy
Set listenerProxy = New StudyValueListenerProxy
listenerProxy.study = Me
listenerProxy.studyValueListener = listener
mStudyValueListeners.add listenerProxy
Set addStudyValueListener = mStudy.addStudyValueListener( _
                                listenerProxy, _
                                valueName, _
                                numberOfValuesToReplay, _
                                taskName, _
                                taskData)
End Function

Public Function getStudyValue( _
        ByVal valueName As String, _
        ByVal ref As Long) As TradeBuild.StudyValueStruct
Dim sv As TradeBuildSP.StudyValue
Dim parms() As Variant

If valueName = "" Then valueName = "$default"

parms = parameters
sv = mStudy.getStudyValue(valueName, ref)
getStudyValue.barNumber = sv.barNumber
getStudyValue.timestamp = sv.timestamp
getStudyValue.value = sv.value
End Function

Public Sub removeStudyValueListener( _
                ByVal listener As TradeBuild.studyValueListener)
Dim i As Long
Dim lListener As StudyValueListenerProxy
For i = mStudyValueListeners.Count To 1 Step -1
    Set lListener = mStudyValueListeners(i)
    If lListener.studyValueListener Is listener Then
        mStudyValueListeners.remove i
    End If
Next
mStudy.removeStudyValueListener lListener
End Sub

'================================================================================
' Helper Functions
'================================================================================


