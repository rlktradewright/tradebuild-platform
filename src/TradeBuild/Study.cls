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

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let defaultParameters(ByVal value As TradeBuild.parameters)
mStudy.defaultParameters = value
End Property

Public Property Get defaultParameters() As TradeBuild.parameters
Set defaultParameters = mStudy.defaultParameters
End Property

Public Property Get id() As String
id = mStudy.id
End Property

Public Property Get instanceName() As String
instanceName = mStudy.instanceName
End Property

Public Property Get instancePath() As String
instancePath = mStudy.instancePath
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

'================================================================================
' Methods
'================================================================================

Public Sub addStudyValueListener( _
                ByVal listener As TradeBuild.studyValueListener, _
                valueNames() As String)
Dim listenerProxy As StudyValueListenerProxy
Set listenerProxy = New StudyValueListenerProxy
listenerProxy.studyValueListener = listener
mStudy.addStudyValueListener listenerProxy, valueNames
End Sub

Public Function getStudyValue( _
        ByVal ValueName As String, _
        ByVal ref As Long, _
        ByRef parameters() As Variant _
        ) As TradeBuild.StudyValueStruct
Dim sv As TradeBuildSP.StudyValue
sv = mStudy.getStudyValue(ValueName, ref, parameters)
getStudyValue.barNumber = sv.barNumber
getStudyValue.Timestamp = sv.Timestamp
getStudyValue.value = sv.value
End Function

'================================================================================
' Helper Functions
'================================================================================


