VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyConfiguration"
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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

'Private mStudyDefinition As studyDefinition
Private mParameters                         As parameters
Private mStudyValueConfigurations           As studyValueConfigurations
Private mStudyValueHandlers                 As studyValueHandlers
Private mStudyHorizontalRules               As studyHorizontalRules
Private mHorizontalRulesLineSeries          As LineSeries
Private mName                               As String
Private mStudyLibraryName                   As String
'Private mInstanceName As String
Private mInstanceFullyQualifiedName         As String
Private mStudy                              As study
Private mUnderlyingStudy                    As study
Private mInputValueNames()                  As String
'Private mNotifyNames() As String
Private mChartRegionName                    As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mInputValueNames(0) As String
mInputValueNames(0) = "$default"
Set mStudyValueConfigurations = New studyValueConfigurations
Set mStudyValueHandlers = New studyValueHandlers
Set mStudyHorizontalRules = New studyHorizontalRules
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let chartRegionName(ByVal value As String)
mChartRegionName = value
End Property

Public Property Get chartRegionName() As String
chartRegionName = mChartRegionName
End Property

Public Property Let horizontalRulesLineSeries( _
                ByVal value As LineSeries)
Set mHorizontalRulesLineSeries = value
End Property

Public Property Get horizontalRulesLineSeries() As LineSeries
Set horizontalRulesLineSeries = mHorizontalRulesLineSeries
End Property

Public Property Let inputValueNames(ByRef value() As String)
mInputValueNames = value
End Property

Public Property Get inputValueNames() As String()
inputValueNames = mInputValueNames
End Property

'Public Property Let instanceName(ByVal value As String)
'mInstanceName = value
'End Property
'
'Public Property Get instanceName() As String
'instanceName = mInstanceName
'End Property

Friend Property Let instanceFullyQualifiedName(ByVal value As String)
mInstanceFullyQualifiedName = value
End Property

Public Property Get instanceFullyQualifiedName() As String
instanceFullyQualifiedName = mInstanceFullyQualifiedName
End Property

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

'Public Property Let notifyNames(ByRef value() As String)
'mNotifyNames = value
'End Property
'
'Public Property Get notifyNames() As String()
'notifyNames = mNotifyNames
'End Property

Public Property Let parameters(ByVal value As parameters)
Set mParameters = value
End Property

Public Property Get parameters() As parameters
Set parameters = mParameters
End Property

Public Property Let StudyLibraryName(ByVal value As String)
mStudyLibraryName = value
End Property

Public Property Get StudyLibraryName() As String
StudyLibraryName = mStudyLibraryName
End Property

'Public Property Let studyDefinition(ByVal value As studyDefinition)
'Set mStudyDefinition = value
'End Property

'Public Property Get studyDefinition() As studyDefinition
'Set studyDefinition = mStudy.studyDefinition
'End Property

Public Property Get studyHorizontalRules() As studyHorizontalRules
Set studyHorizontalRules = mStudyHorizontalRules
End Property

Public Property Let study(ByVal value As study)
Set mStudy = value
End Property

Public Property Get study() As study
Set study = mStudy
End Property

Public Property Get studyValueConfigurations() As studyValueConfigurations
Set studyValueConfigurations = mStudyValueConfigurations
End Property

Public Property Get studyValueHandlers() As studyValueHandlers
Set studyValueHandlers = mStudyValueHandlers
End Property

Public Property Let underlyingStudy(ByVal value As study)
Set mUnderlyingStudy = value
End Property

Public Property Get underlyingStudy() As study
Set underlyingStudy = mUnderlyingStudy
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As StudyConfiguration
Set clone = New StudyConfiguration
clone.chartRegionName = mChartRegionName
clone.inputValueNames = mInputValueNames
clone.name = mName
clone.parameters = mParameters.clone
clone.StudyLibraryName = mStudyLibraryName

Dim hr As StudyHorizontalRule
Dim newHr As StudyHorizontalRule
Dim newHrs As studyHorizontalRules
Set newHrs = clone.studyHorizontalRules
For Each hr In mStudyHorizontalRules
    Set newHr = newHrs.add
    newHr.color = hr.color
    newHr.style = hr.style
    newHr.thickness = hr.thickness
    newHr.y = hr.y
Next

clone.study = mStudy

Dim svc As StudyValueConfiguration
Dim newSvc As StudyValueConfiguration
Dim newSvcs As studyValueConfigurations

Set newSvcs = clone.studyValueConfigurations

For Each svc In mStudyValueConfigurations
    Set newSvc = newSvcs.add(svc.valueName)
    newSvc.barStyle = svc.barStyle
    newSvc.chartRegionName = svc.chartRegionName
    newSvc.dataPointStyle = svc.dataPointStyle
    newSvc.includeInChart = svc.includeInChart
    newSvc.layer = svc.layer
    newSvc.lineStyle = svc.lineStyle
    newSvc.textStyle = svc.textStyle
Next

clone.underlyingStudy = mUnderlyingStudy
End Function

Public Sub finish()
Dim svh As StudyValueHandler
For Each svh In mStudyValueHandlers
    svh.finish False
    mStudy.removeStudyValueListener svh
Next
Set mStudyValueHandlers = Nothing
Set mHorizontalRulesLineSeries = Nothing
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

