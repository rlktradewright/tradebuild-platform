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

Private mStudyDefinition As StudiesSupp.studyDefinition
Private mParameters As ParamsSupp.parameters
Private mStudyValueConfigurations As studyValueConfigurations
Private mStudyValueHandlers As studyValueHandlers
Private mStudyHorizontalRules As studyHorizontalRules
Private mHorizontalRulesLineSeries As ChartSkil.LineSeries
Private mName As String
Private mServiceProviderName As String
Private mInstanceName As String
Private mInstanceFullyQualifiedName As String
Private mStudyId As String
Private mUnderlyingStudyId As String
Private mInputValueNames() As String
'Private mNotifyNames() As String
Private mChartRegionName As String

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
ReDim mInputValueNames(0) As String
mInputValueNames(0) = "$default"
Set mStudyValueConfigurations = New studyValueConfigurations
Set mStudyValueHandlers = New studyValueHandlers
Set mStudyHorizontalRules = New studyHorizontalRules
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

Public Property Let chartRegionName(ByVal value As String)
mChartRegionName = value
End Property

Public Property Get chartRegionName() As String
chartRegionName = mChartRegionName
End Property

Public Property Let horizontalRulesLineSeries( _
                ByVal value As ChartSkil.LineSeries)
Set mHorizontalRulesLineSeries = value
End Property

Public Property Get horizontalRulesLineSeries() As ChartSkil.LineSeries
Set horizontalRulesLineSeries = mHorizontalRulesLineSeries
End Property

Public Property Let inputValueNames(ByRef value() As String)
mInputValueNames = value
End Property

Public Property Get inputValueNames() As String()
inputValueNames = mInputValueNames
End Property

Public Property Let instanceName(ByVal value As String)
mInstanceName = value
End Property

Public Property Get instanceName() As String
instanceName = mInstanceName
End Property

Public Property Let instanceFullyQualifiedName(ByVal value As String)
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

Public Property Let parameters(ByVal value As ParamsSupp.parameters)
Set mParameters = value
End Property

Public Property Get parameters() As ParamsSupp.parameters
Set parameters = mParameters
End Property

Public Property Let serviceProviderName(ByVal value As String)
mServiceProviderName = value
End Property

Public Property Get serviceProviderName() As String
serviceProviderName = mServiceProviderName
End Property

Public Property Let studyDefinition(ByVal value As StudiesSupp.studyDefinition)
Set mStudyDefinition = value
End Property

Public Property Get studyDefinition() As StudiesSupp.studyDefinition
Set studyDefinition = mStudyDefinition
End Property

Public Property Get studyHorizontalRules() As studyHorizontalRules
Set studyHorizontalRules = mStudyHorizontalRules
End Property

Public Property Let studyId(ByVal value As String)
mStudyId = value
End Property

Public Property Get studyId() As String
studyId = mStudyId
End Property

Public Property Get studyValueConfigurations() As studyValueConfigurations
Set studyValueConfigurations = mStudyValueConfigurations
End Property

Public Property Get studyValueHandlers() As studyValueHandlers
Set studyValueHandlers = mStudyValueHandlers
End Property

Public Property Let underlyingStudyId(ByVal value As String)
mUnderlyingStudyId = value
End Property

Public Property Get underlyingStudyId() As String
underlyingStudyId = mUnderlyingStudyId
End Property

'================================================================================
' Methods
'================================================================================

Public Function clone() As StudyConfiguration
Set clone = New StudyConfiguration
clone.chartRegionName = mChartRegionName
clone.inputValueNames = mInputValueNames
clone.instanceFullyQualifiedName = mInstanceFullyQualifiedName
clone.instanceName = mInstanceName
clone.name = mName
'clone.notifyNames = mNotifyNames
clone.parameters = mParameters.clone
clone.serviceProviderName = mServiceProviderName
clone.studyDefinition = mStudyDefinition

Dim hr As StudyHorizontalRule
Dim newHr As StudyHorizontalRule
Dim newHrs As studyHorizontalRules
Set newHrs = clone.studyHorizontalRules
For Each hr In mStudyHorizontalRules
    Set newHr = newHrs.add
    newHr.Color = hr.Color
    newHr.style = hr.style
    newHr.thickness = hr.thickness
    newHr.y = hr.y
Next

clone.studyId = mStudyId

Dim svc As StudyValueConfiguration
Dim newSvc As StudyValueConfiguration
Dim newSvcs As studyValueConfigurations

Set newSvcs = clone.studyValueConfigurations

For Each svc In mStudyValueConfigurations
    Set newSvc = newSvcs.add(svc.valueName)
    newSvc.chartRegionName = svc.chartRegionName
    newSvc.Color = svc.Color
    newSvc.displayMode = svc.displayMode
    newSvc.histogramBarWidth = svc.histogramBarWidth
    newSvc.includeInAutoscale = svc.includeInAutoscale
    newSvc.includeInChart = svc.includeInChart
    newSvc.layer = svc.layer
    newSvc.lineStyle = svc.lineStyle
    newSvc.lineThickness = svc.lineThickness
Next

clone.underlyingStudyId = mUnderlyingStudyId
End Function

Public Sub finish()
Dim svh As StudyValueHandler
For Each svh In mStudyValueHandlers
    svh.finish False
Next
Set mStudyValueHandlers = Nothing
Set mHorizontalRulesLineSeries = Nothing
End Sub

'================================================================================
' Helper Functions
'================================================================================

