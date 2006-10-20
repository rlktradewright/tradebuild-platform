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

Private mStudyDefinition As TradeBuild.studyDefinition
Private mParameters As TradeBuild.parameters
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
Private mInputValueName As String
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

Friend Property Let chartRegionName(ByVal value As String)
mChartRegionName = value
End Property

Public Property Get chartRegionName() As String
chartRegionName = mChartRegionName
End Property

Friend Property Let horizontalRulesLineSeries( _
                ByVal value As ChartSkil.LineSeries)
Set mHorizontalRulesLineSeries = value
End Property

Friend Property Get horizontalRulesLineSeries() As ChartSkil.LineSeries
Set horizontalRulesLineSeries = mHorizontalRulesLineSeries
End Property

Friend Property Let inputValueName(ByRef value As String)
mInputValueName = value
End Property

Public Property Get inputValueName() As String
inputValueName = mInputValueName
End Property

Friend Property Let instanceName(ByVal value As String)
mInstanceName = value
End Property

Public Property Get instanceName() As String
instanceName = mInstanceName
End Property

Friend Property Let instanceFullyQualifiedName(ByVal value As String)
mInstanceFullyQualifiedName = value
End Property

Public Property Get instanceFullyQualifiedName() As String
instanceFullyQualifiedName = mInstanceFullyQualifiedName
End Property

Friend Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

Friend Property Let parameters(ByVal value As TradeBuild.parameters)
Set mParameters = value
End Property

Public Property Get parameters() As TradeBuild.parameters
Set parameters = mParameters
End Property

Friend Property Let serviceProviderName(ByVal value As String)
mServiceProviderName = value
End Property

Public Property Get serviceProviderName() As String
serviceProviderName = mServiceProviderName
End Property

Friend Property Let studyDefinition(ByVal value As TradeBuild.studyDefinition)
Set mStudyDefinition = value
End Property

Public Property Get studyDefinition() As TradeBuild.studyDefinition
Set studyDefinition = mStudyDefinition
End Property

Public Property Get studyHorizontalRules() As studyHorizontalRules
Set studyHorizontalRules = mStudyHorizontalRules
End Property

Friend Property Let studyId(ByVal value As String)
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

Friend Property Let underlyingStudyId(ByVal value As String)
mUnderlyingStudyId = value
End Property

Public Property Get underlyingStudyId() As String
underlyingStudyId = mUnderlyingStudyId
End Property

'================================================================================
' Methods
'================================================================================

Friend Function clone() As StudyConfiguration
Set clone = New StudyConfiguration
clone.chartRegionName = mChartRegionName
clone.inputValueName = mInputValueName
clone.instanceFullyQualifiedName = mInstanceFullyQualifiedName
clone.instanceName = mInstanceName
clone.name = mName
clone.parameters = mParameters.clone
clone.serviceProviderName = mServiceProviderName
clone.studyDefinition = mStudyDefinition

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

clone.studyId = mStudyId

Dim svc As StudyValueConfiguration
Dim newSvc As StudyValueConfiguration
Dim newSvcs As studyValueConfigurations

Set newSvcs = clone.studyValueConfigurations

For Each svc In mStudyValueConfigurations
    Set newSvc = newSvcs.add(svc.valueName)
    newSvc.chartRegionName = svc.chartRegionName
    newSvc.color = svc.color
    newSvc.displayMode = svc.displayMode
    newSvc.histogramBarWidth = svc.histogramBarWidth
    newSvc.includeInAutoscale = svc.includeInAutoscale
    newSvc.includeInChart = svc.includeInChart
    newSvc.layer = svc.layer
    newSvc.lineStyle = svc.lineStyle
    newSvc.lineThickness = svc.lineThickness
    newSvc.multipleValuesPerBar = svc.multipleValuesPerBar
Next

clone.underlyingStudyId = mUnderlyingStudyId
End Function

Friend Sub finish()
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

