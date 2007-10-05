VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueConfigurations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private mStudyValueConfigurations As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStudyValueConfigurations = New Collection
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

Public Property Get count() As Long
count = mStudyValueConfigurations.count
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function add( _
                ByVal valueName As String) As StudyValueConfiguration
Set add = New StudyValueConfiguration
add.valueName = valueName
mStudyValueConfigurations.add add, valueName
End Function

Public Function clone() As studyValueConfigurations
Dim svc As StudyValueConfiguration
Dim newSvc As StudyValueConfiguration

Set clone = New studyValueConfigurations

For Each svc In mStudyValueConfigurations
    newSvc = clone.add(svc.valueName)
    newSvc.barStyle = svc.barStyle
    newSvc.chartRegionName = svc.chartRegionName
    newSvc.dataPointStyle = svc.dataPointStyle
    newSvc.includeInChart = svc.includeInChart
    newSvc.layer = svc.layer
    newSvc.lineStyle = svc.lineStyle
    newSvc.textStyle = svc.textStyle
Next
End Function

Public Function item( _
                ByVal index As Long) As StudyValueConfiguration
Set item = mStudyValueConfigurations.item(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mStudyValueConfigurations.[_NewEnum]
End Function

'@================================================================================
' Helper Functions
'@================================================================================

