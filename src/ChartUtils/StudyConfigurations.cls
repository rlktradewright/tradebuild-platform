VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyConfigurations"
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

Event ItemAdded(ByVal studyConfig As StudyConfiguration)
Event ItemRemoved(ByVal studyConfig As StudyConfiguration)

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

Private mStudyConfigurations As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStudyConfigurations = New Collection
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
count = mStudyConfigurations.count
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub add( _
                ByVal studyConfig As StudyConfiguration)
Dim key As String
Dim keyInc As String
Dim i As Long
Dim lStudyConfig As StudyConfiguration

key = studyConfig.study.instancePath

On Error Resume Next
If studyConfig.instanceFullyQualifiedName = "" Then
    Set lStudyConfig = mStudyConfigurations(key)
Else
    Set lStudyConfig = mStudyConfigurations(studyConfig.instanceFullyQualifiedName)
End If
On Error GoTo 0

If lStudyConfig Is studyConfig Then
    ' this studyConfig has already been added
    Exit Sub
End If

If lStudyConfig Is Nothing Then
    studyConfig.instanceFullyQualifiedName = key
Else
    Do
        i = i + 1
        keyInc = key & "." & i
        
        On Error Resume Next
        Set lStudyConfig = Nothing
        Set lStudyConfig = mStudyConfigurations(keyInc)
        On Error GoTo 0
    
    Loop Until lStudyConfig Is Nothing
    studyConfig.instanceFullyQualifiedName = keyInc
End If


mStudyConfigurations.add studyConfig, studyConfig.instanceFullyQualifiedName
RaiseEvent ItemAdded(studyConfig)
End Sub

Public Function item( _
                ByVal index As Variant) As StudyConfiguration
Set item = mStudyConfigurations.item(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mStudyConfigurations.[_NewEnum]
End Function

Public Sub remove( _
                ByVal studyConfig As StudyConfiguration)
mStudyConfigurations.remove studyConfig.instanceFullyQualifiedName
RaiseEvent ItemRemoved(studyConfig)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



