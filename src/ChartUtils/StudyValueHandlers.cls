VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueHandlers"
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

Private mStudyValueHandlers As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStudyValueHandlers = New Collection
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
count = mStudyValueHandlers.count
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Add( _
                ByVal chartMgr As ChartManager, _
                ByVal Region As ChartRegion, _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition, _
                ByVal UpdatePerTick As Boolean) As StudyValueHandler
Set Add = New StudyValueHandler
Add.Initialise chartMgr, _
                Region, _
                studyValueConfig, _
                studyValueDef, _
                UpdatePerTick
mStudyValueHandlers.Add Add, studyValueConfig.ValueName
End Function

Friend Function Clear()
Set mStudyValueHandlers = New Collection
End Function

Public Function Item( _
                ByVal index As Variant) As StudyValueHandler
Attribute Item.VB_UserMemId = 0
Set Item = mStudyValueHandlers.Item(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Set NewEnum = mStudyValueHandlers.[_NewEnum]
End Function

'@================================================================================
' Helper Functions
'@================================================================================





