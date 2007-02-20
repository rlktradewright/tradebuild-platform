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

Private mStudyValueHandlers As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mStudyValueHandlers = New Collection
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

Public Property Get count() As Long
count = mStudyValueHandlers.count
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                ByVal chartMgr As ChartManager, _
                ByVal region As ChartRegion, _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition, _
                ByVal updatePerTick As Boolean) As StudyValueHandler
Set add = New StudyValueHandler
add.initialise chartMgr, _
                region, _
                studyValueConfig, _
                studyValueDef, _
                updatePerTick
mStudyValueHandlers.add add
End Function

Public Function item( _
                ByVal index As Long) As StudyValueHandler
Set item = mStudyValueHandlers.item(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Set NewEnum = mStudyValueHandlers.[_NewEnum]
End Function

'================================================================================
' Helper Functions
'================================================================================





