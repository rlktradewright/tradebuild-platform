VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueListenerProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Implements TradeBuildSP.IStudyValueListener

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

Private mStudyValueListener As TradeBuild.studyValueListener

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IStudyValueListener Interface Members
'================================================================================

Private Sub IStudyValueListener_Notify(ev As TradeBuildSP.StudyValueEvent)
Dim tradeBuildEv As TradeBuild.StudyValueEvent
Set tradeBuildEv.source = ev.source
tradeBuildEv.value = ev.value
tradeBuildEv.ValueName = ev.ValueName
tradeBuildEv.Timestamp = ev.Timestamp
tradeBuildEv.barNumber = ev.barNumber
mStudyValueListener.notify tradeBuildEv
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let studyValueListener( _
                ByVal value As TradeBuild.studyValueListener)
Set mStudyValueListener = value
End Property

Friend Property Get studyValueListener() As TradeBuild.studyValueListener
Set studyValueListener = mStudyValueListener
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

