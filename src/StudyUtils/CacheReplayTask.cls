VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CacheReplayTask"
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

Implements Tasks.task
Implements ValueReplayTask

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

Private mTaskContext As TaskContext

Private mTarget As Object
Private mTargetIsStudy As Boolean

Private mTargetStudy As study
Private mStudyToNotify As study
Private mTargetListener As StudyValueListener
Private mSourceStudy As study
Private mDiscriminator As Long
Private mValueCache As StudyValueCache

Private mNextValueNumber As Long
Private mValuesNotified As Long

Private mValueName As String
Private mValueNames(0) As String
Private mNotifyName As String
Private mNotifyNames(0) As String

Private mSessionBuilder As SessionBuilder

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' ValueReplayTask Interface Members
'================================================================================

Private Property Get ValueReplayTask_discriminator() As Long
ValueReplayTask_discriminator = mDiscriminator
End Property

Private Property Get ValueReplayTask_notifyName() As String
ValueReplayTask_notifyName = mNotifyName
End Property

Private Property Get ValueReplayTask_notifyNames() As String()
ValueReplayTask_notifyNames = mNotifyNames
End Property

Private Property Get ValueReplayTask_sourceStudy() As study
Set ValueReplayTask_sourceStudy = mSourceStudy
End Property

Private Property Get ValueReplayTask_target() As Object
Set ValueReplayTask_target = mTarget
End Property

Private Property Get ValueReplayTask_taskName() As String
ValueReplayTask_taskName = mTaskContext.name
End Property

Private Property Get ValueReplayTask_ValueName() As String
ValueReplayTask_ValueName = mValueName
End Property

Private Property Get ValueReplayTask_valueNames() As String()
ValueReplayTask_valueNames = mValueNames
End Property

'================================================================================
' Task Interface Members
'================================================================================

Private Property Get task_TaskName() As String
task_TaskName = mTaskContext.name
End Property

Private Sub Task_run()
Dim ev As StudyValueEvent
Dim valueEntry As StudyValue

If mValueCache.count = 0 Then
    mTaskContext.finish
    Set mTaskContext = Nothing
    Exit Sub
End If

valueEntry = mValueCache.getStudyValue(mNextValueNumber)
If Not IsEmpty(valueEntry.value) Then
    Set ev.source = mSourceStudy
    ev.timestamp = valueEntry.timestamp
    ev.barNumber = valueEntry.barNumber
    ev.barStartTime = valueEntry.barStartTime
    If TypeOf valueEntry.value Is Object  Then
        Set ev.value = valueEntry.value
    Else
        ev.value = valueEntry.value
    End If
    
    If mTargetIsStudy Then
        mSessionBuilder.setSessionCurrentTime ev.timestamp
        ev.valueName = mNotifyName
        mStudyToNotify.notify ev
    Else
        ev.valueName = mValueName
        mTargetListener.notify ev
    End If
    
End If

mNextValueNumber = mNextValueNumber + 1
mValuesNotified = mValuesNotified + 1
If mNextValueNumber > mValueCache.count Then
    Debug.Print "Cache for " & mValueName & " replayed " & mValuesNotified & " values"
    mTaskContext.finish
    Set mTaskContext = Nothing
End If

End Sub

Private Property Let Task_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get discriminator() As Long
discriminator = mDiscriminator
End Property

Public Property Get notifyName() As String
notifyName = mNotifyName
End Property

Public Property Get notifyNames() As String()
notifyNames = mNotifyNames
End Property

Public Property Get sourceStudy() As study
Set sourceStudy = mSourceStudy
End Property

Public Property Get target() As Object
Set target = mTarget
End Property

Public Property Get valueName() As String
valueName = mValueName
End Property

Public Property Get valueNames() As String()
valueNames = mValueNames
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub initialise( _
                ByVal valueCache As StudyValueCache, _
                ByVal target As Object, _
                ByVal sourceStudy As study, _
                ByVal numberOfValuesToReplay As Long, _
                ByVal discriminator As Long)
Dim numValuesToNotify As Long
Dim studyEntry As StudyTableEntry

Set mValueCache = valueCache
Set mTarget = target
Set mSourceStudy = sourceStudy
mValueName = valueCache.valueName
mValueNames(0) = mValueName
mDiscriminator = discriminator

If TypeOf target Is study Then
    mTargetIsStudy = True
    Set mTargetStudy = target
    Set mStudyToNotify = mTargetStudy.baseStudy
    
    studyEntry = Globals.StudiesCollection(mTargetStudy.id)
    Set mSessionBuilder = studyEntry.sessionBldr
    
    mNotifyName = mTargetStudy.studyDefinition.studyInputDefinitions.Item(1).name
    mNotifyNames(0) = mNotifyName
    If numberOfValuesToReplay = ReplayNumbers.ReplayAsRequired Then
        ' just replay the number of values needed by the study
        numValuesToNotify = mTargetStudy.numberOfBarsRequired
    Else
        numValuesToNotify = numberOfValuesToReplay
    End If
Else
    Set mTargetListener = target
    numValuesToNotify = numberOfValuesToReplay
End If
If numValuesToNotify > mValueCache.count Or _
    numValuesToNotify = ReplayNumbers.ReplayAll _
Then
    numValuesToNotify = mValueCache.count
End If

mNextValueNumber = mValueCache.count - numValuesToNotify + 1

End Sub

'================================================================================
' Helper Functions
'================================================================================





