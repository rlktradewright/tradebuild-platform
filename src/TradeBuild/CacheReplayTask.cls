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

Implements Task
Implements TradeBuildSP.IValueReplayTask

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

Private mTargetStudy As TradeBuildSP.IStudy
Private mStudyToNotify As TradeBuildSP.IStudy
Private mTargetListener As TradeBuildSP.IStudyValueListener
Private mSourceStudy As TradeBuildSP.IStudy
Private mDiscriminator As Long
Private mValueCache As TradeBuild.StudyValueCache

Private mNextValueNumber As Long
Private mValuesNotified As Long

Private mValueName As String
Private mValueNames() As String
Private mNotifyName As String

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IValueReplayTask Interface Members
'================================================================================

Private Property Get IValueReplayTask_discriminator() As Long
IValueReplayTask_discriminator = mDiscriminator
End Property

Private Property Get IValueReplayTask_notifyName() As String
IValueReplayTask_notifyName = mNotifyName
End Property

Private Property Get IValueReplayTask_sourceStudy() As TradeBuildSP.IStudy
Set IValueReplayTask_sourceStudy = mSourceStudy
End Property

Private Property Get IValueReplayTask_target() As Object
Set IValueReplayTask_target = mTarget
End Property

Private Property Get IValueReplayTask_ValueName() As String
IValueReplayTask_ValueName = mValueName
End Property

Private Property Get IValueReplayTask_valueNames() As String()
IValueReplayTask_valueNames = mValueNames
End Property

'================================================================================
' Task Interface Members
'================================================================================

Private Sub Task_run()
Dim ev As TradeBuildSP.StudyValueEvent
Dim valueEntry As StudyValueStruct

If mValueCache.Count = 0 Then
    mTaskContext.finish
    Set mTaskContext = Nothing
    Exit Sub
End If

valueEntry = mValueCache.getStudyValue(mNextValueNumber)
If Not IsEmpty(valueEntry.value) Then
    Set ev.source = mSourceStudy
    ev.timestamp = valueEntry.timestamp
    ev.barNumber = valueEntry.barNumber
    ev.value = valueEntry.value
    
    If mTargetIsStudy Then
        ev.valueName = mNotifyName
        mStudyToNotify.notify ev
    Else
        ev.valueName = mValueName
        mTargetListener.notify ev
    End If
    
End If

mNextValueNumber = mNextValueNumber + 1
mValuesNotified = mValuesNotified + 1
If mNextValueNumber = mValueCache.Count Then
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

Public Property Get sourceStudy() As TradeBuildSP.IStudy
Set sourceStudy = mSourceStudy
End Property

Public Property Get target() As Object
Set target = mTarget
End Property

Public Property Get valueName() As String
valueName = mValueName
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Initialise( _
                ByVal valueCache As StudyValueCache, _
                ByVal target As Object, _
                ByVal sourceStudy As TradeBuildSP.IStudy, _
                ByVal numUnderlyingValuesToUse As Long, _
                ByVal discriminator As Long)
Dim numValuesToNotify As Long

Set mValueCache = valueCache
Set mTarget = target
Set mSourceStudy = sourceStudy
mValueName = valueCache.valueName
mDiscriminator = discriminator

If TypeOf target Is TradeBuildSP.IStudy Then
    mTargetIsStudy = True
    Set mTargetStudy = target
    Set mStudyToNotify = mTargetStudy.baseStudy
    mNotifyName = mTargetStudy.studyDefinition.studyInputDefinitions.Item(1).name
    If numUnderlyingValuesToUse < 0 Then
        ' just replay the number of values needed by the study
        numValuesToNotify = mTargetStudy.numberOfBarsRequired
    Else
        numValuesToNotify = numUnderlyingValuesToUse
    End If
Else
    Set mTargetListener = target
    If numUnderlyingValuesToUse > 0 Then
        numValuesToNotify = numUnderlyingValuesToUse
    End If
End If
If numValuesToNotify > mValueCache.Count Then numValuesToNotify = mValueCache.Count

mNextValueNumber = mValueCache.Count - numValuesToNotify + 1

End Sub

'================================================================================
' Helper Functions
'================================================================================



