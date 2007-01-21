VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueReplayTask"
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

Implements task
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

Private mNextValueNumber As Long
Private mValuesNotified As Long

Private mValueName As String
Private mValueNames() As String
Private mNotifyNames() As String

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' ValueReplayTask Interface Members
'================================================================================

Private Property Get ValueReplayTask_discriminator() As Long
ValueReplayTask_discriminator = mDiscriminator
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
Dim i As Long
Dim ev As StudyValueEvent
Dim valueEntry As StudyValue

If mSourceStudy.numberOfCachedValues = 0 Then
    mTaskContext.finish
    Set mTaskContext = Nothing
    Exit Sub
End If

For i = 0 To UBound(mValueNames)
    valueEntry = mSourceStudy.getStudyValue(mValueNames(i), mNextValueNumber)
     
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
            ev.valueName = mNotifyNames(i)
            mStudyToNotify.notify ev
        Else
            ev.valueName = mValueNames(i)
            mTargetListener.notify ev
        End If
    End If
Next

mNextValueNumber = mNextValueNumber + 1
mValuesNotified = mValuesNotified + 1
If mNextValueNumber = mSourceStudy.numberOfCachedValues Then
    Debug.Print "Study " & mSourceStudy.instanceName & " replayed " & mValuesNotified & " values"
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
                ByVal target As Object, _
                ByVal sourceStudy As study, _
                ByRef valueNames() As String, _
                ByVal numberOfValuesToReplay As Long, _
                ByVal discriminator As Long)
Dim numValuesToNotify As Long
Dim inputDefs As studyInputDefinitions
Dim i As Long

Set mTarget = target
Set mSourceStudy = sourceStudy
mValueNames = valueNames
mValueName = mValueNames(0)

mDiscriminator = discriminator

If TypeOf target Is study Then
    mTargetIsStudy = True
    Set mTargetStudy = target
    Set mStudyToNotify = mTargetStudy.baseStudy
    
    Set inputDefs = mTargetStudy.studyDefinition.studyInputDefinitions
    ReDim mNotifyNames(inputDefs.count - 1) As String
    For i = 0 To inputDefs.count - 1
        mNotifyNames(i) = inputDefs.Item(i + 1).name
    Next
    
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
If numValuesToNotify > mSourceStudy.numberOfCachedValues Or _
    numValuesToNotify = ReplayNumbers.ReplayAll _
Then
    numValuesToNotify = mSourceStudy.numberOfCachedValues
End If

mNextValueNumber = mSourceStudy.numberOfCachedValues - numValuesToNotify + 1

End Sub

'================================================================================
' Helper Functions
'================================================================================







