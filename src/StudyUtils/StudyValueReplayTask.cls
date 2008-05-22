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

Implements task
Implements ValueReplayTask

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

Private mTaskContext As TaskContext

Private mTarget As Object
Private mTargetIsStudy As Boolean

Private mTargetStudy As study
Private mStudyToNotify As study
Private mTargetListener As StudyValueListener
Private mSourceStudy As study
Private mDiscriminator As Long

Private mFirstValueNumber As Long
Private mNextValueNumber As Long
Private mValuesNotified As Long

Private mValueName As String
Private mValueNames() As String
Private mNotifyName As String
Private mNotifyNames() As String

Private mReplayValueNames() As String
Private mReplayNotifyNames() As String

Private mSessionBuilder As SessionBuilder

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ValueReplayTask Interface Members
'@================================================================================

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

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub task_cancel()

End Sub

Private Property Get task_TaskName() As String
task_TaskName = mTaskContext.name
End Property

Private Sub Task_run()
Dim i As Long
Dim ev As StudyValueEvent
Dim valueEntry As SValue

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

If mSourceStudy.numberOfCachedValues = 0 Then
    mTaskContext.progress 100
    mTaskContext.finish 0, False
    Set mTaskContext = Nothing
    Exit Sub
End If

If mTaskContext.cancelPending Then
    mTaskContext.progress 100 * mValuesNotified / (mSourceStudy.numberOfCachedValues - mFirstValueNumber + 1)
    mTaskContext.finish 0, True
    Set mTaskContext = Nothing
    Exit Sub
End If

For i = 0 To UBound(mReplayValueNames)
    
    failpoint = 200
    
    valueEntry = mSourceStudy.getStudyValue(mReplayValueNames(i), mNextValueNumber)
     
    If Not IsEmpty(valueEntry.value) Then
        Set ev.source = mSourceStudy
        ev.sVal = valueEntry
        
        If mTargetIsStudy Then
            
            failpoint = 300
            
            mSessionBuilder.setSessionCurrentTime ev.sVal.timestamp
            
            ev.valueName = mReplayNotifyNames(i)
            mStudyToNotify.notify ev
        Else
        
            failpoint = 400
            
            ev.valueName = mReplayValueNames(i)
            mTargetListener.notify ev
        End If
    End If
Next

mNextValueNumber = mNextValueNumber + 1
mValuesNotified = mValuesNotified + 1

If mValuesNotified Mod 20 = 0 Then
    ' notify progress each 20 values
    mTaskContext.progress 100 * mValuesNotified / (mSourceStudy.numberOfCachedValues - mFirstValueNumber + 1)
End If

If mNextValueNumber > mSourceStudy.numberOfCachedValues Then
    Debug.Print "Study (" & mSourceStudy.id & ") " & _
                mSourceStudy.instanceName & _
                " replayed " & mValuesNotified & _
                " values for " & Join(mReplayValueNames(), ",")
    mTaskContext.progress 100
    mTaskContext.finish 0, False
    Set mTaskContext = Nothing
End If

Exit Sub

Err:
Err.Raise Err.Number, _
        "StudiesSupp" & "." & "StudyValueReplayTask" & "::" & "Task_run" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.description


End Sub

Private Property Let Task_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get discriminator() As Long
discriminator = mDiscriminator
End Property

Public Property Get notifyName() As String
notifyName = mNotifyName
End Property

Public Property Get notifyNames() As String()
notifyNames = notifyNames
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

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal target As Object, _
                ByVal sourceStudy As study, _
                ByRef replayValueNames() As String, _
                ByRef replayNotifyNames() As String, _
                ByVal numberOfValuesToReplay As Long, _
                ByRef valueNames() As String, _
                ByRef notifyNames() As String, _
                ByVal discriminator As Long)
Dim numValuesToNotify As Long
Dim studyEntry As StudyTableEntry

Set mTarget = target
Set mSourceStudy = sourceStudy
mReplayValueNames = replayValueNames
mReplayNotifyNames = replayNotifyNames
mValueNames = valueNames
mValueName = mValueNames(0)
mNotifyNames = notifyNames
mNotifyName = mNotifyNames(0)
mDiscriminator = discriminator

If TypeOf target Is study Then
    mTargetIsStudy = True
    Set mTargetStudy = target
    Set mStudyToNotify = mTargetStudy.baseStudy
    
    studyEntry = Globals.StudiesCollection(mTargetStudy.id)
    Set mSessionBuilder = studyEntry.sessionBldr
    
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
mFirstValueNumber = mNextValueNumber
End Sub

'@================================================================================
' Helper Functions
'@================================================================================







