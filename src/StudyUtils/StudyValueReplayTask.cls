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

Implements Task
Implements ValueReplayTask

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "StudyValueReplayTask"

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

Private mTargetStudy As Study
Private mStudyToNotify As Study
Private mTargetListener As StudyValueListener
Private mSourceStudy As Study

' The following no longer used as the cookie argument to startTask is used instead
'Private mDiscriminator As Long

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

Private mResult As ValueReplayResult

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ValueReplayTask Interface Members
'@================================================================================

Private Property Get ValueReplayTask_Discriminator() As Long
ValueReplayTask_Discriminator = 0
End Property

Private Property Get ValueReplayTask_NotifyName() As String
ValueReplayTask_NotifyName = mNotifyName
End Property

Private Property Get ValueReplayTask_NotifyNames() As String()
ValueReplayTask_NotifyNames = mNotifyNames
End Property

Private Property Get ValueReplayTask_SourceStudy() As Study
Set ValueReplayTask_SourceStudy = mSourceStudy
End Property

Private Property Get ValueReplayTask_Target() As Object
Set ValueReplayTask_Target = mTarget
End Property

Private Property Get ValueReplayTask_TaskName() As String
ValueReplayTask_TaskName = mTaskContext.Name
End Property

Private Property Get ValueReplayTask_ValueName() As String
ValueReplayTask_ValueName = mValueName
End Property

Private Property Get ValueReplayTask_ValueNames() As String()
ValueReplayTask_ValueNames = mValueNames
End Property

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub task_cancel()
mTaskContext.Finish Empty, True
End Sub

Private Property Get task_TaskName() As String
task_TaskName = mTaskContext.Name
End Property

Private Sub Task_run()
Dim i As Long
Dim ev As StudyValueEvent
Dim valueEntry As SValue

Const ProcName As String = "Task_run"
On Error GoTo Err

If mSourceStudy.NumberOfCachedValues = 0 Then
    mTaskContext.Progress 100
    mTaskContext.Finish mResult, False
    Set mTaskContext = Nothing
    Exit Sub
End If

If mTaskContext.CancelPending Then
    mTaskContext.Progress 100 * mValuesNotified / (mSourceStudy.NumberOfCachedValues - mFirstValueNumber + 1)
    mTaskContext.Finish mResult, True
    Set mTaskContext = Nothing
    Exit Sub
End If

For i = 0 To UBound(mReplayValueNames)
    
    valueEntry = mSourceStudy.GetStudyValue(mReplayValueNames(i), mNextValueNumber)
     
    If Not IsEmpty(valueEntry.value) Then
        Set ev.source = mSourceStudy
        ev.sVal = valueEntry
        
        If mTargetIsStudy Then
            
            mSessionBuilder.SetSessionCurrentTime ev.sVal.Timestamp
            
            ev.valueName = mReplayNotifyNames(i)
            mStudyToNotify.Notify ev
        Else
        
            ev.valueName = mReplayValueNames(i)
            mTargetListener.Notify ev
        End If
    End If
Next

mNextValueNumber = mNextValueNumber + 1
mValuesNotified = mValuesNotified + 1

If mValuesNotified Mod 20 = 0 Then
    ' Notify progress each 20 values
    mTaskContext.Progress 100 * mValuesNotified / (mSourceStudy.NumberOfCachedValues - mFirstValueNumber + 1)
End If

If mNextValueNumber > mSourceStudy.NumberOfCachedValues Then
    Debug.Print "Study (" & mSourceStudy.Id & ") " & _
                mSourceStudy.InstanceName & _
                " replayed " & mValuesNotified & _
                " values for " & Join(mReplayValueNames(), ",")
    mTaskContext.Progress 100
    mTaskContext.Finish mResult, False
    Set mTaskContext = Nothing
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
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

Public Property Get Discriminator() As Long
Discriminator = 0
End Property

Public Property Get notifyName() As String
notifyName = mNotifyName
End Property

Public Property Get notifyNames() As String()
notifyNames = notifyNames
End Property

Public Property Get sourceStudy() As Study
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
                ByVal sourceStudy As Study, _
                ByRef replayValueNames() As String, _
                ByRef replayNotifyNames() As String, _
                ByVal numberOfValuesToReplay As Long, _
                ByRef valueNames() As String, _
                ByRef notifyNames() As String, _
                ByVal Discriminator As Long)
Dim numValuesToNotify As Long
Dim studyEntry As StudyTableEntry

Const ProcName As String = "initialise"
On Error GoTo Err

Set mTarget = target
Set mSourceStudy = sourceStudy
mReplayValueNames = replayValueNames
mReplayNotifyNames = replayNotifyNames
mValueNames = valueNames
mValueName = mValueNames(0)
mNotifyNames = notifyNames
mNotifyName = mNotifyNames(0)
'mDiscriminator = Discriminator

Set mResult = New ValueReplayResult
mResult.initialise target, sourceStudy, valueNames, notifyNames

If TypeOf target Is Study Then
    mTargetIsStudy = True
    Set mTargetStudy = target
    Set mStudyToNotify = mTargetStudy.BaseStudy
    
    studyEntry = Globals.StudiesCollection(mTargetStudy.Id)
    Set mSessionBuilder = studyEntry.sessionBldr
    
    If numberOfValuesToReplay = ReplayNumbers.ReplayAsRequired Then
        ' just replay the number of values needed by the study
        numValuesToNotify = mTargetStudy.NumberOfBarsRequired
    Else
        numValuesToNotify = numberOfValuesToReplay
    End If
Else
    Set mTargetListener = target
    numValuesToNotify = numberOfValuesToReplay
End If

If numValuesToNotify > mSourceStudy.NumberOfCachedValues Or _
    numValuesToNotify = ReplayNumbers.ReplayAll _
Then
    numValuesToNotify = mSourceStudy.NumberOfCachedValues
End If

mNextValueNumber = mSourceStudy.NumberOfCachedValues - numValuesToNotify + 1
mFirstValueNumber = mNextValueNumber

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================







