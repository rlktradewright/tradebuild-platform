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

Implements ITask
Implements IValueReplayTask

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

Private mStudyToNotify As IStudy
Private mTargetListener As IStudyValueListener
Private mSourceStudy As IStudy

Private mFirstValueNumber As Long
Private mNextValueNumber As Long
Private mValuesNotified As Long

Private mValueName As String
Private mValueNames() As String
Private mNotifyName As String
Private mNotifyNames() As String

Private mReplayValueNames() As String
Private mReplayNotifyNames() As String

Private mNumberOfCachedValues() As Long

Private mSessionBuilder As SessionBuilder

Private mResult As ValueReplayResult

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IValueReplayTask Interface Members
'@================================================================================

Private Property Get IValueReplayTask_Discriminator() As Long
IValueReplayTask_Discriminator = 0
End Property

Private Property Get IValueReplayTask_NotifyName() As String
IValueReplayTask_NotifyName = mNotifyName
End Property

Private Property Get IValueReplayTask_NotifyNames() As String()
IValueReplayTask_NotifyNames = mNotifyNames
End Property

Private Property Get IValueReplayTask_SourceStudy() As IStudy
Set IValueReplayTask_SourceStudy = mSourceStudy
End Property

Private Property Get IValueReplayTask_Target() As Object
Set IValueReplayTask_Target = mTarget
End Property

Private Property Get IValueReplayTask_TaskName() As String
IValueReplayTask_TaskName = mTaskContext.Name
End Property

Private Property Get IValueReplayTask_ValueName() As String
IValueReplayTask_ValueName = mValueName
End Property

Private Property Get IValueReplayTask_ValueNames() As String()
IValueReplayTask_ValueNames = mValueNames
End Property

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_cancel()
doCancel
End Sub

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

If mSourceStudy.NumberOfCachedValues = 0 Then
    mTaskContext.Progress 100, Empty
    mTaskContext.Finish mResult, False
    Set mTaskContext = Nothing
    Exit Sub
End If

If mTaskContext.CancelPending Then
    doCancel
    Exit Sub
End If

If mTargetIsStudy And mSessionBuilder Is Nothing Then
    mTaskContext.Suspend -1
    Exit Sub
End If

If processValue(mNextValueNumber) Then
    mNextValueNumber = mNextValueNumber + 1
    mValuesNotified = mValuesNotified + 1
    
    If mValuesNotified Mod 20 = 0 Then
        ' Notify progress each 20 values
        mTaskContext.Progress 100 * mValuesNotified / (mSourceStudy.NumberOfCachedValues - mFirstValueNumber + 1), Empty
    End If
Else
    GStudies.Logger.Log "Study (" & mSourceStudy.Id & ") " & mSourceStudy.InstanceName & _
                " replayed " & mValuesNotified & _
                " values for " & Join(mReplayValueNames(), ","), _
                ProcName, ModuleName, LogLevelHighDetail
    mTaskContext.Progress 100, Empty
    mTaskContext.Finish mResult, False
    Set mTaskContext = Nothing
End If

Exit Sub

Err:
GStudies.NotifyUnhandledError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

Set mSessionBuilder = ev.Future.Value
mTaskContext.Continue

Exit Sub

Err:
GStudies.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Discriminator() As Long
Discriminator = 0
End Property

Public Property Get NotifyName() As String
NotifyName = mNotifyName
End Property

Public Property Get NotifyNames() As String()
NotifyNames = NotifyNames
End Property

Public Property Get SourceStudy() As IStudy
Set SourceStudy = mSourceStudy
End Property

Public Property Get Target() As Object
Set Target = mTarget
End Property

Public Property Get ValueName() As String
ValueName = mValueName
End Property

Public Property Get ValueNames() As String()
ValueNames = mValueNames
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pStudyManager As StudyManager, _
                ByVal Target As Object, _
                ByVal SourceStudy As IStudy, _
                ByRef replayValueNames() As String, _
                ByRef replayNotifyNames() As String, _
                ByVal numberOfValuesToReplay As Long, _
                ByRef ValueNames() As String, _
                ByRef NotifyNames() As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument Not (UBound(ValueNames) > 0 And numberOfValuesToReplay <> ReplayNumbers.ReplayAll), "Must replay all values if more than one Value name is being replayed"

Set mTarget = Target
Set mSourceStudy = SourceStudy
mReplayValueNames = replayValueNames
mReplayNotifyNames = replayNotifyNames
mValueNames = ValueNames
mValueName = mValueNames(0)
mNotifyNames = NotifyNames
mNotifyName = mNotifyNames(0)

ReDim mNumberOfCachedValues(UBound(mReplayValueNames)) As Long
Dim i As Long
For i = 0 To UBound(mNumberOfCachedValues)
    mNumberOfCachedValues(i) = mSourceStudy.NumberOfCachedValues(mReplayValueNames(i))
Next

Set mResult = New ValueReplayResult
mResult.Initialise Target, SourceStudy, ValueNames, NotifyNames

Dim numValuesToNotify As Long
If TypeOf Target Is IStudy Then
    mTargetIsStudy = True
    Dim lTargetStudy As IStudy
    Set lTargetStudy = Target
    Set mStudyToNotify = lTargetStudy.BaseStudy
    
    Set mSessionBuilder = pStudyManager.GetStudyEntry(lTargetStudy).SessionBuilder
    
    If numberOfValuesToReplay = ReplayNumbers.ReplayAsRequired Then
        ' just replay the number of values needed by the study
        numValuesToNotify = lTargetStudy.NumberOfBarsRequired
    Else
        numValuesToNotify = numberOfValuesToReplay
    End If
Else
    Set mTargetListener = Target
    numValuesToNotify = numberOfValuesToReplay
    mReplayNotifyNames = mReplayValueNames
End If

If numValuesToNotify > mSourceStudy.NumberOfCachedValues Or _
    numValuesToNotify = ReplayNumbers.ReplayAll _
Then
    numValuesToNotify = mSourceStudy.NumberOfCachedValues
End If

mNextValueNumber = mSourceStudy.NumberOfCachedValues - numValuesToNotify + 1
mFirstValueNumber = mNextValueNumber

GStudies.Logger.Log "Study (" & mSourceStudy.Id & ") " & _
            mSourceStudy.InstanceName & _
            " replaying values for " & Join(mReplayValueNames(), ","), _
            ProcName, ModuleName, LogLevelHighDetail

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub doCancel()
Const ProcName As String = "doCancel"
On Error GoTo Err

mTaskContext.Progress 100 * mValuesNotified / (mSourceStudy.NumberOfCachedValues - mFirstValueNumber + 1), Empty
mTaskContext.Finish mResult, True
Set mTaskContext = Nothing

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub NotifyValue(ByRef pValue As SValue, ByVal pNotifyName As String)
Const ProcName As String = "notifyValue"
On Error GoTo Err

If IsEmpty(pValue.Value) Then Exit Sub

Dim ev As StudyValueEventData
Set ev.Source = mSourceStudy
ev.sVal = pValue
ev.ValueName = pNotifyName
If mTargetIsStudy Then
    mSessionBuilder.SetSessionCurrentTime ev.sVal.Timestamp
    mStudyToNotify.Notify ev
Else
    mTargetListener.Notify ev
End If

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processValue(ByVal pValueNumber As Long) As Boolean
Const ProcName As String = "processValue"
On Error GoTo Err

processValue = False
Dim i As Long
For i = 0 To UBound(mReplayValueNames)
    If mNextValueNumber <= mNumberOfCachedValues(i) Then
        processValue = True
        Dim valueEntry As SValue
        valueEntry = mSourceStudy.GetStudyValue(mReplayValueNames(i), pValueNumber)
        NotifyValue valueEntry, mReplayNotifyNames(i)
    End If
Next

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function



