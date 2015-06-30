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

Private Const ModuleName                As String = "CacheReplayTask"

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

Private mTargetStudy As IStudy
Private mStudyToNotify As IStudy
Private mTargetListener As IStudyValueListener
Private mSourceStudy As IStudy

Private mValueCache As valueCache

Private mFirstValueNumber As Long
Private mNextValueNumber As Long
Private mValuesNotified As Long

Private mValueName As String
Private mValueNames(0) As String
Private mNotifyName As String
Private mNotifyNames(0) As String

Private mSessionBuilder As SessionBuilder

Private mResult As ValueReplayResult

Private WithEvents mFutureWaiter As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

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

If mValueCache.Count = 0 Then
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

Dim valueEntry As SValue
valueEntry = mValueCache.GetSValue(mNextValueNumber)
If Not IsEmpty(valueEntry.Value) Then
    Dim ev As StudyValueEventData
    Set ev.Source = mSourceStudy
    ev.sVal = valueEntry
    
    If mTargetIsStudy Then
        mSessionBuilder.SetSessionCurrentTime ev.sVal.Timestamp
        ev.ValueName = mNotifyName
        mStudyToNotify.Notify ev
    Else
        ev.ValueName = mValueName
        mTargetListener.Notify ev
    End If
    
End If

mNextValueNumber = mNextValueNumber + 1
mValuesNotified = mValuesNotified + 1

If mValuesNotified Mod 20 = 0 Then
    ' Notify progress each 20 values
    mTaskContext.Progress 100 * mValuesNotified / (mSourceStudy.NumberOfCachedValues - mFirstValueNumber + 1), Empty
End If

If mNextValueNumber > mValueCache.Count Then
    gLogger.Log "Study (" & mSourceStudy.Id & ") " & mSourceStudy.InstanceName & _
                " replayed " & mValuesNotified & " values for " & mValueName, _
                ProcName, ModuleName, LogLevelHighDetail
    mTaskContext.Progress 100, Empty
    mTaskContext.Finish mResult, False
    Set mTaskContext = Nothing
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
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
gNotifyUnhandledError ProcName, ModuleName
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
NotifyNames = mNotifyNames
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
                ByVal valueCache As valueCache, _
                ByVal Target As Object, _
                ByVal SourceStudy As IStudy, _
                ByVal numberOfValuesToReplay As Long, _
                ByVal Discriminator As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mValueCache = valueCache
Set mTarget = Target
Set mSourceStudy = SourceStudy
mValueName = valueCache.ValueName
mValueNames(0) = mValueName

Set mResult = New ValueReplayResult
mResult.Initialise Target, SourceStudy, ValueNames, NotifyNames

Dim numValuesToNotify As Long
If TypeOf Target Is IStudy Then
    mTargetIsStudy = True
    Set mTargetStudy = Target
    Set mStudyToNotify = mTargetStudy.BaseStudy
    
    Dim lSessionBuilderFuture As IFuture
    Set lSessionBuilderFuture = pStudyManager.GetStudyEntry(mTargetStudy).SessionBuilderFuture
    
    If lSessionBuilderFuture.IsAvailable Then
        Set mSessionBuilder = lSessionBuilderFuture.Value
    Else
        Set mFutureWaiter = New FutureWaiter
        mFutureWaiter.Add lSessionBuilderFuture
    End If
    
    mNotifyName = mTargetStudy.StudyDefinition.StudyInputDefinitions.Item(1).Name
    mNotifyNames(0) = mNotifyName
    If numberOfValuesToReplay = ReplayNumbers.ReplayAsRequired Then
        ' just replay the number of values needed by the study
        numValuesToNotify = mTargetStudy.NumberOfBarsRequired
    Else
        numValuesToNotify = numberOfValuesToReplay
    End If
Else
    Set mTargetListener = Target
    numValuesToNotify = numberOfValuesToReplay
End If
If numValuesToNotify > mValueCache.Count Or _
    numValuesToNotify = ReplayNumbers.ReplayAll _
Then
    numValuesToNotify = mValueCache.Count
End If

mNextValueNumber = mValueCache.Count - numValuesToNotify + 1
mFirstValueNumber = mNextValueNumber

gLogger.Log "Study (" & mSourceStudy.Id & ") " & _
            mSourceStudy.InstanceName & _
            " replaying  values for " & mValueName, ProcName, ModuleName, LogLevelHighDetail

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub doCancel()
Const ProcName As String = "doCancel"
On Error GoTo Err

If Not mFutureWaiter Is Nothing Then mFutureWaiter.Cancel

mTaskContext.Progress 100 * mValuesNotified / (mSourceStudy.NumberOfCachedValues - mFirstValueNumber + 1), Empty
mTaskContext.Finish mResult, True
Set mTaskContext = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



