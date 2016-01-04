VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements ITaskCompletionListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "StudyFoundation"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type ValueCacheTableEntry
    valueCache          As valueCache
    IsDefault           As Boolean
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mName As String
Private mLibraryName As String

Private mStudyManagerRef As WeakReference

Private mStudyRef As WeakReference
Private mShortName As String

Private mIdent As String

Private mInputValueNames() As String
Private mUnderlyingStudyRef As WeakReference

Private mParams As Parameters
Private mStudies As Studies
Private mListeners As StudyValueListeners

Private mValueCaches() As ValueCacheTableEntry

Private mBarNumber As Long
Private mBarStartTime As Date
Private mLocalBarNumber As Long

Private mIsNewBar As Boolean

Private mReplayUsingStudyValues As Boolean

Private mNumberOfValuesToCache As Long

Private mSessionFuture As IFuture

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITaskCompletionListener Interface Members
'@================================================================================

Private Sub ITaskCompletionListener_taskCompleted( _
                ev As TaskCompletionEventData)
Const ProcName As String = "ITaskCompletionListener_taskCompleted"
On Error GoTo Err

Dim result As ValueReplayResult
Set result = ev.result

Dim tc As TaskController
Set tc = ev.Source
Select Case CLng(tc.Cookie)
Case TaskDiscriminators.TaskAddStudy
    mStudies.Add result.Target, result.ValueNames
Case TaskDiscriminators.TaskAddStudyValueListener
    mListeners.Add result.Target, result.ValueName
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get BarNumber() As Long
Const ProcName As String = "BarNumber"
On Error GoTo Err

BarNumber = mBarNumber

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Id() As String
Const ProcName As String = "Id"
On Error GoTo Err

Id = mIdent

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get InputValueNames() As String()
InputValueNames = mInputValueNames
End Property

Public Property Get InstanceName() As String
Const ProcName As String = "InstanceName"
On Error GoTo Err

InstanceName = mShortName & "(" & _
                Join(mInputValueNames, ",") & "," & _
                ParamsToString & ")"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get InstancePath() As String
Const ProcName As String = "InstancePath"
On Error GoTo Err

InstancePath = mShortName & "(" & _
                Join(mInputValueNames, ",") & "," & _
                ParamsToString & ")." & _
                UnderlyingStudy.InstancePath

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsNewBar() As Boolean
Const ProcName As String = "IsNewBar"
On Error GoTo Err

IsNewBar = mIsNewBar

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LibraryName() As String
LibraryName = mLibraryName
End Property

Public Property Get LocalBarNumber() As Long
Const ProcName As String = "LocalBarNumber"
On Error GoTo Err

LocalBarNumber = mLocalBarNumber

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Get Parameters() As Parameters
Const ProcName As String = "Parameters"
On Error GoTo Err

Set Parameters = mParams.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ReplayUsingStudyValues( _
                ByVal Value As Boolean)
Const ProcName As String = "ReplayUsingStudyValues"
On Error GoTo Err

mReplayUsingStudyValues = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SessionFuture() As IFuture
Set SessionFuture = mSessionFuture
End Property

Public Property Get StudyManager() As StudyManager
Set StudyManager = mStudyManagerRef.Target
End Property

Public Property Get UnderlyingStudy() As IStudy
Set UnderlyingStudy = mUnderlyingStudyRef.Target
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AddStudy( _
                ByVal studyToAdd As IStudy, _
                ByRef ValueNames() As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal TaskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "AddStudy"
On Error GoTo Err

Dim inputDefs As StudyInputDefinitions
Set inputDefs = studyToAdd.StudyDefinition.StudyInputDefinitions

Dim NotifyNames() As String
ReDim NotifyNames(inputDefs.Count - 1) As String

Dim i As Long
For i = 0 To inputDefs.Count - 1
    NotifyNames(i) = inputDefs.Item(i + 1).Name
Next

If TypeOf Study Is IBarStudy Then
    Dim lBarStudy As IBarStudy
    Set lBarStudy = Study
    
    ' if the study being added needs close values, we must make sure that
    ' open, high and low are also replayed
    
    Dim replayValueNames() As String
    ReDim replayValueNames(UBound(ValueNames)) As String
    
    Dim replayNotifyNames() As String
    ReDim replayNotifyNames(UBound(ValueNames)) As String
    Set inputDefs = studyToAdd.StudyDefinition.StudyInputDefinitions
    
    Dim CloseValueName  As String
    CloseValueName = UCase$(lBarStudy.CloseValueName)

    Dim j As Long
    
    For i = 0 To UBound(ValueNames)
        Dim ValueName As String
        ValueName = UCase$(ValueNames(i))
        If ValueName = CloseValueName Or _
            ValueName = DefaultStudyValueNameStr _
        Then
            ReDim Preserve replayValueNames(UBound(replayValueNames) + 3) As String
            ReDim Preserve replayNotifyNames(UBound(replayValueNames)) As String
            
            replayValueNames(j) = lBarStudy.OpenValueName
            replayNotifyNames(j) = inputDefs.Item(i + 1).Name
            
            replayValueNames(j + 1) = lBarStudy.HighValueName
            replayNotifyNames(j + 1) = inputDefs.Item(i + 1).Name
            
            replayValueNames(j + 2) = lBarStudy.LowValueName
            replayNotifyNames(j + 2) = inputDefs.Item(i + 1).Name
            
            replayValueNames(j + 3) = lBarStudy.CloseValueName
            replayNotifyNames(j + 3) = inputDefs.Item(i + 1).Name
            
            j = j + 4
        Else
            replayValueNames(j) = ValueNames(i)
            replayNotifyNames(j) = inputDefs.Item(i + 1).Name
            j = j + 1
        End If
    Next
    
    Set AddStudy = addStudyHelper(studyToAdd, _
                            ValueNames, _
                            NotifyNames, _
                            replayValueNames, _
                            replayNotifyNames, _
                            numberOfValuesToReplay, _
                            TaskName, _
                            taskData)
Else
    Set AddStudy = addStudyHelper(studyToAdd, _
                            ValueNames, _
                            NotifyNames, _
                            ValueNames, _
                            NotifyNames, _
                            numberOfValuesToReplay, _
                            TaskName, _
                            taskData)
End If
    

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function AddStudyValueListener( _
                ByVal listener As IStudyValueListener, _
                ByVal ValueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal TaskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "AddStudyValueListener"
On Error GoTo Err

AssertArgument numberOfValuesToReplay >= 0 Or _
    numberOfValuesToReplay = ReplayNumbers.ReplayAll, "numberOfValuesToReplay must be >=0 or set to ReplayNumbers.ReplayAll"

If numberOfValuesToReplay = 0 Or Study.NumberOfCachedValues = 0 Then
    mListeners.Add listener, ValueName
    Exit Function
End If
    
Dim valueDefs As StudyValueDefinitions
Set valueDefs = Study.StudyDefinition.StudyValueDefinitions
ValueName = UCase$(ValueName)
If mReplayUsingStudyValues Then
    Dim valNames(0) As String
    valNames(0) = ValueName
    
    Dim NotifyNames(0) As String 'dummy
    Dim replayTask As New StudyValueReplayTask
    replayTask.Initialise StudyManager, _
                            listener, _
                            Study, _
                            valNames, _
                            NotifyNames, _
                            numberOfValuesToReplay, _
                            valNames, _
                            NotifyNames
    Set AddStudyValueListener = StartTask(replayTask, PriorityNormal, TaskName, TaskDiscriminators.TaskAddStudyValueListener, taskData)
    AddStudyValueListener.AddTaskCompletionListener Me
Else
    Set AddStudyValueListener = replayCacheValues(FindValueCache(ValueName), _
                                    listener, _
                                    Study, _
                                    numberOfValuesToReplay, _
                                    TaskDiscriminators.TaskAddStudyValueListener, _
                                    TaskName, _
                                    taskData)
    AddStudyValueListener.AddTaskCompletionListener Me
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FindValueCache( _
                ByVal ValueName As String) As valueCache
Const ProcName As String = "FindValueCache"
On Error GoTo Err

ValueName = UCase$(ValueName)

Dim i As Long
For i = 0 To UBound(mValueCaches)
    If UCase$(mValueCaches(i).valueCache.ValueName) = ValueName Or _
        (ValueName = DefaultStudyValueNameStr And _
        mValueCaches(i).IsDefault) _
    Then
        Set FindValueCache = mValueCaches(i).valueCache
        Exit For
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetStudyValue( _
                ByVal ValueName As String, _
                ByVal ref As Long) As SValue
Const ProcName As String = "GetStudyValue"
On Error GoTo Err
Dim UnderlyingStudy
Dim valueCache As valueCache
Set valueCache = FindValueCache(ValueName)
GetStudyValue = valueCache.GetSValue(ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Initialise( _
                ByVal pName As String, _
                ByVal pLibraryName As String, _
                ByVal pStudyManager As StudyManager, _
                ByVal Study As IStudy, _
                ByVal Id As String, _
                ByVal Parameters As Parameters, _
                ByVal pNumberOfValuesToCache As Long, _
                ByRef InputValueNames() As String, _
                ByVal pUnderlyingStudy As IStudy, _
                ByVal pSessionFuture As IFuture)
Const ProcName As String = "Initialise"
On Error GoTo Err

mName = pName
mLibraryName = pLibraryName

Set mStudyManagerRef = CreateWeakReference(pStudyManager)
Set mStudyRef = CreateWeakReference(Study)
mShortName = Study.StudyDefinition.ShortName
mIdent = Id
Set mParams = Parameters
mInputValueNames = InputValueNames
Set mUnderlyingStudyRef = CreateWeakReference(pUnderlyingStudy)

Set mStudies = New Studies
mStudies.DefaultValueName = Study.StudyDefinition.DefaultValueName

Set mListeners = New StudyValueListeners
mListeners.DefaultValueName = Study.StudyDefinition.DefaultValueName

mNumberOfValuesToCache = pNumberOfValuesToCache
createValueCaches mNumberOfValuesToCache

Set mSessionFuture = pSessionFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Public Sub NotifyInput( _
                ByRef ev As StudyValueEventData)
Const ProcName As String = "NotifyInput"
On Error GoTo Err

If ev.sVal.BarNumber <> mBarNumber Then
    mBarNumber = ev.sVal.BarNumber
    mLocalBarNumber = mLocalBarNumber + 1
    mBarStartTime = ev.sVal.BarStartTime
    mIsNewBar = True
Else
    mIsNewBar = False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
Public Sub NotifyValue( _
                ByRef ev As StudyValueEventData)
Const ProcName As String = "NotifyValue"
On Error GoTo Err

If ev.sVal.BarStartTime = 0 Then ev.sVal.BarStartTime = mBarStartTime
mStudies.Notify ev
mListeners.Notify ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
Public Function NumberOfBarsRequired( _
                ByVal studyNeeds As Long) As Long
Const ProcName As String = "NumberOfBarsRequired"
On Error GoTo Err

studyNeeds = studyNeeds + mNumberOfValuesToCache

Dim superiorStudiesNeeds
superiorStudiesNeeds = mStudies.NumberOfBarsRequired
NumberOfBarsRequired = IIf(studyNeeds > superiorStudiesNeeds, _
                            studyNeeds, _
                            superiorStudiesNeeds)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NumberOfCachedValues( _
                Optional ByVal ValueName As String = "") As Long
Const ProcName As String = "NumberOfCachedValues"
On Error GoTo Err

If ValueName = "" Then
    NumberOfCachedValues = mValueCaches(0).valueCache.Count
Else
    NumberOfCachedValues = FindValueCache(ValueName).Count
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub RemoveStudyValueListener( _
                ByVal listener As IStudyValueListener)
Const ProcName As String = "RemoveStudyValueListener"
On Error GoTo Err

mListeners.Remove listener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Function addStudyHelper( _
                ByVal studyToAdd As IStudy, _
                ByRef ValueNames() As String, _
                ByRef NotifyNames() As String, _
                ByRef replayValueNames() As String, _
                ByRef replayNotifyNames() As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal TaskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "addStudyHelper"
On Error GoTo Err

AssertArgument numberOfValuesToReplay >= 0 Or _
    numberOfValuesToReplay = ReplayNumbers.ReplayAsRequired Or _
    numberOfValuesToReplay = ReplayNumbers.ReplayAll, "numberOfValuesToReplay must be >=0 or set to ReplayNumbers.ReplayAll or ReplayNumbers.ReplayAsRequired"

If numberOfValuesToReplay = 0 Or Study.NumberOfCachedValues = 0 Then
    mStudies.AddEx studyToAdd, ValueNames, NotifyNames
    Exit Function
End If

Dim valueDefs As StudyValueDefinitions
Set valueDefs = Study.StudyDefinition.StudyValueDefinitions

If UBound(ValueNames) = 0 And (Not mReplayUsingStudyValues) Then
    Dim ValueName As String
    ValueName = UCase$(ValueNames(0))
    
    Dim i As Long

    Dim valueDef As StudyValueDefinition
    For Each valueDef In valueDefs
        If UCase$(valueDef.Name) = ValueName Or _
            (ValueName = DefaultStudyValueNameStr And _
            valueDef.IsDefault) _
        Then
            Set addStudyHelper = replayCacheValues(mValueCaches(i).valueCache, _
                                        studyToAdd, _
                                        Study, _
                                        numberOfValuesToReplay, _
                                        TaskDiscriminators.TaskAddStudy, _
                                        TaskName, _
                                        taskData)
            addStudyHelper.AddTaskCompletionListener Me
            Exit For
        End If
        i = i + 1
    Next
Else
    Dim replayTask As New StudyValueReplayTask
    replayTask.Initialise StudyManager, _
                            studyToAdd, _
                            Study, _
                            replayValueNames, _
                            replayNotifyNames, _
                            numberOfValuesToReplay, _
                            ValueNames, _
                            NotifyNames
    Set addStudyHelper = StartTask(replayTask, PriorityNormal, TaskName, TaskDiscriminators.TaskAddStudy, taskData)
    addStudyHelper.AddTaskCompletionListener Me
    
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub createValueCaches( _
                ByVal numberOfValuesToCache As Long)
Const ProcName As String = "createValueCaches"
On Error GoTo Err

Dim valueDefs As StudyValueDefinitions
Set valueDefs = Study.StudyDefinition.StudyValueDefinitions

ReDim mValueCaches(valueDefs.Count - 1) As ValueCacheTableEntry

Dim i As Long
Dim valueDef As StudyValueDefinition
For Each valueDef In valueDefs
    Set mValueCaches(i).valueCache = CreateValueCache(numberOfValuesToCache, valueDef.Name)
    mValueCaches(i).IsDefault = valueDef.IsDefault
    i = i + 1
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function ParamsToString() As String
Const ProcName As String = "ParamsToString"
On Error GoTo Err

On Error Resume Next

Dim paramDefs As StudyParameterDefinitions
Set paramDefs = Study.StudyDefinition.StudyParameterDefinitions

Dim i As Long
For i = 1 To paramDefs.Count
    Dim paramDef As StudyParameterDefinition
    Set paramDef = paramDefs.Item(i)
    If Len(ParamsToString) = 0 Then
        ParamsToString = mParams.GetParameterValue(paramDef.Name)
    Else
        ParamsToString = ParamsToString & "," & mParams.GetParameterValue(paramDef.Name)
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function replayCacheValues( _
                ByVal cache As valueCache, _
                ByVal Target As Object, _
                ByVal SourceStudy As IStudy, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal Discriminator As Long, _
                Optional ByVal TaskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "replayCacheValues"
On Error GoTo Err

AssertArgument numberOfValuesToReplay >= 0 Or _
    numberOfValuesToReplay = ReplayNumbers.ReplayAll, "numberOfValuesToReplay must be >=0 or set to ReplayNumbers.ReplayAll"

Dim replayTask As New CacheReplayTask
replayTask.Initialise StudyManager, _
                    cache, _
                    Target, _
                    SourceStudy, _
                    numberOfValuesToReplay, _
                    0

Set replayCacheValues = StartTask( _
                                replayTask, _
                                PriorityNormal, _
                                IIf(TaskName <> "", TaskName, GenerateGUIDString), _
                                Discriminator, _
                                taskData)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function Study() As IStudy
Const ProcName As String = "Study"
On Error GoTo Err

Set Study = mStudyRef.Target

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

