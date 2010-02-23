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

Implements TaskCompletionListener

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

Public Enum TaskDiscriminators
    TaskAddStudy
    TaskAddStudyValueListener
End Enum

Private Type ValueCacheTableEntry
    valueCache          As valueCache
    IsDefault           As Boolean
End Type

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mStudyRef As WeakReference
Private mShortName As String

Private mIdent As String

Private mInputValueNames() As String
Private mUnderlyingStudy As Study

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

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

End Sub

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_taskCompleted( _
                ev As TaskCompletionEvent)
Dim tc As TaskController
Dim result As ValueReplayResult
Const ProcName As String = "TaskCompletionListener_taskCompleted"
On Error GoTo Err

Set result = ev.result
Set tc = ev.source
Select Case CLng(tc.cookie)
Case TaskDiscriminators.TaskAddStudy
    mStudies.Add result.target, result.valueNames
Case TaskDiscriminators.TaskAddStudyValueListener
    mListeners.Add result.target, result.valueName
End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Id() As String
Const ProcName As String = "Id"
On Error GoTo Err

Id = mIdent

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get InstanceName() As String
Const ProcName As String = "InstanceName"
On Error GoTo Err

InstanceName = mShortName & "(" & _
                Join(mInputValueNames, ",") & "," & _
                ParamsToString & ")"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get InstancePath() As String
Const ProcName As String = "InstancePath"
On Error GoTo Err

InstancePath = mShortName & "(" & _
                Join(mInputValueNames, ",") & "," & _
                ParamsToString & ")." & _
                mUnderlyingStudy.InstancePath

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IsNewBar() As Boolean
Const ProcName As String = "IsNewBar"
On Error GoTo Err

IsNewBar = mIsNewBar

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LocalBarNumber() As Long
Const ProcName As String = "LocalBarNumber"
On Error GoTo Err

LocalBarNumber = mLocalBarNumber

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Parameters() As Parameters
Const ProcName As String = "Parameters"
On Error GoTo Err

Set Parameters = mParams.Clone

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ReplayUsingStudyValues( _
                ByVal value As Boolean)
Const ProcName As String = "ReplayUsingStudyValues"
On Error GoTo Err

mReplayUsingStudyValues = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AddStudy( _
                ByVal studyToAdd As Study, _
                ByRef valueNames() As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Dim notifyNames() As String
Dim replayValueNames() As String
Dim replayNotifyNames() As String
Dim inputDefs As StudyInputDefinitions
Dim lBarStudy As BarStudy
Dim i As Long
Dim j As Long

Const ProcName As String = "AddStudy"
On Error GoTo Err

Set inputDefs = studyToAdd.StudyDefinition.StudyInputDefinitions

ReDim notifyNames(inputDefs.Count - 1) As String

For i = 0 To inputDefs.Count - 1
    notifyNames(i) = inputDefs.Item(i + 1).Name
Next

If TypeOf Study Is BarStudy Then
    Set lBarStudy = Study
    
    ' if the study being added needs close values, we must make sure that
    ' open, high and low are also replayed
    
    ReDim replayValueNames(UBound(valueNames)) As String
    ReDim replayNotifyNames(UBound(valueNames)) As String
    Set inputDefs = studyToAdd.StudyDefinition.StudyInputDefinitions
    
    Dim CloseValueName  As String
    CloseValueName = UCase$(lBarStudy.CloseValueName)
    For i = 0 To UBound(valueNames)
        Dim valueName As String
        valueName = UCase$(valueNames(i))
        If valueName = CloseValueName Or _
            valueName = DefaultStudyValueNameStr _
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
            replayValueNames(j) = valueNames(i)
            replayNotifyNames(j) = inputDefs.Item(i + 1).Name
            j = j + 1
        End If
    Next
    
    Set AddStudy = addStudyHelper(studyToAdd, _
                            valueNames, _
                            notifyNames, _
                            replayValueNames, _
                            replayNotifyNames, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)
Else
    Set AddStudy = addStudyHelper(studyToAdd, _
                            valueNames, _
                            notifyNames, _
                            valueNames, _
                            notifyNames, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)
End If
    

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Public Function AddStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Dim valueDefs As StudyValueDefinitions
Dim valNames(0) As String
Dim notifyNames(0) As String 'dummy

Const ProcName As String = "AddStudyValueListener"
On Error GoTo Err

If numberOfValuesToReplay >= 0 Or _
    numberOfValuesToReplay = ReplayNumbers.ReplayAll _
Then
    ' ok
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & "StudyFoundation" & ":" & "AddStudyValueListener", _
            "numberOfValuesToReplay must be >=0 or set to ReplayNumbers.ReplayAll"
    
End If

If numberOfValuesToReplay = 0 Or Study.NumberOfCachedValues = 0 Then
    mListeners.Add listener, valueName
    Exit Function
End If
    
Set valueDefs = Study.StudyDefinition.StudyValueDefinitions
valueName = UCase$(valueName)
If mReplayUsingStudyValues Then
    Dim replayTask As New StudyValueReplayTask
    valNames(0) = valueName
    replayTask.initialise listener, _
                            Study, _
                            valNames, _
                            notifyNames, _
                            numberOfValuesToReplay, _
                            valNames, _
                            notifyNames, _
                            0
    Set AddStudyValueListener = StartTask(replayTask, PriorityNormal, taskName, TaskDiscriminators.TaskAddStudyValueListener, taskData)
    AddStudyValueListener.AddTaskCompletionListener Me
Else
    Set AddStudyValueListener = replayCacheValues(FindValueCache(valueName), _
                                    listener, _
                                    Study, _
                                    numberOfValuesToReplay, _
                                    TaskDiscriminators.TaskAddStudyValueListener, _
                                    taskName, _
                                    taskData)
    AddStudyValueListener.AddTaskCompletionListener Me
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function FindValueCache( _
                ByVal valueName As String) As valueCache
Dim i As Long

Const ProcName As String = "FindValueCache"
On Error GoTo Err

valueName = UCase$(valueName)
For i = 0 To UBound(mValueCaches)
    If UCase$(mValueCaches(i).valueCache.valueName) = valueName Or _
        (valueName = DefaultStudyValueNameStr And _
        mValueCaches(i).IsDefault) _
    Then
        Set FindValueCache = mValueCaches(i).valueCache
        Exit For
    End If
Next

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function GetStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Dim valueCache As valueCache

Const ProcName As String = "GetStudyValue"
On Error GoTo Err

Set valueCache = FindValueCache(valueName)
GetStudyValue = valueCache.GetSValue(ref)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub initialise( _
                ByVal Study As Study, _
                ByVal Id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study)

Const ProcName As String = "initialise"
On Error GoTo Err

Set mStudyRef = CreateWeakReference(Study)
mShortName = Study.StudyDefinition.ShortName
mIdent = Id
Set mParams = Parameters
mInputValueNames = inputValueNames
Set mUnderlyingStudy = underlyingStudy

Set mStudies = New Studies
mStudies.DefaultValueName = Study.StudyDefinition.DefaultValueName

Set mListeners = New StudyValueListeners
mListeners.DefaultValueName = Study.StudyDefinition.DefaultValueName

mNumberOfValuesToCache = numberOfValuesToCache
createValueCaches mNumberOfValuesToCache

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Public Sub NotifyInput( _
                ByRef ev As StudyValueEvent)
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub
                
Public Sub NotifyValue( _
                ByRef ev As StudyValueEvent)
Const ProcName As String = "NotifyValue"
On Error GoTo Err

If ev.sVal.BarStartTime = 0 Then ev.sVal.BarStartTime = mBarStartTime
mStudies.Notify ev
mListeners.Notify ev

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub
                
Public Function NumberOfBarsRequired( _
                ByVal studyNeeds As Long) As Long
Dim superiorStudiesNeeds
Const ProcName As String = "NumberOfBarsRequired"
On Error GoTo Err

studyNeeds = studyNeeds + mNumberOfValuesToCache
superiorStudiesNeeds = mStudies.NumberOfBarsRequired
NumberOfBarsRequired = IIf(studyNeeds > superiorStudiesNeeds, _
                            studyNeeds, _
                            superiorStudiesNeeds)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function NumberOfCachedValues( _
                Optional ByVal valueName As String = "") As Long
Const ProcName As String = "NumberOfCachedValues"
On Error GoTo Err

If valueName = "" Then
    NumberOfCachedValues = mValueCaches(0).valueCache.Count
Else
    NumberOfCachedValues = FindValueCache(valueName).Count
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub RemoveStudyValueListener( _
                ByVal listener As StudyValueListener)
Const ProcName As String = "RemoveStudyValueListener"
On Error GoTo Err

mListeners.Remove listener

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Function addStudyHelper( _
                ByVal studyToAdd As Study, _
                ByRef valueNames() As String, _
                ByRef notifyNames() As String, _
                ByRef replayValueNames() As String, _
                ByRef replayNotifyNames() As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Dim valueDefs As StudyValueDefinitions
Dim valueDef As StudyValueDefinition
Dim i As Long

Const ProcName As String = "addStudyHelper"
On Error GoTo Err

If numberOfValuesToReplay >= 0 Or _
    numberOfValuesToReplay = ReplayNumbers.ReplayAsRequired Or _
    numberOfValuesToReplay = ReplayNumbers.ReplayAll _
Then
    ' ok
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & "StudyFoundation" & ":" & "AddStudy", _
            "numberOfValuesToReplay must be >=0 or set to ReplayNumbers.ReplayAll"
    
End If

If numberOfValuesToReplay = 0 Or Study.NumberOfCachedValues = 0 Then
    mStudies.AddEx studyToAdd, valueNames, notifyNames
    Exit Function
End If

Set valueDefs = Study.StudyDefinition.StudyValueDefinitions
If UBound(valueNames) = 0 And (Not mReplayUsingStudyValues) Then
    Dim valueName As String
    valueName = UCase$(valueNames(0))
    For Each valueDef In valueDefs
        If UCase$(valueDef.Name) = valueName Or _
            (valueName = DefaultStudyValueNameStr And _
            valueDef.IsDefault) _
        Then
            Set addStudyHelper = replayCacheValues(mValueCaches(i).valueCache, _
                                        studyToAdd, _
                                        Study, _
                                        numberOfValuesToReplay, _
                                        TaskDiscriminators.TaskAddStudy, _
                                        taskName, _
                                        taskData)
            addStudyHelper.AddTaskCompletionListener Me
            Exit For
        End If
        i = i + 1
    Next
Else
    Dim replayTask As New StudyValueReplayTask
    replayTask.initialise studyToAdd, _
                            Study, _
                            replayValueNames, _
                            replayNotifyNames, _
                            numberOfValuesToReplay, _
                            valueNames, _
                            notifyNames, _
                            0
    Set addStudyHelper = StartTask(replayTask, PriorityNormal, taskName, TaskDiscriminators.TaskAddStudy, taskData)
    addStudyHelper.AddTaskCompletionListener Me
    
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
                
End Function

Private Sub createValueCaches( _
                ByVal numberOfValuesToCache As Long)
Dim valueDefs As StudyValueDefinitions
Dim valueDef As StudyValueDefinition
Dim i As Long

Const ProcName As String = "createValueCaches"
On Error GoTo Err

Set valueDefs = Study.StudyDefinition.StudyValueDefinitions

ReDim mValueCaches(valueDefs.Count - 1) As ValueCacheTableEntry
For Each valueDef In valueDefs
    Set mValueCaches(i).valueCache = CreateValueCache(numberOfValuesToCache, valueDef.Name)
    mValueCaches(i).IsDefault = valueDef.IsDefault
    i = i + 1
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function ParamsToString() As String
Dim paramDefs As StudyParameterDefinitions
Dim paramDef As StudyParameterDefinition
Dim i As Long

Const ProcName As String = "ParamsToString"
On Error GoTo Err

On Error Resume Next
Set paramDefs = Study.StudyDefinition.StudyParameterDefinitions
For i = 1 To paramDefs.Count
    Set paramDef = paramDefs.Item(i)
    If Len(ParamsToString) = 0 Then
        ParamsToString = mParams.GetParameterValue(paramDef.Name)
    Else
        ParamsToString = ParamsToString & "," & mParams.GetParameterValue(paramDef.Name)
    End If
Next

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function replayCacheValues( _
                ByVal cache As valueCache, _
                ByVal target As Object, _
                ByVal sourceStudy As Study, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal discriminator As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Dim replayTask As New CacheReplayTask

Const ProcName As String = "replayCacheValues"
On Error GoTo Err

If numberOfValuesToReplay >= 0 Or _
    numberOfValuesToReplay = ReplayNumbers.ReplayAll _
Then
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & "." & ":" & "replayValues", _
            "numberOfValuesToReplay must be >=0 or set to ReplayNumbers.ReplayAll"

End If
replayTask.initialise cache, _
                    target, _
                    sourceStudy, _
                    numberOfValuesToReplay, _
                    0

Set replayCacheValues = StartTask( _
                                replayTask, _
                                PriorityNormal, _
                                IIf(taskName <> "", taskName, GenerateGUIDString), _
                                discriminator, _
                                taskData)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function Study() As Study
Const ProcName As String = "Study"
On Error GoTo Err

Set Study = mStudyRef.target

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

