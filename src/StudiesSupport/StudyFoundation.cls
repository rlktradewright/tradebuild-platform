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

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

Implements Tasks.TaskCompletionListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

Public Enum TaskDiscriminators
    TaskAddStudy
    TaskAddStudyValueListener
End Enum

Private Type ValueCacheTableEntry
    valueCache          As StudyValueCache
    isDefault           As Boolean
End Type

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mStudyRef As WeakReference
Private mShortName As String

Private mIdent As String

Private mInputValueNames() As String
Private mUnderlyingStudy As study

Private mParams As Parameters2.parameters
Private mStudies As Studies
Private mListeners As StudyValueListeners

Private mValueCaches() As ValueCacheTableEntry

Private mBarNumber As Long
Private mBarStartTime As Date
Private mLocalBarNumber As Long

Private mIsNewBar As Boolean

Private mReplayUsingStudyValues As Boolean

Private mNumberOfValuesToCache As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

End Sub

'================================================================================
' TaskCompletionListener Interface Members
'================================================================================

Private Sub TaskCompletionListener_taskCompleted( _
                ev As Tasks.TaskCompletionEvent)
Dim task As ValueReplayTask
Set task = ev.source
Select Case task.discriminator
Case TaskDiscriminators.TaskAddStudy
    Dim targetStudy As study
    Set targetStudy = task.target
    mStudies.add targetStudy, task.valueNames
Case TaskDiscriminators.TaskAddStudyValueListener
    Dim targetListener As StudyValueListener
    Set targetListener = task.target
    mListeners.add targetListener, task.valueName
End Select
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get barNumber() As Long
barNumber = mBarNumber
End Property

Public Property Get id() As String
id = mIdent
End Property

Public Property Get instanceName() As String
instanceName = mShortName & "(" & _
                Join(mInputValueNames, ",") & "," & _
                ParamsToString & ")"
End Property

Public Property Get instancePath() As String
instancePath = mShortName & "(" & _
                Join(mInputValueNames, ",") & "," & _
                ParamsToString & ")." & _
                mUnderlyingStudy.instancePath
End Property

Public Property Get isNewBar() As Boolean
isNewBar = mIsNewBar
End Property

Public Property Get localBarNumber() As Long
localBarNumber = mLocalBarNumber
End Property

Public Property Get parameters() As Parameters2.parameters
Set parameters = mParams.clone
End Property

Public Property Let ReplayUsingStudyValues( _
                ByVal value As Boolean)
mReplayUsingStudyValues = value
End Property

Public Property Get sessionEndTime() As Date
sessionEndTime = mUnderlyingStudy.sessionEndTime
End Property

Public Property Get sessionStartTime() As Date
sessionStartTime = mUnderlyingStudy.sessionStartTime
End Property

'================================================================================
' Methods
'================================================================================

Public Function addStudy( _
                ByVal studyToAdd As study, _
                ByRef valueNames() As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As Tasks.TaskCompletion
Dim valueDefs As studyValueDefinitions
Dim valueDef As StudyValueDefinition
Dim i As Long

If numberOfValuesToReplay >= 0 Or _
    numberOfValuesToReplay = ReplayNumbers.ReplayAsRequired Or _
    numberOfValuesToReplay = ReplayNumbers.ReplayAll _
Then
    ' ok
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "StudiesSupp" & "." & "StudyFoundation" & ":" & "addStudy", _
            "numberOfValuesToReplay must be >=0 or set to ReplayNumbers.ReplayAll"
    
End If

If numberOfValuesToReplay = 0 Or study.numberOfCachedValues = 0 Then
    mStudies.add studyToAdd, valueNames
    Exit Function
End If

Set valueDefs = study.studyDefinition.studyValueDefinitions
If UBound(valueNames) = 0 And (Not mReplayUsingStudyValues) Then
    Dim valueName As String
    valueName = UCase$(valueNames(0))
    For Each valueDef In valueDefs
        If UCase$(valueDef.name) = valueName Then
            Set addStudy = mValueCaches(i).valueCache.replayValues(studyToAdd, _
                                        study, _
                                        numberOfValuesToReplay, _
                                        TaskDiscriminators.TaskAddStudy, _
                                        taskName, _
                                        taskData)
            addStudy.addTaskCompletionListener Me
        End If
        i = i + 1
    Next
Else
    Dim replayTask As New StudyValueReplayTask
    replayTask.initialise studyToAdd, _
                            study, _
                            valueNames, _
                            numberOfValuesToReplay, _
                            TaskDiscriminators.TaskAddStudy
    Set addStudy = Tasks.startTask(replayTask, PriorityNormal, taskName, taskData)
    addStudy.addTaskCompletionListener Me
    
End If
End Function

Public Function addStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As Tasks.TaskCompletion
Dim valueDefs As studyValueDefinitions
Dim valNames(0) As String

If numberOfValuesToReplay >= 0 Or _
    numberOfValuesToReplay = ReplayNumbers.ReplayAll _
Then
    ' ok
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "StudiesSupp" & "." & "StudyFoundation" & ":" & "addStudyValueListener", _
            "numberOfValuesToReplay must be >=0 or set to ReplayNumbers.ReplayAll"
    
End If

If numberOfValuesToReplay = 0 Or study.numberOfCachedValues = 0 Then
    mListeners.add listener, valueName
    Exit Function
End If
    
Set valueDefs = study.studyDefinition.studyValueDefinitions
valueName = UCase$(valueName)
If mReplayUsingStudyValues Then
    Dim replayTask As New StudyValueReplayTask
    valNames(0) = valueName
    replayTask.initialise listener, _
                            study, _
                            valNames, _
                            numberOfValuesToReplay, _
                            TaskDiscriminators.TaskAddStudyValueListener
    Set addStudyValueListener = Tasks.startTask(replayTask, PriorityNormal, taskName, taskData)
    addStudyValueListener.addTaskCompletionListener Me
Else
    Set addStudyValueListener = FindValueCache(valueName).replayValues( _
                                    listener, _
                                    study, _
                                    numberOfValuesToReplay, _
                                    TaskDiscriminators.TaskAddStudyValueListener, _
                                    taskName, _
                                    taskData)
    addStudyValueListener.addTaskCompletionListener Me
End If
End Function

Public Function FindValueCache( _
                ByVal valueName As String) As StudyValueCache
Dim i As Long

valueName = UCase$(valueName)
For i = 0 To UBound(mValueCaches)
    If UCase$(mValueCaches(i).valueCache.valueName) = valueName Or _
        (valueName = DefaultStudyValueNameStr And _
        mValueCaches(i).isDefault) _
    Then
        Set FindValueCache = mValueCaches(i).valueCache
        Exit For
    End If
Next
End Function

Public Function getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As StudyValue
Dim valueCache As StudyValueCache

Set valueCache = FindValueCache(valueName)
getStudyValue = valueCache.getStudyValue(ref)
End Function

Public Sub initialise( _
                ByVal study As study, _
                ByVal id As String, _
                ByVal parameters As Parameters2.parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As study)

Set mStudyRef = WeakRefs.createWeakReference(study)
mShortName = study.studyDefinition.shortName
mIdent = id
Set mParams = parameters
mInputValueNames = inputValueNames
Set mUnderlyingStudy = underlyingStudy
Set mStudies = New Studies
Set mListeners = New StudyValueListeners

mNumberOfValuesToCache = numberOfValuesToCache
createValueCaches mNumberOfValuesToCache

End Sub

Public Sub NotifyInput( _
                ByRef ev As StudyValueEvent)
If ev.barNumber <> mBarNumber Then
    mBarNumber = ev.barNumber
    mLocalBarNumber = mLocalBarNumber + 1
    mBarStartTime = ev.barStartTime
    mIsNewBar = True
Else
    mIsNewBar = False
End If
End Sub
                
Public Sub NotifyValue( _
                ByRef ev As StudyValueEvent)
If ev.barStartTime = 0 Then ev.barStartTime = mBarStartTime
mStudies.notify ev
mListeners.notify ev
End Sub
                
Public Function numberOfBarsRequired( _
                ByVal studyNeeds As Long) As Long
Dim superiorStudiesNeeds
studyNeeds = studyNeeds + mNumberOfValuesToCache
superiorStudiesNeeds = mStudies.numberOfBarsRequired
numberOfBarsRequired = IIf(studyNeeds > superiorStudiesNeeds, _
                            studyNeeds, _
                            superiorStudiesNeeds)
End Function

Public Function numberOfCachedValues( _
                Optional ByVal valueName As String = "") As Long
If valueName = "" Then
    numberOfCachedValues = mValueCaches(0).valueCache.count
Else
    numberOfCachedValues = FindValueCache(valueName).count
End If
End Function

Public Sub removeStudyValueListener( _
                ByVal listener As StudyValueListener)
mListeners.remove listener
End Sub


'================================================================================
' Helper Functions
'================================================================================

Private Sub createValueCaches( _
                ByVal numberOfValuesToCache As Long)
Dim valueDefs As studyValueDefinitions
Dim valueDef As StudyValueDefinition
Dim i As Long

Set valueDefs = study.studyDefinition.studyValueDefinitions

ReDim mValueCaches(valueDefs.count - 1) As ValueCacheTableEntry
For Each valueDef In valueDefs
    Set mValueCaches(i).valueCache = New StudyValueCache
    mValueCaches(i).valueCache.initialise numberOfValuesToCache, valueDef.name
    mValueCaches(i).isDefault = valueDef.isDefault
    i = i + 1
Next
End Sub

Private Function ParamsToString() As String
Dim paramDefs As studyParameterDefinitions
Dim paramDef As StudyParameterDefinition
Dim i As Long

On Error Resume Next
Set paramDefs = study.studyDefinition.studyParameterDefinitions
For i = 1 To paramDefs.count
    Set paramDef = paramDefs.Item(i)
    If Len(ParamsToString) = 0 Then
        ParamsToString = mParams.getParameterValue(paramDef.name)
    Else
        ParamsToString = ParamsToString & "," & mParams.getParameterValue(paramDef.name)
    End If
Next
End Function

Private Function study() As study
Set study = mStudyRef.target
End Function

