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
Private mUnderlyingStudy As TradeBuildSP.IStudy

Private mParams As IParameters
Private mStudies As IStudies
Private mListeners As IStudyValueListeners

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
Dim task As IValueReplayTask
Set task = ev.source
Select Case task.discriminator
Case TaskDiscriminators.TaskAddStudy
    Dim targetStudy As IStudy
    Set targetStudy = task.target
    mStudies.add targetStudy, task.valueNames
Case TaskDiscriminators.TaskAddStudyValueListener
    Dim targetListener As IStudyValueListener
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

Public Property Get InstanceName() As String
InstanceName = mShortName & "(" & _
                Join(mInputValueNames, ",") & "," & _
                ParamsToString & ")"
End Property

Public Property Get InstancePath() As String
InstancePath = mShortName & "(" & _
                Join(mInputValueNames, ",") & "," & _
                ParamsToString & ")." & _
                mUnderlyingStudy.InstancePath
End Property

Public Property Get isNewBar() As Boolean
isNewBar = mIsNewBar
End Property

Public Property Get localBarNumber() As Long
localBarNumber = mLocalBarNumber
End Property

Public Property Get parameters() As IParameters
Set parameters = mParams.clone
End Property

Public Property Let ReplayUsingStudyValues( _
                ByVal value As Boolean)
mReplayUsingStudyValues = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function AddStudy( _
                ByVal studyToAdd As IStudy, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As Tasks.TaskCompletion
Dim valueDefs As IStudyValueDefinitions
Dim valueDef As IStudyValueDefinition
Dim i As Long

If numUnderlyingValuesToUse = 0 Or study.numberOfCachedValues = 0 Then
    mStudies.add studyToAdd, valueNames
    Exit Function
End If

Set valueDefs = study.studyDefinition.studyValueDefinitions
If UBound(valueNames) = 0 And (Not mReplayUsingStudyValues) Then
    Dim valueName As String
    valueName = UCase$(valueNames(0))
    For Each valueDef In valueDefs
        If UCase$(valueDef.name) = valueName Then
            Set AddStudy = mValueCaches(i).valueCache.replayValues(studyToAdd, _
                                        study, _
                                        numUnderlyingValuesToUse, _
                                        TaskDiscriminators.TaskAddStudy, _
                                        taskName, _
                                        taskData)
            AddStudy.addTaskCompletionListener Me
        End If
        i = i + 1
    Next
Else
    Dim replayTask As New StudyValueReplayTask
    replayTask.initialise studyToAdd, _
                            study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            TaskDiscriminators.TaskAddStudy
    Set AddStudy = Tasks.startTask(replayTask, PriorityNormal, taskName, taskData)
    AddStudy.addTaskCompletionListener Me
    
End If
End Function

Public Function AddStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As Tasks.TaskCompletion
Dim valueDefs As IStudyValueDefinitions
Dim valueDef As IStudyValueDefinition
Dim valNames(0) As String

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
    Set AddStudyValueListener = Tasks.startTask(replayTask, PriorityNormal, taskName, taskData)
    AddStudyValueListener.addTaskCompletionListener Me
Else
    Set AddStudyValueListener = FindValueCache(valueName).replayValues( _
                                    listener, _
                                    study, _
                                    numberOfValuesToReplay, _
                                    TaskDiscriminators.TaskAddStudyValueListener, _
                                    taskName, _
                                    taskData)
    AddStudyValueListener.addTaskCompletionListener Me
End If
End Function

Public Function FindValueCache( _
                ByVal valueName As String) As IStudyValueCache
Dim i As Long

valueName = UCase$(valueName)
For i = 0 To UBound(mValueCaches)
    If UCase$(mValueCaches(i).valueCache.valueName) = valueName Or _
        (valueName = DefaultStudyValueName And _
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
                ByVal study As IStudy, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)

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
                ByRef ev As TradeBuildSP.StudyValueEvent)
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
                ByRef ev As TradeBuildSP.StudyValueEvent)
If mBarStartTime <> 0 Then ev.barStartTime = mBarStartTime
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

Public Sub RemoveStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
mListeners.remove listener
End Sub


'================================================================================
' Helper Functions
'================================================================================

Private Sub createValueCaches( _
                ByVal numberOfValuesToCache As Long)
Dim valueDefs As IStudyValueDefinitions
Dim valueDef As IStudyValueDefinition
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
Dim paramDefs As TradeBuildSP.IStudyParameterDefinitions
Dim paramDef As TradeBuildSP.IStudyParameterDefinition
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

Private Function study() As IStudy
Set study = mStudyRef.target
End Function

