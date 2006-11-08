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

Implements ITaskCompletionListener

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
    valueCache          As TradeBuildSP.IStudyValueCache
    isDefault           As Boolean
End Type

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mStudyServiceConsumer As IStudyServiceConsumer

Private mStudyRef As IWeakReference
Private mShortName As String

Private mIdent As String

Private mInputValueNames() As String
Private mUnderlyingStudy As TradeBuildSP.IStudy

Private mParams As IParameters
Private mStudies As IStudies
Private mListeners As IStudyValueListeners

Private mValueCaches() As ValueCacheTableEntry

Private mBarNumber As Long
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
' XXXX Interface Members
'================================================================================

Private Sub ITaskCompletionListener_taskCompleted( _
                ev As TradeBuildSP.TaskCompletionEvent)
Dim task As IValueReplayTask
Set task = ev.Source
Select Case task.discriminator
Case TaskDiscriminators.TaskAddStudy
    Dim targetStudy As IStudy
    Set targetStudy = task.Target
    mStudies.Add targetStudy, task.ValueNames
Case TaskDiscriminators.TaskAddStudyValueListener
    Dim targetListener As IStudyValueListener
    Set targetListener = task.Target
    mListeners.Add targetListener, task.ValueName
End Select
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get barnumber() As Long
barnumber = mBarNumber
End Property

Public Property Let commonServiceConsumer( _
                ByVal value As ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property

Public Property Get Id() As String
Id = mIdent
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

Public Property Get Parameters() As IParameters
Set Parameters = mParams.Clone
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
                ByRef ValueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion
Dim valueDefs As IStudyValueDefinitions
Dim valueDef As IStudyValueDefinition
Dim i As Long

If numUnderlyingValuesToUse = 0 Or study.NumberOfCachedValues = 0 Then
    mStudies.Add studyToAdd, ValueNames
    Exit Function
End If

Set valueDefs = study.studyDefinition.StudyValueDefinitions
If UBound(ValueNames) = 0 And (Not mReplayUsingStudyValues) Then
    Dim ValueName As String
    ValueName = UCase$(ValueNames(0))
    For Each valueDef In valueDefs
        If UCase$(valueDef.Name) = ValueName Then
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
    Set AddStudy = mStudyServiceConsumer.replayStudyValues( _
                                studyToAdd, _
                                study, _
                                ValueNames, _
                                numUnderlyingValuesToUse, _
                                TaskDiscriminators.TaskAddStudy, _
                                taskName, _
                                taskData)
    AddStudy.addTaskCompletionListener Me
    
End If
End Function

Public Function AddStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByVal ValueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion
Dim valueDefs As IStudyValueDefinitions
Dim valueDef As IStudyValueDefinition
Dim valNames(0) As String

If numberOfValuesToReplay = 0 Or study.NumberOfCachedValues = 0 Then
    mListeners.Add listener, ValueName
    Exit Function
End If
    
Set valueDefs = study.studyDefinition.StudyValueDefinitions
ValueName = UCase$(ValueName)
If mReplayUsingStudyValues Then
    valNames(0) = ValueName
    Set AddStudyValueListener = mStudyServiceConsumer.replayStudyValues( _
                                    listener, _
                                    study, _
                                    valNames, _
                                    numberOfValuesToReplay, _
                                    TaskDiscriminators.TaskAddStudyValueListener, _
                                    taskName, _
                                    taskData)
    AddStudyValueListener.addTaskCompletionListener Me
Else
    Set AddStudyValueListener = FindValueCache(ValueName).replayValues( _
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
                ByVal ValueName As String) As IStudyValueCache
Dim i As Long

ValueName = UCase$(ValueName)
For i = 0 To UBound(mValueCaches)
    If UCase$(mValueCaches(i).valueCache.ValueName) = ValueName Or _
        (ValueName = DefaultStudyValueName And _
        mValueCaches(i).isDefault) _
    Then
        Set FindValueCache = mValueCaches(i).valueCache
        Exit For
    End If
Next
End Function

Public Function GetStudyValue( _
                ByVal ValueName As String, _
                ByVal ref As Long) As StudyValue
Dim valueCache As IStudyValueCache

Set valueCache = FindValueCache(ValueName)
GetStudyValue = valueCache.GetStudyValue(ref)
End Function

Public Sub initialise( _
                ByVal study As IStudy, _
                ByVal Id As String, _
                ByVal Parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy, _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer)

Set mCommonServiceConsumer = commonServiceConsumer
Set mStudyServiceConsumer = studyServiceConsumer

Set mStudyRef = mCommonServiceConsumer.NewWeakReference(study)
mShortName = study.studyDefinition.shortName
mIdent = Id
Set mParams = Parameters
mInputValueNames = inputValueNames
Set mUnderlyingStudy = underlyingStudy
Set mStudies = mCommonServiceConsumer.NewStudies
Set mListeners = mCommonServiceConsumer.NewStudyValueListeners

mNumberOfValuesToCache = numberOfValuesToCache
createValueCaches mNumberOfValuesToCache

End Sub

Public Sub NotifyInput( _
                ByRef ev As TradeBuildSP.StudyValueEvent)
mIsNewBar = newBar(ev.barnumber)
End Sub
                
Public Sub NotifyValue( _
                ByRef ev As TradeBuildSP.StudyValueEvent)
mStudies.notify ev
mListeners.notify ev
End Sub
                
Public Function NumberOfBarsRequired( _
                ByVal studyNeeds As Long) As Long
Dim superiorStudiesNeeds
studyNeeds = studyNeeds + mNumberOfValuesToCache
superiorStudiesNeeds = mStudies.NumberOfBarsRequired
NumberOfBarsRequired = IIf(studyNeeds > superiorStudiesNeeds, _
                            studyNeeds, _
                            superiorStudiesNeeds)
End Function

Public Function NumberOfCachedValues( _
                Optional ByVal ValueName As String = "") As Long
If ValueName = "" Then
    NumberOfCachedValues = mValueCaches(0).valueCache.Count
Else
    NumberOfCachedValues = FindValueCache(ValueName).Count
End If
End Function

Public Sub RemoveStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
mListeners.Remove listener
End Sub


'================================================================================
' Helper Functions
'================================================================================

Private Sub createValueCaches( _
                ByVal numberOfValuesToCache As Long)
Dim valueDefs As IStudyValueDefinitions
Dim valueDef As IStudyValueDefinition
Dim i As Long

Set valueDefs = study.studyDefinition.StudyValueDefinitions

ReDim mValueCaches(valueDefs.Count - 1) As ValueCacheTableEntry
For Each valueDef In valueDefs
    Set mValueCaches(i).valueCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, valueDef.Name)
    mValueCaches(i).isDefault = valueDef.isDefault
    i = i + 1
Next
End Sub

Private Function ParamsToString() As String
Dim paramDefs As TradeBuildSP.IStudyParameterDefinitions
Dim paramDef As TradeBuildSP.IStudyParameterDefinition
Dim i As Long

On Error Resume Next
Set paramDefs = study.studyDefinition.StudyParameterDefinitions
For i = 1 To paramDefs.Count
    Set paramDef = paramDefs.Item(i)
    If Len(ParamsToString) = 0 Then
        ParamsToString = mParams.getParameterValue(paramDef.Name)
    Else
        ParamsToString = ParamsToString & "," & mParams.getParameterValue(paramDef.Name)
    End If
Next
End Function

Private Function newBar(barnumber As Long) As Boolean
If barnumber <> mBarNumber Then
    mBarNumber = barnumber
    mLocalBarNumber = mLocalBarNumber + 1
    newBar = True
Else
    newBar = False
End If
End Function

Private Function study() As IStudy
Set study = mStudyRef.Target
End Function

