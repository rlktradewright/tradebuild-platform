VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlowStochastic"
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
' Interfaces
'================================================================================

Implements TradeBuildSP.IStudy
Implements TradeBuildSP.ITaskCompletionListener

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

Private mStudyServiceConsumer As IStudyServiceConsumer
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mIdent As String

Private mInputValueName As String
Private mUnderlyingStudy As TradeBuildSP.IStudy

Private mParams As IParameters
Private mStudies As IStudies
Private mListeners As IStudyValueListeners

Private mKPeriods As Long
Private mKDPeriods As Long
Private mDPeriods As Long

Private mK As Variant
Private mD As Variant

Private mKCache As IStudyValueCache
Private mDCache As IStudyValueCache

Private mStoch As Stochastic
Private mStochStudy As TradeBuildSP.IStudy
Private mMASmoothing As SMA
Private mMASmoothingStudy As TradeBuildSP.IStudy

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mParams = GSlowStochastic.defaultParameters
End Sub

Private Sub Class_Terminate()
Debug.Print "SlowStochastic terminated"
End Sub

'================================================================================
' IStudy Interface Members
'================================================================================

Private Function IStudy_addStudy( _
                ByVal study As IStudy, _
                ByVal ValueName As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion

Select Case UCase$(ValueName)
Case UCase$(SStochValueK)
    Set IStudy_addStudy = mKCache.replayValues(study, _
                                Me, _
                                numUnderlyingValuesToUse, _
                                TaskDiscriminators.TaskAddStudy, _
                                taskName, _
                                taskData)
    IStudy_addStudy.addTaskCompletionListener Me
Case UCase$(SStochValueD)
    Set IStudy_addStudy = mDCache.replayValues(study, _
                                Me, _
                                numUnderlyingValuesToUse, _
                                TaskDiscriminators.TaskAddStudy, _
                                taskName, _
                                taskData)
    IStudy_addStudy.addTaskCompletionListener Me
End Select

End Function

Private Function IStudy_addStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByVal ValueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion
Select Case UCase$(ValueName)
Case UCase$(SStochValueK)
    If mKCache.Count = 0 Then
        mListeners.Add listener, ValueName
    Else
        Set IStudy_addStudyValueListener = mKCache.replayValues(listener, _
                                    Me, _
                                    numberOfValuesToReplay, _
                                    TaskDiscriminators.TaskAddStudyValueListener, _
                                    taskName, _
                                    taskData)
        IStudy_addStudyValueListener.addTaskCompletionListener Me
    End If
Case UCase$(SStochValueD)
    If mDCache.Count = 0 Then
        mListeners.Add listener, ValueName
    Else
        Set IStudy_addStudyValueListener = mDCache.replayValues(listener, _
                                    Me, _
                                    numberOfValuesToReplay, _
                                    TaskDiscriminators.TaskAddStudyValueListener, _
                                    taskName, _
                                    taskData)
        IStudy_addStudyValueListener.addTaskCompletionListener Me
    End If
End Select
End Function

Private Property Get IStudy_baseStudy() As IStudy
Set IStudy_baseStudy = mStochStudy
End Property

Private Property Let IStudy_defaultParameters(ByVal RHS As TradeBuildSP.IParameters)
GStochastic.defaultParameters = RHS
End Property

Private Property Get IStudy_defaultParameters() As TradeBuildSP.IParameters
Set IStudy_defaultParameters = GSlowStochastic.defaultParameters
End Property

Private Function IStudy_getStudyValue( _
                ByVal ValueName As String, _
                ByVal ref As Long, _
                ByRef params() As Variant) As StudyValue
Select Case UCase$(ValueName)
Case UCase$(DefaultStudyValueName), UCase$(SStochValueK)
    IStudy_getStudyValue = percentK(ref)
Case UCase$(SStochValueD)
    IStudy_getStudyValue = percentD(ref)
End Select
End Function

Private Property Get IStudy_id() As String
IStudy_id = mIdent
End Property

Private Sub IStudy_initialise( _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByVal inputValueName As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)
mIdent = id
mInputValueName = inputValueName
Set mUnderlyingStudy = underlyingStudy
processParameters parameters
Set mKCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, SStochValueK)
Set mDCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, SStochValueD)

End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = SStochShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams, GSlowStochastic.studyDefinition) & ")"
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = SStochShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams, GSlowStochastic.studyDefinition) & ")." & _
                mUnderlyingStudy.instancePath
End Property

Private Sub IStudy_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
Dim isNewBar As Boolean
Static sPrevVal As Variant
Dim newEv As TradeBuildSP.StudyValueEvent

isNewBar = newBar(ev.barnumber)

If Not isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
    sPrevVal = ev.value
End If

mK = ev.value
mMASmoothingStudy.notify ev
mD = mMASmoothing.ma.value

newEv.barnumber = ev.barnumber
newEv.timestamp = ev.timestamp
Set newEv.Source = Me
newEv.value = mK
newEv.ValueName = SStochValueK
mStudies.notify newEv
mListeners.notify newEv
newEv.value = mD
newEv.ValueName = SStochValueD
mStudies.notify newEv
mListeners.notify newEv
End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long
Dim thisNeeds As Long
Dim superiorStudiesNeeds
thisNeeds = mKPeriods + mDPeriods + mKCache.cyclicSize
superiorStudiesNeeds = mStudies.numberOfBarsRequired
IStudy_numberOfBarsRequired = IIf(thisNeeds > superiorStudiesNeeds, _
                            thisNeeds, _
                            superiorStudiesNeeds)
End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal ValueName As String = "") As Long
IStudy_numberOfCachedValues = mKCache.Count
End Function

Private Property Get IStudy_parameters() As IParameters
Set IStudy_parameters = mParams.Clone
End Property

Private Sub IStudy_removeStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
mListeners.Remove listener
End Sub

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = GSlowStochastic.studyDefinition
End Property

'================================================================================
' ITaskCompletionListener Interface Members
'================================================================================

Private Sub ITaskCompletionListener_taskCompleted( _
                ev As TradeBuildSP.TaskCompletionEvent)
Dim task As IValueReplayTask
Set task = ev.Source
Select Case task.discriminator
Case TaskDiscriminators.TaskAddStudy
    Dim targetStudy As IStudy
    Set targetStudy = task.Target
    mStudies.Add targetStudy, task.ValueName
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

Public Property Get percentD(Optional ByVal ref As Integer) As Variant
percentD = mKCache.getStudyValue(ref)
End Property

Public Property Get percentK(Optional ByVal ref As Integer) As Variant
percentK = mKCache.getStudyValue(ref)
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub initialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                Optional ByVal numberOfValuesToCache As Long, _
                Optional ByVal params As IParameters)

Set mCommonServiceConsumer = commonServiceConsumer
Set mStudyServiceConsumer = studyServiceConsumer
Set mStudies = mCommonServiceConsumer.NewStudies
Set mListeners = mCommonServiceConsumer.NewStudyValueListeners

Set mKCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, SStochValueK)
Set mDCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, SStochValueD)

If Not params Is Nothing Then processParameters params
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function newBar(barnumber As Long) As Boolean
Static previousBarNumber As Long
If barnumber <> previousBarNumber Then
    previousBarNumber = barnumber
    newBar = True
Else
    newBar = False
End If
End Function

Private Sub processParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter
Dim lParams As IParameters

Set mParams = params.Clone

Set param = mParams.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(SStochParamKPeriods)
        mKPeriods = param.value
    Case UCase$(SStochParamKDPeriods)
        mKDPeriods = param.value
    Case UCase$(StochParamDPeriods)
        mDPeriods = param.value
        Set mMASmoothing = New SMA
        Set mMASmoothingStudy = mMASmoothing
        Set lParams = mMASmoothingStudy.defaultParameters
        lParams.setParameterValue ParamPeriods, mDPeriods
        mMASmoothing.initialise mCommonServiceConsumer, mStudyServiceConsumer, 1, lParams
    End Select
    Set param = mParams.getNextParameter
Loop

Set mStoch = New Stochastic
Set mStochStudy = mStoch
Set lParams = mStochStudy.defaultParameters
lParams.setParameterValue StochParamKPeriods, mKPeriods
lParams.setParameterValue StochParamDPeriods, mKDPeriods
mStoch.initialise mCommonServiceConsumer, mStudyServiceConsumer, 1, lParams
mStochStudy.addStudy Me, StochValueD, 1
End Sub





