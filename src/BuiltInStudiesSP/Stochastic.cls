VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stochastic"
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
Private mDPeriods As Long

Private mKCache As IStudyValueCache
Private mDCache As IStudyValueCache

Private mBars As Bars
Private mMASmoothing As SMA
Private mMASmoothingStudy As IStudy

Private mK As Variant
Private mD As Variant

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mParams = GStochastic.defaultParameters
End Sub

Private Sub Class_Terminate()
Debug.Print "Stochastic terminated"
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
Case UCase$(StochValueK)
    Set IStudy_addStudy = mKCache.replayValues(study, _
                                Me, _
                                numUnderlyingValuesToUse, _
                                TaskDiscriminators.TaskAddStudy, _
                                taskName, _
                                taskData)
    IStudy_addStudy.addTaskCompletionListener Me
Case UCase$(StochValueD)
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
Case UCase$(StochValueK)
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
Case UCase$(StochValueD)
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
Set IStudy_baseStudy = Me
End Property

Private Property Let IStudy_defaultParameters(ByVal RHS As TradeBuildSP.IParameters)
GStochastic.defaultParameters = RHS
End Property

Private Property Get IStudy_defaultParameters() As TradeBuildSP.IParameters
Set IStudy_defaultParameters = GStochastic.defaultParameters
End Property

Private Function IStudy_getStudyValue( _
                ByVal ValueName As String, _
                ByVal ref As Long, _
                ByRef params() As Variant) As StudyValue
Select Case UCase$(ValueName)
Case UCase$(DefaultStudyValueName), UCase$(StochValueK)
    IStudy_getStudyValue = percentK(ref)
Case UCase$(StochValueD)
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
Set mKCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, StochValueK)
Set mDCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, StochValueD)

End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = StochShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams, GStochastic.studyDefinition) & ")"
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = StochShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams, GStochastic.studyDefinition) & ")." & _
                mUnderlyingStudy.instancePath
End Property

Private Sub IStudy_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
Dim isNewBar As Boolean
Dim hh As Variant
Dim ll As Variant
Static sPrevVal As Variant
Dim newEv As TradeBuildSP.StudyValueEvent

isNewBar = newBar(ev.barnumber)

If Not isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
    sPrevVal = ev.value
End If

mBars.notify ev
'If ev.barNumber = 500 Then Stop
If mBars.Count > mKPeriods Then
    hh = mBars.highestHighValue(mKPeriods)
    ll = mBars.lowestLowValue(mKPeriods)
End If
If IsEmpty(hh) Or IsEmpty(ll) Then Exit Sub

If hh = ll Then
    ' leave the previous value for mK
Else
    mK = 100 * (mBars.closeValue - ll) / (hh - ll)
End If

newEv.barnumber = ev.barnumber
newEv.timestamp = ev.timestamp
newEv.value = mK
mMASmoothingStudy.notify newEv
mD = mMASmoothing.ma.value

If isNewBar Then
    mKCache.addValue mK, ev.barnumber, ev.timestamp
    mDCache.addValue mD, ev.barnumber, ev.timestamp
Else
    mKCache.updateValue mK, ev.barnumber, ev.timestamp
    mDCache.updateValue mD, ev.barnumber, ev.timestamp
End If
            
Set newEv.Source = Me
newEv.value = mK
newEv.ValueName = StochValueK
mStudies.notify newEv
mListeners.notify newEv
newEv.value = mD
newEv.ValueName = StochValueD
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
Set IStudy_studyDefinition = GStochastic.studyDefinition
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
percentD = mDCache.getStudyValue(ref)
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

Set mKCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, StochValueK)
Set mDCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, StochValueD)

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
    Case UCase$(StochParamKPeriods)
        mKPeriods = param.value
        Set mBars = New Bars
        mBars.initialise mCommonServiceConsumer, mKPeriods + 1
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

End Sub

