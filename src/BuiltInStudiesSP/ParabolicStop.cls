VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParabolicStop"
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

Private mInitialStopForBar As Double
Private mCurrentStop As Double
Private mPsCache As IStudyValueCache
Private mExtremePoint As Double
Private mSignificantPoint As Double

Private mAccelerationFactor As Double
Private mStartFactor As Double
Private mIncrement As Double
Private mMaxFactor As Double

Private mLong As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mParams = GParabolicStop.defaultParameters
mCurrentStop = MaxDouble
mLong = True    ' arbitrary initial setting
End Sub

Private Sub Class_Terminate()
Debug.Print "Parabolic Stop terminated"
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
Case UCase$(PsValuePs)
    Set IStudy_addStudy = mPsCache.replayValues(study, _
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
Case UCase$(PsValuePs)
    If mPsCache.Count = 0 Then
        mListeners.Add listener, ValueName
    Else
        Set IStudy_addStudyValueListener = mPsCache.replayValues(listener, _
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
GParabolicStop.defaultParameters = RHS
End Property

Private Property Get IStudy_defaultParameters() As TradeBuildSP.IParameters
Set IStudy_defaultParameters = GParabolicStop.defaultParameters
End Property

Private Function IStudy_getStudyValue( _
                ByVal ValueName As String, _
                ByVal ref As Long, _
                ByRef params() As Variant) As StudyValue
Select Case UCase$(ValueName)
Case UCase$(DefaultStudyValueName), UCase$(PsValuePs)
    IStudy_getStudyValue = ps(ref)
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
Set mPsCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, PsValuePs)

End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = PsShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams, GParabolicStop.studyDefinition) & ")"
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = PsShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams, GParabolicStop.studyDefinition) & ")." & _
                mUnderlyingStudy.instancePath
End Property

Private Sub IStudy_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
Dim newValue As Double
Dim isNewBar As Boolean
Dim evOut As StudyValueEvent
Static accelAdjustedThisbar As Boolean
Static sPrevVal As Variant

If IsEmpty(ev.value) Then Exit Sub

isNewBar = newBar(ev.barnumber)

If Not isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
    sPrevVal = ev.value
End If

newValue = CDbl(ev.value)

If mCurrentStop = MaxDouble Then
    mCurrentStop = newValue * 0.999
    mExtremePoint = newValue * 1.01
    mInitialStopForBar = mCurrentStop
End If

If isNewBar Then
    accelAdjustedThisbar = False
    mInitialStopForBar = mCurrentStop
    mPsCache.addValue mCurrentStop, ev.barnumber, ev.timestamp
    If mAccelerationFactor = 0 Then
        mAccelerationFactor = mStartFactor
    End If
    If mLong Then
        mCurrentStop = mInitialStopForBar + mAccelerationFactor * (mExtremePoint - mInitialStopForBar)
    Else
        mCurrentStop = mInitialStopForBar - mAccelerationFactor * (mInitialStopForBar - mExtremePoint)
    End If
    mPsCache.updateValue mCurrentStop, ev.barnumber, ev.timestamp
End If

If mPsCache.Count < 2 Then Exit Sub
        
If mLong Then
    If newValue > mCurrentStop Then
        If newValue > mExtremePoint Then
            mExtremePoint = newValue
            If Not accelAdjustedThisbar And mAccelerationFactor <> 0 Then
                If Abs(mMaxFactor - mAccelerationFactor) > 0.001 Then
                    mAccelerationFactor = mAccelerationFactor + mIncrement
                    accelAdjustedThisbar = True
                End If
            End If
            mCurrentStop = mInitialStopForBar + mAccelerationFactor * (mExtremePoint - mInitialStopForBar)
            mPsCache.updateValue mCurrentStop, ev.barnumber, ev.timestamp
        End If
        If newValue > mSignificantPoint Then mSignificantPoint = newValue
    ElseIf newValue > mPsCache.getValue(-1) Then
        mCurrentStop = mPsCache.getValue(-1)
    Else
        mLong = False
        mCurrentStop = mSignificantPoint
        mInitialStopForBar = mCurrentStop
        mPsCache.updateValue mCurrentStop, ev.barnumber, ev.timestamp
        mAccelerationFactor = 0
        mExtremePoint = newValue
        mSignificantPoint = newValue
    End If
Else
    If newValue < mCurrentStop Then
        If newValue < mExtremePoint Then
            mExtremePoint = newValue
            If Not accelAdjustedThisbar And mAccelerationFactor <> 0 Then
                If Abs(mMaxFactor - mAccelerationFactor) > 0.001 Then
                    mAccelerationFactor = mAccelerationFactor + mIncrement
                    accelAdjustedThisbar = True
                End If
            End If
            mCurrentStop = mInitialStopForBar - mAccelerationFactor * (mInitialStopForBar - mExtremePoint)
            mPsCache.updateValue mCurrentStop, ev.barnumber, ev.timestamp
        End If
        If newValue < mSignificantPoint Then mSignificantPoint = newValue
    ElseIf newValue < mPsCache.getValue(-1) Then
        mCurrentStop = mPsCache.getValue(-1)
    Else
        mLong = True
        mCurrentStop = mSignificantPoint
        mInitialStopForBar = mCurrentStop
        mPsCache.updateValue mCurrentStop, ev.barnumber, ev.timestamp
        mAccelerationFactor = 0
        mExtremePoint = newValue
        mSignificantPoint = newValue
    End If
End If

Set evOut.Source = Me
evOut.barnumber = ev.barnumber
evOut.timestamp = ev.timestamp
evOut.value = mCurrentStop
evOut.ValueName = PsValuePs
mStudies.notify evOut
mListeners.notify evOut

End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long
Dim thisNeeds As Long
Dim superiorStudiesNeeds
thisNeeds = 100 ' difficult to say!
superiorStudiesNeeds = mStudies.numberOfBarsRequired
IStudy_numberOfBarsRequired = IIf(thisNeeds > superiorStudiesNeeds, _
                            thisNeeds, _
                            superiorStudiesNeeds)
End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal ValueName As String = "") As Long
IStudy_numberOfCachedValues = mPsCache.Count
End Function

Private Property Get IStudy_parameters() As IParameters
Set IStudy_parameters = mParams.Clone
End Property

Private Sub IStudy_removeStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
mListeners.Remove listener
End Sub

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = GParabolicStop.studyDefinition
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

Public Property Get ps(Optional ByVal ref As Integer) As StudyValue
ps = mPsCache.getStudyValue(ref)
End Property

Public Sub setStop(ByVal stopValue As Double, _
                    ByVal isLong As Boolean, _
                    ByVal extremePoint As Double)
mLong = isLong
mCurrentStop = stopValue
mInitialStopForBar = mCurrentStop
mPsCache.updateValue mCurrentStop
mAccelerationFactor = 0
mExtremePoint = extremePoint
mSignificantPoint = extremePoint
End Sub

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

Set mPsCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, PsValuePs)

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
    Case UCase$(PsParamStartFactor)
        mStartFactor = param.value
    Case UCase$(PsParamIncrement)
        mIncrement = param.value
    Case UCase$(PsParamMaxFactor)
        mMaxFactor = param.value
    End Select
    Set param = mParams.getNextParameter
Loop

End Sub


