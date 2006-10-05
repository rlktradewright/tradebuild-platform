VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMA"
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
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuildSP.IStudy
Implements MovingAverageStudy
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

Private mPeriods As Long
Private mExpFactor As Double
Private mSlopeThreshold As Double

Private mMaCache As TradeBuildSP.IStudyValueCache

Private mPrevMA As Variant  ' MA for the previous bar
Private mMA As Variant

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mParams = GEMA.defaultParameters
End Sub

Private Sub Class_Terminate()
Debug.Print "SimpleMovingAverageStudy terminated"
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


If ValueName = EMAValueEMA Then
    Set IStudy_addStudy = mMaCache.replayValues(study, Me, numUnderlyingValuesToUse, taskName, taskData)
    IStudy_addStudy.addTaskCompletionListener Me
End If

End Function

Private Sub IStudy_addStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                valueNames() As String)
mListeners.Add listener, valueNames
End Sub

Private Property Get IStudy_baseStudy() As IStudy
Set IStudy_baseStudy = Me
End Property

Private Property Let IStudy_defaultParameters(ByVal RHS As TradeBuildSP.IParameters)
GEMA.defaultParameters = RHS
End Property

Private Property Get IStudy_defaultParameters() As TradeBuildSP.IParameters
Set IStudy_defaultParameters = GEMA.defaultParameters
End Property

Private Function IStudy_getStudyValue( _
                ByVal ValueName As String, _
                ByVal ref As Long, _
                params() As Variant) As StudyValue
Select Case LCase$(ValueName)
Case UCase$(DefaultStudyValueName), UCase$(EMAValueEMA)
    IStudy_getStudyValue = ma(ref)
Case "FALLING"
    If UBound(params) < 0 Then
        IStudy_getStudyValue = falling(ref)
    Else
        IStudy_getStudyValue = falling(ref, params(0))
    End If
Case "FLAT"
    If UBound(params) < 0 Then
        IStudy_getStudyValue = flat(ref)
    Else
        IStudy_getStudyValue = flat(ref, params(0))
    End If
Case "RISING"
    If UBound(params) < 0 Then
        IStudy_getStudyValue = rising(ref)
    Else
        IStudy_getStudyValue = rising(ref, params(0))
    End If
Case "SLOPE"
    If UBound(params) < 0 Then
        IStudy_getStudyValue = slope(ref)
    Else
        IStudy_getStudyValue = slope(ref, params(0))
    End If
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
Set mMaCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, EMAValueEMA)

End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = EmaShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams) & ")"
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = EmaShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams) & ")." & _
                mUnderlyingStudy.instancePath
End Property

Private Sub IStudy_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
Dim isNewBar As Boolean
Static sPrevVal As Variant
Dim evOut As StudyValueEvent

isNewBar = newBar(ev.barnumber)

If Not isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
    sPrevVal = ev.value
End If

If isNewBar Then
    If IsEmpty(mPrevMA) Then
        mPrevMA = ev.value
    Else
        mPrevMA = mMA
    End If
End If

If Not IsEmpty(ev.value) Then
    mMA = (mExpFactor * ev.value) + mPrevMA * (1 - mExpFactor)
End If

If isNewBar Then
    mMaCache.addValue mMA, ev.barnumber, ev.timestamp
Else
    mMaCache.updateValue mMA, ev.barnumber, ev.timestamp
End If
            
Set evOut.Source = Me
evOut.barnumber = ev.barnumber
evOut.timestamp = ev.timestamp
evOut.value = mMA
evOut.ValueName = EMAValueEMA
mStudies.Notify evOut
mListeners.Notify evOut
End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long
Dim thisNeeds As Long
Dim superiorStudiesNeeds
If mMaCache.cyclicSize = 0 Then
    IStudy_numberOfBarsRequired = 0
Else
    thisNeeds = 2 * mPeriods + mMaCache.cyclicSize
    superiorStudiesNeeds = mStudies.numberOfBarsRequired
    IStudy_numberOfBarsRequired = IIf(thisNeeds > superiorStudiesNeeds, _
                                thisNeeds, _
                                superiorStudiesNeeds)
End If
End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal ValueName As String = "") As Long
IStudy_numberOfCachedValues = mMaCache.Count
End Function

Private Property Get IStudy_parameters() As IParameters
Set IStudy_parameters = mParams.Clone
End Property

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = GEMA.studyDefinition
End Property

'================================================================================
' ITaskCompletionListener Interface Members
'================================================================================

Private Sub ITaskCompletionListener_taskCompleted( _
                ev As TradeBuildSP.TaskCompletionEvent)
Dim task As IValueReplayTask
Set task = ev.Source
mStudies.Add task.targetStudy, task.ValueName
End Sub

'================================================================================
' MovingAverageStudy Interface Members
'================================================================================

Public Function MovingAverageStudy_addStudy( _
                study As IStudy, _
                ByVal ValueName As String, _
                ByVal numUnderlyingValuesToUse As Long) As TradeBuildSP.ITaskCompletion

Select Case UCase$(ValueName)
Case UCase$(EMAValueEMA)
    Set MovingAverageStudy_addStudy = mMaCache.replayValues(study, Me, numUnderlyingValuesToUse)
    MovingAverageStudy_addStudy.addTaskCompletionListener Me
End Select

End Function

Public Property Get MovingAverageStudy_falling( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As StudyValue
MovingAverageStudy_falling = falling(ref, periods)
End Property

Public Property Get MovingAverageStudy_flat( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As StudyValue
MovingAverageStudy_flat = flat(ref, periods)
End Property

Public Sub MovingAverageStudy_initialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal numberOfValuesToCache As Long, _
                ByVal params As IParameters)
initialise commonServiceConsumer, studyServiceConsumer, numberOfValuesToCache, params
End Sub

Public Property Get MovingAverageStudy_ma( _
                Optional ByVal ref As Long) As StudyValue
MovingAverageStudy_ma = ma(ref)
End Property

Private Function MovingAverageStudy_numberOfCachedValues( _
                Optional ByVal ValueName As String = "") As Long
MovingAverageStudy_numberOfCachedValues = mMaCache.Count
End Function

Public Property Get MovingAverageStudy_rising( _
                Optional ByVal ref As Long, _
                Optional ByVal periods As Long = 2) As StudyValue
MovingAverageStudy_rising = rising(ref, periods)
End Property

Public Property Get MovingAverageStudy_slope( _
                Optional ByVal ref As Long, _
                Optional ByVal periods As Long = 2) As StudyValue
MovingAverageStudy_slope = slope(ref, periods)
End Property

Public Property Let MovingAverageStudy_slopeThreshold( _
                ByVal value As Long)
mSlopeThreshold = value
End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get falling( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As StudyValue
Dim sv As StudyValue
sv = slope(ref, periods)
sv.value = (sv.value <= (-1 * mSlopeThreshold))
falling = sv
End Property

Public Property Get flat( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As StudyValue
Dim sv As StudyValue
sv = slope(ref, periods)
sv.value = (sv.value < mSlopeThreshold And _
    sv.value > (-1 * mSlopeThreshold))
flat = sv
End Property

Public Property Get ma(Optional ByVal ref As Long) As StudyValue
ma = mMaCache.getStudyValue(ref)
End Property

Public Property Get rising( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As StudyValue
Dim sv As StudyValue
sv = slope(ref, periods)
sv.value = (sv.value >= mSlopeThreshold)
rising = sv
End Property

Public Property Get slope( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As StudyValue
Dim sv As StudyValue
sv = ma(ref)
sv.value = (sv.value - ma(ref - periods + 1).value) / periods
slope = sv
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

Set mMaCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, EMAValueEMA)

If Not params Is Nothing Then processParameters params
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcExpFactor()
mExpFactor = 2 / (mPeriods + 1)
End Sub

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

Set mParams = params.Clone

Set param = mParams.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(EMAParamPeriods)
        mPeriods = param.value
        calcExpFactor
    Case UCase$(EMAParamSlopeThreshold)
        mSlopeThreshold = param.value
    End Select
    Set param = mParams.getNextParameter
Loop
        
End Sub







