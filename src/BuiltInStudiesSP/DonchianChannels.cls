VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DonchianChannels"
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

Private mHighCache As IStudyValueCache
Private mLowcache As IStudyValueCache
Private mCurrHigh As Double
Private mCurrLow As Double

Private mUpper As Double
Private mLower As Double

Private mUpperCache As IStudyValueCache
Private mLowerCache As IStudyValueCache

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mParams = GDonchianChannels.defaultParameters
End Sub

Private Sub Class_Terminate()
Debug.Print "Donchian Channels study terminated"
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
Case UCase$(DoncValueLower)
    Set IStudy_addStudy = mLowerCache.replayValues( _
                                study, _
                                Me, _
                                numUnderlyingValuesToUse, _
                                TaskDiscriminators.TaskAddStudy, _
                                taskName, _
                                taskData)
    IStudy_addStudy.addTaskCompletionListener Me
Case UCase$(DoncValueUpper)
    Set IStudy_addStudy = mUpperCache.replayValues( _
                                study, _
                                Me, _
                                numUnderlyingValuesToUse, _
                                TaskDiscriminators.TaskAddStudyValueListener, _
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
Case UCase$(DoncValueLower)
    Set IStudy_addStudyValueListener = mLowerCache.replayValues(listener, _
                                Me, _
                                numberOfValuesToReplay, _
                                TaskDiscriminators.TaskAddStudyValueListener, _
                                taskName, _
                                taskData)
    IStudy_addStudyValueListener.addTaskCompletionListener Me
Case UCase$(DoncValueUpper)
    Set IStudy_addStudyValueListener = mUpperCache.replayValues(listener, _
                                Me, _
                                numberOfValuesToReplay, _
                                TaskDiscriminators.TaskAddStudyValueListener, _
                                taskName, _
                                taskData)
    IStudy_addStudyValueListener.addTaskCompletionListener Me
End Select
End Function

Private Property Get IStudy_baseStudy() As IStudy
Set IStudy_baseStudy = Me
End Property

Private Property Let IStudy_defaultParameters(ByVal RHS As TradeBuildSP.IParameters)
GDonchianChannels.defaultParameters = RHS
End Property

Private Property Get IStudy_defaultParameters() As TradeBuildSP.IParameters
Set IStudy_defaultParameters = GDonchianChannels.defaultParameters
End Property

Private Function IStudy_getStudyValue( _
                ByVal ValueName As String, _
                ByVal ref As Long, _
                ByRef params() As Variant) As StudyValue
Select Case UCase$(ValueName)
Case UCase$(DoncValueLower)
    IStudy_getStudyValue = lower(ref)
Case UCase$(DoncValueUpper)
    IStudy_getStudyValue = upper(ref)
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
Set mLowerCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, DoncValueLower)
Set mUpperCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, DoncValueUpper)

End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = DoncShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams) & ")"
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = DoncShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams) & ")." & _
                mUnderlyingStudy.instancePath
End Property

Private Sub IStudy_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
Dim isNewBar As Boolean
Static sPrevVal As Variant
Dim evOut As StudyValueEvent
Static numValues As Long
Dim i As Long

If IsEmpty(ev.value) Then Exit Sub

isNewBar = newBar(ev.barnumber)

If Not isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
    sPrevVal = ev.value
End If

If isNewBar Then
    numValues = numValues + 1
    mCurrHigh = ev.value
    mCurrLow = ev.value
    mHighCache.addValue mCurrHigh, ev.barnumber, ev.timestamp
    mLowcache.addValue mCurrLow, ev.barnumber, ev.timestamp
    If numValues >= mPeriods Then
        mUpper = DummyHigh
        For i = 1 To mPeriods - 1
            If mHighCache.getValue(i) > mUpper Then
                mUpper = mHighCache.getValue(i)
            End If
        Next
        mUpperCache.addValue mUpper, ev.barnumber, ev.timestamp
        notifyValue mUpper, DoncValueUpper, ev.barnumber, ev.timestamp
        
        mLower = DummyLow
        For i = 1 To mPeriods - 1
            If mLowcache.getValue(i) < mLower Then
                mLower = mLowcache.getValue(i)
            End If
        Next
        mLowerCache.addValue mLower, ev.barnumber, ev.timestamp
        notifyValue mLower, DoncValueLower, ev.barnumber, ev.timestamp
    End If
End If

If ev.value > mCurrHigh Then
    mCurrHigh = ev.value
    mHighCache.updateValue mCurrHigh, ev.barnumber, ev.timestamp
End If
If ev.value < mCurrLow Then
    mCurrLow = ev.value
    mLowcache.updateValue mCurrLow, ev.barnumber, ev.timestamp
End If
    
If numValues < mPeriods Then Exit Sub

If ev.value > mUpper Then
    mUpper = ev.value
    mUpperCache.updateValue mUpper, ev.barnumber, ev.timestamp
    notifyValue mUpper, DoncValueUpper, ev.barnumber, ev.timestamp
End If
If ev.value < mLower Then
    mLower = ev.value
    mLowerCache.updateValue mLower, ev.barnumber, ev.timestamp
    notifyValue mLower, DoncValueLower, ev.barnumber, ev.timestamp
End If
        
End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long
Dim thisNeeds As Long
Dim superiorStudiesNeeds
If mHighCache.cyclicSize = 0 Then
    IStudy_numberOfBarsRequired = 0
Else
    thisNeeds = mPeriods + mHighCache.cyclicSize
    superiorStudiesNeeds = mStudies.numberOfBarsRequired
    IStudy_numberOfBarsRequired = IIf(thisNeeds > superiorStudiesNeeds, _
                                thisNeeds, _
                                superiorStudiesNeeds)
End If
End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal ValueName As String = "") As Long
IStudy_numberOfCachedValues = mUpperCache.Count
End Function

Private Property Get IStudy_parameters() As IParameters
Set IStudy_parameters = mParams.Clone
End Property

Private Sub IStudy_removeStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
mListeners.Remove listener
End Sub

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = GDonchianChannels.studyDefinition
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
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get lower(Optional ByVal ref As Long) As Variant
If ref = 0 Then
    lower = mLower
Else
    lower = mLowerCache.getValue(ref)
End If
End Property

Public Property Get upper(Optional ByVal ref As Long) As Variant
If ref = 0 Then
    upper = mUpper
Else
    upper = mUpperCache.getValue(ref)
End If
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

Set mLowerCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, DoncValueLower)
Set mUpperCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, DoncValueUpper)

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

Private Sub notifyValue( _
                ByVal value As Variant, _
                ByVal ValueName As String, _
                ByVal barnumber As Long, _
                ByVal timestamp As Date)
Dim ev As StudyValueEvent
If Not IsEmpty(value) Then
    ev.value = value
    ev.ValueName = ValueName
    Set ev.Source = Me
    ev.barnumber = barnumber
    ev.timestamp = timestamp
    mStudies.Notify ev
    mListeners.Notify ev
End If
End Sub

Private Sub processParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter

Set mParams = params.Clone

Set param = mParams.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(DoncParamPeriods)
        mPeriods = param.value
        Set mHighCache = mCommonServiceConsumer.NewStudyValueCache(mPeriods, "highs")
        Set mLowcache = mCommonServiceConsumer.NewStudyValueCache(mPeriods, "lows")
    End Select
    Set param = mParams.getNextParameter
Loop

End Sub







