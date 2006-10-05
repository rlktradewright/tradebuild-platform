VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ATR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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

Private mParams As IParameters
Private mListeners As IStudyValueListeners
Private mStudies As IStudies

Private mInputValueName As String
Private mUnderlyingStudy As TradeBuildSP.IStudy

Private mPeriods As Long

Private mMA As MovingAverageStudy
Private mMAStudy As IStudy


Private mTR As Double
Private mATR As Variant
Private mPrevATR As Variant
Private mATRCache As TradeBuildSP.IStudyValueCache

Private mPrevClose As Variant   ' this is Variant type to cater for cases where
                                ' a bar is empty (eg the first bar has no previous bar)
Private mCurrHigh As Double
Private mCurrLow As Double
Private mCurrClose As Double

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mParams = GATR.defaultParameters
End Sub

Private Sub Class_Terminate()
Debug.Print "ATR study terminated"
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
Case UCase$(AtrValueATR)
    Set IStudy_addStudy = mATRCache.replayValues(study, Me, numUnderlyingValuesToUse, taskName, taskData)
    IStudy_addStudy.addTaskCompletionListener Me
End Select

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
GATR.defaultParameters = RHS
End Property

Private Property Get IStudy_defaultParameters() As TradeBuildSP.IParameters
Set IStudy_defaultParameters = GATR.defaultParameters
End Property

Private Function IStudy_getStudyValue( _
                ByVal ValueName As String, _
                ByVal ref As Long, _
                ByRef params() As Variant) As StudyValue
Select Case UCase$(ValueName)
Case UCase$(AtrValueATR)
    IStudy_getStudyValue = ATR(ref)
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
Set mATRCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, AtrValueATR)

End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = AtrShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams) & ")"
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = AtrShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams) & ")." & _
                mUnderlyingStudy.instancePath
End Property

Private Sub IStudy_Notify(ev As TradeBuildSP.StudyValueEvent)
Static numValues As Long
Dim isNewBar As Boolean
Dim newval As Double
Dim evMa As StudyValueEvent
Dim evOut As StudyValueEvent
Static prevVal As Variant

isNewBar = newBar(ev.barnumber)

If Not isNewBar Then
    If ev.value = prevVal Then Exit Sub
    prevVal = ev.value
End If

newval = CDbl(ev.value)

evMa = ev

If isNewBar Then
    mPrevClose = mCurrClose
    mCurrHigh = newval
    mCurrLow = newval
    mCurrClose = newval
    
    mTR = trueRange(newval)
    
    evMa.value = mTR
    mMAStudy.Notify evMa
    mATR = mMA.ma.value
    mATRCache.addValue mATR, ev.barnumber, ev.timestamp
Else
    If newval > mCurrHigh Then
        mCurrHigh = newval
    End If
    If newval < mCurrLow Then
        mCurrLow = newval
    End If
    mCurrClose = newval
    
    mTR = trueRange(newval)
    
    evMa.value = mTR
    mMAStudy.Notify evMa
    mATR = mMA.ma.value
    mATRCache.updateValue mATR, ev.barnumber, ev.timestamp
End If

Set evOut.Source = Me
evOut.barnumber = ev.barnumber
evOut.timestamp = ev.timestamp
evOut.value = mATR
evOut.ValueName = AtrValueATR
mStudies.Notify evOut
mListeners.Notify evOut
End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long
Dim thisNeeds As Long
Dim superiorStudiesNeeds
If mATRCache.cyclicSize = 0 Then
    IStudy_numberOfBarsRequired = 0
Else
    thisNeeds = mPeriods + 1 + mATRCache.cyclicSize
    superiorStudiesNeeds = mStudies.numberOfBarsRequired
    IStudy_numberOfBarsRequired = IIf(thisNeeds > superiorStudiesNeeds, _
                                    thisNeeds, _
                                    superiorStudiesNeeds)
End If
End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal ValueName As String = "") As Long
IStudy_numberOfCachedValues = mATRCache.Count
End Function

Private Property Get IStudy_parameters() As IParameters
Set IStudy_parameters = mParams.Clone
End Property

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = GATR.studyDefinition
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
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get ATR(Optional ByVal ref As Long) As StudyValue
ATR = mATRCache.getValue(ref)
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

Set mATRCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, AtrValueATR)

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
Dim maType As String

Set mParams = params.Clone

Set param = mParams.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(AtrParamPeriods)
        mPeriods = param.value
    Case UCase$(AtrParamMAType)
        maType = param.value
    End Select
    Set param = mParams.getNextParameter
Loop

Set mMA = gCreateMA(maType)
Set mMAStudy = mMA
Set lParams = mMAStudy.parameters
lParams.setParameterValue ParamPeriods, mPeriods
mMA.initialise mCommonServiceConsumer, mStudyServiceConsumer, 1, lParams

End Sub

Private Function trueRange(ByVal value As Variant) As Variant
Dim thisBarRange As Double
Dim prevCloseToHigh As Double
Dim prevCloseToLow As Double

If IsEmpty(value) Or IsEmpty(mPrevClose) Then Exit Function

thisBarRange = mCurrHigh - mCurrLow
prevCloseToHigh = mCurrHigh - mPrevClose
prevCloseToLow = mPrevClose - mCurrLow

trueRange = IIf(thisBarRange >= prevCloseToHigh, thisBarRange, prevCloseToHigh)
trueRange = IIf(trueRange >= prevCloseToLow, trueRange, prevCloseToLow)
End Function





