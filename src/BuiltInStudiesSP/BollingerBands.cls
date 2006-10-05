VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BollingerBands"
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

Private mParams As IParameters
Private mStudies As IStudies
Private mListeners As IStudyValueListeners

Private mInputValueName As String
Private mUnderlyingStudy As TradeBuildSP.IStudy

Private mPeriods As Long
Private mCentreBandWidth As Double
Private mEdgeBandWidth As Double
Private mSlopeThreshold As Double

Private mSD As StandardDeviation
Private mSDStudy As TradeBuildSP.IStudy

Private mMAStudy As TradeBuildSP.IStudy
Private mMA As MovingAverageStudy

Private mCentre As Variant
Private mTop As Variant
Private mBottom As Variant

Private mDevs As Double

Private mNumberOfValuesToCache As Long

Private mBottomCache As TradeBuildSP.IStudyValueCache
Private mCentreCache As TradeBuildSP.IStudyValueCache
Private mTopCache As TradeBuildSP.IStudyValueCache

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mSD = New StandardDeviation
Set mSDStudy = mSD
End Sub

Private Sub Class_Terminate()
Debug.Print "BollingerBands study terminated"
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
Case UCase$(BBValueBottom)
Case UCase$(BBValueCentre)
Case UCase$(BBValueSpread)
Case UCase$(BBValueTop)
Case Else
    Exit Function
End Select

Set IStudy_addStudy = mStudyServiceConsumer.replayStudyValues( _
                            study, _
                            Me, _
                            ValueName, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)
IStudy_addStudy.addTaskCompletionListener Me

End Function

Private Sub IStudy_addStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByRef valueNames() As String)
mListeners.Add listener, valueNames
End Sub

Private Property Get IStudy_baseStudy() As IStudy
Set IStudy_baseStudy = Me
End Property

Private Property Let IStudy_defaultParameters(ByVal RHS As TradeBuildSP.IParameters)
GBollingerBands.defaultParameters = RHS
End Property

Private Property Get IStudy_defaultParameters() As TradeBuildSP.IParameters
Set IStudy_defaultParameters = GBollingerBands.defaultParameters
End Property

Private Function IStudy_getStudyValue( _
                ByVal ValueName As String, _
                ByVal ref As Long, _
                params() As Variant) As StudyValue
Select Case UCase$(ValueName)
Case UCase$(BBValueBottom)
    IStudy_getStudyValue = Me.bottom(ref)
Case UCase$(BBValueCentre)
    IStudy_getStudyValue = Me.centre(ref)
Case UCase$(BBValueSpread)
    IStudy_getStudyValue = Me.spread(ref)
Case UCase$(BBValueTop)
    IStudy_getStudyValue = Me.top(ref)
Case "bottomfalling"
    If UBound(params) < 0 Then
        IStudy_getStudyValue = Me.bottomFalling(ref)
    Else
        IStudy_getStudyValue = Me.bottomFalling(ref, params(0))
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
mNumberOfValuesToCache = numberOfValuesToCache

End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = BbShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams) & ")"
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = BbShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams) & ")." & _
                mUnderlyingStudy.instancePath
End Property

Private Sub IStudy_Notify(ev As TradeBuildSP.StudyValueEvent)
Dim isNewBar As Boolean
Static sPrevVal As Variant
Dim evOut As TradeBuildSP.StudyValueEvent

isNewBar = newBar(ev.barnumber)

If Not isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
    sPrevVal = ev.value
End If

evOut.barnumber = ev.barnumber
Set evOut.Source = Me
evOut.timestamp = ev.timestamp
mSDStudy.Notify ev
mMAStudy.Notify ev

If IsEmpty(mMA.ma.value) Or IsEmpty(mSD.SD.value) Then Exit Sub

notifyValue mMA.ma.value, BBValueCentre, ev.barnumber, ev.timestamp
notifyValue mMA.ma.value + mDevs * mSD.SD.value, BBValueTop, ev.barnumber, ev.timestamp
notifyValue mMA.ma.value - mDevs * mSD.SD.value, BBValueBottom, ev.barnumber, ev.timestamp
notifyValue 2 * mDevs * mSD.SD.value, BBValueSpread, ev.barnumber, ev.timestamp
End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long
Dim lStudy As IStudy
Set lStudy = mSD
IStudy_numberOfBarsRequired = lStudy.numberOfBarsRequired
End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal ValueName As String = "") As Long
IStudy_numberOfCachedValues = mMAStudy.numberOfCachedValues
End Function

Private Property Get IStudy_parameters() As IParameters
Set IStudy_parameters = mParams.Clone
End Property

Private Sub IStudy_removeStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
mListeners.Remove listener
End Sub

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = GBollingerBands.studyDefinition
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
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get bottom(Optional ByVal ref As Long) As StudyValue
Dim sv As StudyValue
bottom = mMA.ma(ref)
sv = mSD.SD(ref)
If IsEmpty(bottom.value) Or IsEmpty(sv.value) Then Exit Property
bottom.value = bottom.value - mDevs * sv.value
End Property

Public Property Get bottomFalling(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As StudyValue
bottomFalling = bottomSlope(ref, periods)
If IsEmpty(bottomFalling.value) Then Exit Property
bottomFalling.value = (bottomFalling.value <= (-1 * mSlopeThreshold))
End Property

Public Property Get bottomFlat(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As StudyValue
bottomFlat = bottomSlope(ref, periods)
If IsEmpty(bottomFlat.value) Then Exit Property
bottomFlat.value = (bottomFlat.value < mSlopeThreshold And _
                    bottomFlat.value > (-1 * mSlopeThreshold))
End Property

Public Property Get bottomRising(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As StudyValue
bottomRising = bottomSlope(ref, periods)
If IsEmpty(bottomRising.value) Then Exit Property
bottomRising.value = (bottomRising.value >= mSlopeThreshold)
End Property

Public Property Get bottomSlope(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As StudyValue
Dim sv As StudyValue
bottomSlope = bottom(ref)
sv = bottom(ref - periods + 1)
If IsEmpty(bottomSlope.value) Or IsEmpty(sv.value) Then Exit Property
bottomSlope.value = (bottomSlope.value - sv.value) / periods
End Property

Public Property Get centre(Optional ByVal ref As Long) As StudyValue
centre = mMA.ma(ref)
End Property

Public Property Get centreFalling(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As StudyValue
centreFalling = centreSlope(ref, periods)
If IsEmpty(centreFalling.value) Then Exit Property
centreFalling.value = (centreFalling.value <= (-1 * mSlopeThreshold))
End Property

Public Property Get centreFlat(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As StudyValue
centreFlat = centreSlope(ref, periods)
If IsEmpty(centreFlat.value) Then Exit Property
centreFlat.value = (centreFlat.value < mSlopeThreshold And _
    centreFlat.value > (-1 * mSlopeThreshold))
End Property

Public Property Get centreRising(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As StudyValue
centreRising = centreSlope(ref, periods)
If IsEmpty(centreRising.value) Then Exit Property
centreRising.value = (centreRising.value >= mSlopeThreshold)
End Property

Public Property Get centreSlope(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As StudyValue
Dim sv As StudyValue
centreSlope = centre(ref)
sv = centre(ref - periods + 1)
If IsEmpty(centreSlope.value) Or IsEmpty(sv.value) Then Exit Property
centreSlope.value = (centreSlope.value - sv.value) / periods
End Property

Public Property Get inBottomBand(ByVal value As Double, _
                                Optional ByVal ref As Long) As StudyValue
inBottomBand = bottom(ref)
If IsEmpty(inBottomBand.value) Then Exit Property
inBottomBand.value = (value <= (inBottomBand.value + mEdgeBandWidth))
End Property

Public Property Get inCentreBand(ByVal value As Double, _
                                Optional ByVal ref As Long) As StudyValue
Dim bandUpper As Double
Dim bandLower As Double

inCentreBand = centre(ref)
If IsEmpty(inCentreBand.value) Then Exit Property
bandUpper = inCentreBand.value + (mCentreBandWidth / 2)
bandLower = bandUpper - mCentreBandWidth

inCentreBand.value = (value <= bandUpper _
                    And _
                    value >= bandLower)
End Property

Public Property Get inTopBand(ByVal value As Double, _
                            Optional ByVal ref As Long) As StudyValue
inTopBand = top(ref)
If IsEmpty(inTopBand.value) Then Exit Property
inTopBand.value = (value >= (inTopBand.value - mEdgeBandWidth))
End Property

Public Property Get rangeIntersectsCentreBand( _
                    ByVal lowValue As Double, _
                    ByVal highValue As Double, _
                    Optional ByVal ref As Long) As StudyValue
Dim bandUpper As Double
Dim bandLower As Double

rangeIntersectsCentreBand = centre(ref)
If IsEmpty(rangeIntersectsCentreBand.value) Then Exit Property

bandUpper = rangeIntersectsCentreBand.value + (mCentreBandWidth / 2)
bandLower = bandUpper - mCentreBandWidth

If highValue >= bandLower And highValue <= bandUpper Then
    rangeIntersectsCentreBand.value = True
ElseIf lowValue >= bandLower And lowValue <= bandUpper Then
    rangeIntersectsCentreBand.value = True
ElseIf lowValue < bandLower And highValue > bandUpper Then
    rangeIntersectsCentreBand.value = True
Else
    rangeIntersectsCentreBand.value = False
End If
End Property

Public Property Get spread(Optional ByVal ref As Long) As StudyValue
Dim sv As StudyValue
spread = top(ref)
sv = bottom(ref)
If IsEmpty(spread.value) Or IsEmpty(sv.value) Then Exit Property
spread.value = spread.value - sv.value
End Property

Public Property Get top(Optional ByVal ref As Long) As StudyValue
Dim sv As StudyValue
top = mMA.ma(ref)
sv = mSD.SD(ref)
If IsEmpty(top.value) Or IsEmpty(sv.value) Then Exit Property
top.value = top.value + mDevs * sv.value
End Property

Public Property Get topFalling(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As StudyValue
topFalling = topSlope(ref, periods)
If IsEmpty(topFalling.value) Then Exit Property
topFalling.value = (topFalling.value <= (-1 * mSlopeThreshold))
End Property

Public Property Get topFlat(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As StudyValue
topFlat = topSlope(ref, periods)
If IsEmpty(topFlat.value) Then Exit Property
topFlat.value = (topFlat.value < mSlopeThreshold And _
    topFlat.value > (-1 * mSlopeThreshold))
End Property

Public Property Get topRising(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As StudyValue
topRising = topSlope(ref, periods)
If IsEmpty(topRising.value) Then Exit Property
topRising.value = (topRising.value >= mSlopeThreshold)
End Property

Public Property Get topSlope(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As StudyValue
Dim sv As StudyValue
topSlope = top(ref)
sv = top(ref - periods + 1)
If IsEmpty(topSlope.value) Or IsEmpty(sv.value) Then Exit Property
topSlope.value = (topSlope.value - sv.value) / periods
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

Set mStudyServiceConsumer = studyServiceConsumer

mNumberOfValuesToCache = numberOfValuesToCache

If Not params Is Nothing Then
    mNumberOfValuesToCache = numberOfValuesToCache
    processParameters params
End If

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
Dim lParams As IParameters
Dim maType As String

Set mParams = params.Clone

Set param = mParams.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(BBParamCentreBandWidth)
        mCentreBandWidth = param.value
    Case UCase$(BBParamDeviations)
        mDevs = param.value
    Case UCase$(BBParamEdgeBandWidth)
        mEdgeBandWidth = param.value
    Case UCase$(BBParamPeriods)
        mPeriods = param.value
    Case UCase$(BBParamSlopeThreshold)
        mSlopeThreshold = param.value
    Case UCase$(BBParamMAType)
        maType = param.value
    End Select
    Set param = mParams.getNextParameter
Loop

Set mMA = gCreateMA(maType)
Set mMAStudy = mMA
Set lParams = mMAStudy.parameters
lParams.setParameterValue ParamPeriods, mPeriods
mMA.initialise mCommonServiceConsumer, mStudyServiceConsumer, mNumberOfValuesToCache, lParams

Set lParams = mSDStudy.parameters
lParams.setParameterValue ParamPeriods, mPeriods
mSD.initialise mCommonServiceConsumer, mStudyServiceConsumer, mNumberOfValuesToCache, lParams
End Sub


