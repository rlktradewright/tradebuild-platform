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

'-------this variable is required in all studies---------
Private mStudyFoundation As StudyFoundation
'--------------------------------------------------------

Private mStudyServiceConsumer As IStudyServiceConsumer
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mPeriods As Long
Private mCentreBandWidth As Double
Private mEdgeBandWidth As Double
Private mSlopeThreshold As Double

Private mSD As StandardDeviation
Private mSDStudy As TradeBuildSP.IStudy

Private mMAStudy As TradeBuildSP.IStudy
Private mMA As IMovingAverageStudy

Private mDevs As Double

Private mNumberOfValuesToCache As Long

'Private mBottomCache As TradeBuildSP.IStudyValueCache
'Private mCentreCache As TradeBuildSP.IStudyValueCache
'Private mTopCache As TradeBuildSP.IStudyValueCache
'Private mSpreadCache As TradeBuildSP.IStudyValueCache

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
' Custom methods that must be impleemented
'================================================================================

Private Sub customInitialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)
Set mCommonServiceConsumer = commonServiceConsumer
Set mStudyServiceConsumer = studyServiceConsumer
mStudyFoundation.ReplayUsingStudyValues = True
mNumberOfValuesToCache = numberOfValuesToCache
customProcessParameters parameters
End Sub

Private Sub customNotify( _
                ev As TradeBuildSP.StudyValueEvent)
Static sPrevVal As Variant

If Not mStudyFoundation.isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
    sPrevVal = ev.value
End If

mSDStudy.Notify ev
mMAStudy.Notify ev

If IsEmpty(mMA.ma.value) Or IsEmpty(mSD.SD.value) Then Exit Sub

notifyValue mMA.ma.value, BBValueCentre, ev.barnumber, ev.timestamp
notifyValue mMA.ma.value + mDevs * mSD.SD.value, BBValueTop, ev.barnumber, ev.timestamp
notifyValue mMA.ma.value - mDevs * mSD.SD.value, BBValueBottom, ev.barnumber, ev.timestamp
notifyValue 2 * mDevs * mSD.SD.value, BBValueSpread, ev.barnumber, ev.timestamp
End Sub

Private Function customNumberOfBarsRequired() As Long
Dim lStudy As IStudy
Set lStudy = mSD
customNumberOfBarsRequired = lStudy.numberOfBarsRequired
End Function

Private Sub customProcessParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter
Dim lparams As IParameters
Dim maType As String
Dim lStudy As IStudy
Dim valueNames(0) As String

valueNames(0) = "in"

Set param = params.getFirstParameter

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
    Set param = params.getNextParameter
Loop

Set mMA = gCreateMA(maType, _
                    mCommonServiceConsumer, _
                    mStudyServiceConsumer, _
                    mPeriods, _
                    1)
Set mMAStudy = mMA

Set lparams = GStandardDeviation.defaultParameters
lparams.setParameterValue ParamPeriods, mPeriods
Set lStudy = mSD
lStudy.initialise mCommonServiceConsumer, _
                    mStudyServiceConsumer, _
                    mCommonServiceConsumer.GenerateGUIDString, _
                    lparams, _
                    mNumberOfValuesToCache, _
                    valueNames, _
                    Nothing
End Sub

Private Function customStudyDefinition() As IStudyDefinition
Set customStudyDefinition = GBollingerBands.studyDefinition
End Function

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

'================================================================================
' Helper Functions
'================================================================================

Private Function getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As StudyValue
Select Case UCase$(valueName)
Case UCase$(BBValueBottom)
    getStudyValue = Me.bottom(ref)
Case UCase$(BBValueCentre)
    getStudyValue = Me.centre(ref)
Case UCase$(BBValueSpread)
    getStudyValue = Me.spread(ref)
Case UCase$(BBValueTop)
    getStudyValue = Me.top(ref)
End Select
End Function

Private Sub notifyValue( _
                ByVal value As Variant, _
                ByVal valueName As String, _
                ByVal barnumber As Long, _
                ByVal timestamp As Date)
Dim ev As StudyValueEvent
If Not IsEmpty(value) Then
    ev.value = value
    ev.valueName = valueName
    Set ev.Source = Me
    ev.barnumber = barnumber
    ev.timestamp = timestamp
    mStudyFoundation.notifyValue ev
End If
End Sub


'================================================================================
' From this point on the code is identical for all studies - do not change!!!!
'================================================================================


'================================================================================
' IStudy Interface Members
'================================================================================

Private Function IStudy_addStudy( _
                ByVal study As IStudy, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion

Set IStudy_addStudy = mStudyFoundation.addStudy( _
                            study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)
End Function

Private Function IStudy_addStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion
Set IStudy_addStudyValueListener = mStudyFoundation.addStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)
End Function

Private Property Get IStudy_baseStudy() As IStudy
Set IStudy_baseStudy = Me
End Property

Private Function IStudy_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As StudyValue
'Study_getStudyValue = mStudyFoundation.getStudyValue(ValueName, ref)
IStudy_getStudyValue = getStudyValue(valueName, ref)
End Function

Private Property Get IStudy_id() As String
IStudy_id = mStudyFoundation.id
End Property

Private Sub IStudy_initialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)
Set mStudyFoundation = New StudyFoundation
mStudyFoundation.initialise Me, _
                            id, _
                            parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudy, _
                            commonServiceConsumer, _
                            studyServiceConsumer
                            

customInitialise commonServiceConsumer, _
                studyServiceConsumer, _
                id, _
                parameters, _
                numberOfValuesToCache, _
                inputValueNames, _
                underlyingStudy
End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = mStudyFoundation.instanceName
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = mStudyFoundation.instancePath
End Property

Private Sub IStudy_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
mStudyFoundation.NotifyInput ev
customNotify ev
End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long
IStudy_numberOfBarsRequired = mStudyFoundation.numberOfBarsRequired(customNumberOfBarsRequired)
End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal valueName As String = "") As Long
IStudy_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)
End Function

Private Property Get IStudy_parameters() As IParameters
Set IStudy_parameters = mStudyFoundation.parameters
End Property

Private Sub IStudy_removeStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
mStudyFoundation.removeStudyValueListener listener
End Sub

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = customStudyDefinition
End Property



