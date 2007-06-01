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

Implements Study

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




Private mPeriods As Long
Private mCentreBandWidth As Double
Private mEdgeBandWidth As Double
Private mSlopeThreshold As Double

Private mSD As StandardDeviation
Private mSDStudy As Study

Private mMAStudy As Study

Private mDevs As Double

Private mNumberOfValuesToCache As Long

'Private mBottomCache As StudyValueCache
'Private mCentreCache As StudyValueCache
'Private mTopCache As StudyValueCache
'Private mSpreadCache As StudyValueCache

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
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As session)

mStudyFoundation.ReplayUsingStudyValues = True
mNumberOfValuesToCache = numberOfValuesToCache
customProcessParameters Parameters
End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Static sPrevVal As Variant
Dim ma As Variant

If Not mStudyFoundation.isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.value

mSDStudy.notify ev
mMAStudy.notify ev

ma = mMAStudy.getStudyValue(MovingAverageStudyValueName, 0).value
If IsEmpty(ma) Or IsEmpty(mSD.SD.value) Then Exit Sub

notifyValue ma, BBValueCentre, ev.barNumber, ev.timestamp, ev.barStartTime
notifyValue ma + mDevs * mSD.SD.value, BBValueTop, ev.barNumber, ev.timestamp, ev.barStartTime
notifyValue ma - mDevs * mSD.SD.value, BBValueBottom, ev.barNumber, ev.timestamp, ev.barStartTime
notifyValue 2 * mDevs * mSD.SD.value, BBValueSpread, ev.barNumber, ev.timestamp, ev.barStartTime
End Sub

Private Function customNumberOfBarsRequired() As Long
Dim lStudy As Study
Set lStudy = mSD
customNumberOfBarsRequired = lStudy.numberOfBarsRequired
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter
Dim lparams As Parameters
Dim maType As String
Dim lStudy As Study
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

Set mMAStudy = gCreateMA(maType, _
                    mPeriods, _
                    0)

Set lparams = GStandardDeviation.defaultParameters
lparams.setParameterValue ParamPeriods, mPeriods
Set lStudy = mSD
lStudy.initialise GenerateGUIDString, _
                    lparams, _
                    mNumberOfValuesToCache, _
                    valueNames, _
                    Nothing, _
                    Nothing
End Sub

Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GBollingerBands.StudyDefinition
End Function

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get bottom(Optional ByVal ref As Long) As StudyValue
Dim sv As StudyValue
bottom = mMAStudy.getStudyValue(MovingAverageStudyValueName, ref)
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
centre = mMAStudy.getStudyValue(MovingAverageStudyValueName, ref)
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
top = mMAStudy.getStudyValue(MovingAverageStudyValueName, ref)
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
                ByVal barNumber As Long, _
                ByVal barStartTime As Date, _
                ByVal timestamp As Date)
Dim ev As StudyValueEvent
If Not IsEmpty(value) Then
    ev.value = value
    ev.valueName = valueName
    Set ev.Source = Me
    ev.barNumber = barNumber
    ev.barStartTime = barStartTime
    ev.timestamp = timestamp
    mStudyFoundation.notifyValue ev
End If
End Sub


Private Function numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
numberOfCachedValues = mMAStudy.numberOfCachedValues("")
End Function

'================================================================================
' From this point on the code is identical for all studies - do not change!!!!
'================================================================================


'================================================================================
' Study Interface Members
'================================================================================

Private Function Study_addStudy( _
                ByVal Study As Study, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskCompletion

Set Study_addStudy = mStudyFoundation.addStudy( _
                            Study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)
End Function

Private Function Study_addStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskCompletion
Set Study_addStudyValueListener = mStudyFoundation.addStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)
End Function

Private Property Get Study_baseStudy() As Study
Set Study_baseStudy = Me
End Property

Private Function Study_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As StudyValue
'Study_getStudyValue = mStudyFoundation.getStudyValue(ValueName, ref)
Study_getStudyValue = getStudyValue(valueName, ref)
End Function

Private Function Study_getValueTicksize( _
                ByVal valueName As String) As Double

End Function

Private Property Get Study_id() As String
Study_id = mStudyFoundation.id
End Property

Private Sub Study_initialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As session)
Set mStudyFoundation = New StudyFoundation
mStudyFoundation.initialise Me, _
                            id, _
                            Parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudy
                            

customInitialise id, _
                Parameters, _
                numberOfValuesToCache, _
                inputValueNames, _
                underlyingStudy, _
                pSession
End Sub

Private Property Get Study_instanceName() As String
Study_instanceName = mStudyFoundation.instanceName
End Property

Private Property Get Study_instancePath() As String
Study_instancePath = mStudyFoundation.instancePath
End Property

Private Sub Study_Notify( _
                ev As StudyValueEvent)
mStudyFoundation.notifyInput ev
customNotify ev
End Sub

Private Property Get Study_numberOfBarsRequired() As Long
Study_numberOfBarsRequired = mStudyFoundation.numberOfBarsRequired(customNumberOfBarsRequired)
End Property

Private Function Study_numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Study_numberOfCachedValues = numberOfCachedValues(valueName)
End Function

Private Property Get Study_parameters() As Parameters
Set Study_parameters = mStudyFoundation.Parameters
End Property

Private Sub Study_removeStudyValueListener( _
                ByVal listener As StudyValueListener)
mStudyFoundation.removeStudyValueListener listener
End Sub

Private Sub Study_setInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)

End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Set Study_studyDefinition = customStudyDefinition
End Property



