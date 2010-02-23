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

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements Study

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "BollingerBands"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

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

'Private mBottomCache As ValueCache
'Private mCentreCache As ValueCache
'Private mTopCache As ValueCache
'Private mSpreadCache As ValueCache

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mSD = New StandardDeviation
Set mSDStudy = mSD

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub Class_Terminate()
Const ProcName As String = "Class_Terminate"
On Error GoTo Err

Debug.Print "BollingerBands study terminated"

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Custom methods that must be impleemented
'@================================================================================

Private Sub customInitialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)

Const ProcName As String = "customInitialise"
On Error GoTo Err

mStudyFoundation.ReplayUsingStudyValues = True
mNumberOfValuesToCache = numberOfValuesToCache
customProcessParameters Parameters

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Static sPrevVal As Variant
Dim ma As Variant

Const ProcName As String = "customNotify"
On Error GoTo Err

If Not mStudyFoundation.IsNewBar Then
    If ev.sVal.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.value

mSDStudy.Notify ev
mMAStudy.Notify ev

ma = mMAStudy.getStudyValue(MovingAverageStudyValueName, 0).value
If IsEmpty(ma) Or IsEmpty(mSD.SD.value) Then Exit Sub

notifyValue ma, BBValueCentre, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
notifyValue ma + mDevs * mSD.SD.value, BBValueTop, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
notifyValue ma - mDevs * mSD.SD.value, BBValueBottom, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
notifyValue 2 * mDevs * mSD.SD.value, BBValueSpread, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function customNumberOfBarsRequired() As Long
Dim lStudy As Study
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

Set lStudy = mSD
customNumberOfBarsRequired = lStudy.NumberOfBarsRequired

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter
Dim lparams As Parameters
Dim maType As String
Dim lStudy As Study
Dim valueNames(0) As String

Const ProcName As String = "customProcessParameters"
On Error GoTo Err

valueNames(0) = "in"

For Each param In params
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
Next

Set mMAStudy = gCreateMA(maType, _
                    mPeriods, _
                    0)

Set lparams = GStandardDeviation.defaultParameters
lparams.SetParameterValue ParamPeriods, mPeriods
Set lStudy = mSD
lStudy.initialise GenerateGUIDString, _
                    lparams, _
                    mNumberOfValuesToCache, _
                    valueNames, _
                    Nothing, _
                    Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GBollingerBands.StudyDefinition

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get bottom(Optional ByVal ref As Long) As SValue
Dim sv As SValue
Const ProcName As String = "bottom"
On Error GoTo Err

bottom = mMAStudy.getStudyValue(MovingAverageStudyValueName, ref)
sv = mSD.SD(ref)
If IsEmpty(bottom.value) Or IsEmpty(sv.value) Then Exit Property
bottom.value = bottom.value - mDevs * sv.value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get bottomFalling(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "bottomFalling"
On Error GoTo Err

bottomFalling = bottomSlope(ref, periods)
If IsEmpty(bottomFalling.value) Then Exit Property
bottomFalling.value = (bottomFalling.value <= (-1 * mSlopeThreshold))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get bottomFlat(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "bottomFlat"
On Error GoTo Err

bottomFlat = bottomSlope(ref, periods)
If IsEmpty(bottomFlat.value) Then Exit Property
bottomFlat.value = (bottomFlat.value < mSlopeThreshold And _
                    bottomFlat.value > (-1 * mSlopeThreshold))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get bottomRising(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "bottomRising"
On Error GoTo Err

bottomRising = bottomSlope(ref, periods)
If IsEmpty(bottomRising.value) Then Exit Property
bottomRising.value = (bottomRising.value >= mSlopeThreshold)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get bottomSlope(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As SValue
Dim sv As SValue
Const ProcName As String = "bottomSlope"
On Error GoTo Err

bottomSlope = bottom(ref)
sv = bottom(ref - periods + 1)
If IsEmpty(bottomSlope.value) Or IsEmpty(sv.value) Then Exit Property
bottomSlope.value = (bottomSlope.value - sv.value) / periods

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get centre(Optional ByVal ref As Long) As SValue
Const ProcName As String = "centre"
On Error GoTo Err

centre = mMAStudy.getStudyValue(MovingAverageStudyValueName, ref)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get centreFalling(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "centreFalling"
On Error GoTo Err

centreFalling = centreSlope(ref, periods)
If IsEmpty(centreFalling.value) Then Exit Property
centreFalling.value = (centreFalling.value <= (-1 * mSlopeThreshold))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get centreFlat(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "centreFlat"
On Error GoTo Err

centreFlat = centreSlope(ref, periods)
If IsEmpty(centreFlat.value) Then Exit Property
centreFlat.value = (centreFlat.value < mSlopeThreshold And _
    centreFlat.value > (-1 * mSlopeThreshold))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get centreRising(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "centreRising"
On Error GoTo Err

centreRising = centreSlope(ref, periods)
If IsEmpty(centreRising.value) Then Exit Property
centreRising.value = (centreRising.value >= mSlopeThreshold)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get centreSlope(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As SValue
Dim sv As SValue
Const ProcName As String = "centreSlope"
On Error GoTo Err

centreSlope = centre(ref)
sv = centre(ref - periods + 1)
If IsEmpty(centreSlope.value) Or IsEmpty(sv.value) Then Exit Property
centreSlope.value = (centreSlope.value - sv.value) / periods

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get inBottomBand(ByVal value As Double, _
                                Optional ByVal ref As Long) As SValue
Const ProcName As String = "inBottomBand"
On Error GoTo Err

inBottomBand = bottom(ref)
If IsEmpty(inBottomBand.value) Then Exit Property
inBottomBand.value = (value <= (inBottomBand.value + mEdgeBandWidth))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get inCentreBand(ByVal value As Double, _
                                Optional ByVal ref As Long) As SValue
Dim bandUpper As Double
Dim bandLower As Double

Const ProcName As String = "inCentreBand"
On Error GoTo Err

inCentreBand = centre(ref)
If IsEmpty(inCentreBand.value) Then Exit Property
bandUpper = inCentreBand.value + (mCentreBandWidth / 2)
bandLower = bandUpper - mCentreBandWidth

inCentreBand.value = (value <= bandUpper _
                    And _
                    value >= bandLower)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get inTopBand(ByVal value As Double, _
                            Optional ByVal ref As Long) As SValue
Const ProcName As String = "inTopBand"
On Error GoTo Err

inTopBand = top(ref)
If IsEmpty(inTopBand.value) Then Exit Property
inTopBand.value = (value >= (inTopBand.value - mEdgeBandWidth))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get rangeIntersectsCentreBand( _
                    ByVal lowValue As Double, _
                    ByVal highValue As Double, _
                    Optional ByVal ref As Long) As SValue
Dim bandUpper As Double
Dim bandLower As Double

Const ProcName As String = "rangeIntersectsCentreBand"
On Error GoTo Err

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

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get spread(Optional ByVal ref As Long) As SValue
Dim sv As SValue
Const ProcName As String = "spread"
On Error GoTo Err

spread = top(ref)
sv = bottom(ref)
If IsEmpty(spread.value) Or IsEmpty(sv.value) Then Exit Property
spread.value = spread.value - sv.value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get top(Optional ByVal ref As Long) As SValue
Dim sv As SValue
Const ProcName As String = "top"
On Error GoTo Err

top = mMAStudy.getStudyValue(MovingAverageStudyValueName, ref)
sv = mSD.SD(ref)
If IsEmpty(top.value) Or IsEmpty(sv.value) Then Exit Property
top.value = top.value + mDevs * sv.value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get topFalling(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "topFalling"
On Error GoTo Err

topFalling = topSlope(ref, periods)
If IsEmpty(topFalling.value) Then Exit Property
topFalling.value = (topFalling.value <= (-1 * mSlopeThreshold))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get topFlat(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "topFlat"
On Error GoTo Err

topFlat = topSlope(ref, periods)
If IsEmpty(topFlat.value) Then Exit Property
topFlat.value = (topFlat.value < mSlopeThreshold And _
    topFlat.value > (-1 * mSlopeThreshold))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get topRising(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "topRising"
On Error GoTo Err

topRising = topSlope(ref, periods)
If IsEmpty(topRising.value) Then Exit Property
topRising.value = (topRising.value >= mSlopeThreshold)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get topSlope(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As SValue
Dim sv As SValue
Const ProcName As String = "topSlope"
On Error GoTo Err

topSlope = top(ref)
sv = top(ref - periods + 1)
If IsEmpty(topSlope.value) Or IsEmpty(sv.value) Then Exit Property
topSlope.value = (topSlope.value - sv.value) / periods

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Const ProcName As String = "getStudyValue"
On Error GoTo Err

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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub notifyValue( _
                ByVal value As Variant, _
                ByVal valueName As String, _
                ByVal barNumber As Long, _
                ByVal barStartTime As Date, _
                ByVal timestamp As Date)
Dim ev As StudyValueEvent
Const ProcName As String = "notifyValue"
On Error GoTo Err

If Not IsEmpty(value) Then
    ev.sVal.value = value
    ev.valueName = valueName
    Set ev.Source = Me
    ev.sVal.barNumber = barNumber
    ev.sVal.barStartTime = barStartTime
    ev.sVal.timestamp = timestamp
    mStudyFoundation.notifyValue ev
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub


Private Function numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "numberOfCachedValues"
On Error GoTo Err

numberOfCachedValues = mMAStudy.numberOfCachedValues("")

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' From this point on the code is identical for all studies - do not change!!!!
'@================================================================================


'@================================================================================
' Study Interface Members
'@================================================================================

Private Function Study_addStudy( _
                ByVal Study As Study, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController

Const ProcName As String = "Study_addStudy"
On Error GoTo Err

Set Study_addStudy = mStudyFoundation.AddStudy( _
                            Study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function Study_addStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "Study_addStudyValueListener"
On Error GoTo Err

Set Study_addStudyValueListener = mStudyFoundation.AddStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Get Study_baseStudy() As Study
Const ProcName As String = "Study_baseStudy"
On Error GoTo Err

Set Study_baseStudy = Me

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Function Study_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
'Study_getStudyValue = mStudyFoundation.getStudyValue(ValueName, ref)
Const ProcName As String = "Study_getStudyValue"
On Error GoTo Err

Study_getStudyValue = getStudyValue(valueName, ref)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function Study_getValueTicksize( _
                ByVal valueName As String) As Double

Const ProcName As String = "Study_getValueTicksize"
On Error GoTo Err



Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Get Study_id() As String
Const ProcName As String = "Study_id"
On Error GoTo Err

Study_id = mStudyFoundation.id

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub Study_initialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)
Const ProcName As String = "Study_initialise"
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get Study_instanceName() As String
Const ProcName As String = "Study_instanceName"
On Error GoTo Err

Study_instanceName = mStudyFoundation.InstanceName

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get Study_instancePath() As String
Const ProcName As String = "Study_instancePath"
On Error GoTo Err

Study_instancePath = mStudyFoundation.InstancePath

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub Study_Notify( _
                ev As StudyValueEvent)
Const ProcName As String = "Study_Notify"
On Error GoTo Err

mStudyFoundation.NotifyInput ev
customNotify ev

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get Study_numberOfBarsRequired() As Long
Const ProcName As String = "Study_numberOfBarsRequired"
On Error GoTo Err

Study_numberOfBarsRequired = mStudyFoundation.NumberOfBarsRequired(customNumberOfBarsRequired)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Function Study_numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "Study_numberOfCachedValues"
On Error GoTo Err

Study_numberOfCachedValues = numberOfCachedValues(valueName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Get Study_parameters() As Parameters
Const ProcName As String = "Study_parameters"
On Error GoTo Err

Set Study_parameters = mStudyFoundation.Parameters

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub Study_removeStudyValueListener( _
                ByVal listener As StudyValueListener)
Const ProcName As String = "Study_removeStudyValueListener"
On Error GoTo Err

mStudyFoundation.RemoveStudyValueListener listener

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub Study_setInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)

Const ProcName As String = "Study_setInputTicksize"
On Error GoTo Err



Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Const ProcName As String = "Study_studyDefinition"
On Error GoTo Err

Set Study_studyDefinition = customStudyDefinition

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property



