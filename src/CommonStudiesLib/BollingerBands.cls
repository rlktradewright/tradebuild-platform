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

Implements IStudy

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

Private mStudyManager As StudyManager


Private mPeriods As Long
Private mCentreBandWidth As Double
Private mEdgeBandWidth As Double
Private mSlopeThreshold As Double

Private mSD As StandardDeviation
Private mSDStudy As IStudy

Private mMAStudy As IStudy

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
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Class_Terminate()
Const ProcName As String = "Class_Terminate"
On Error GoTo Err

Debug.Print "BollingerBands study terminated"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Custom methods that must be impleemented
'@================================================================================

Private Sub customInitialise( _
                ByVal Id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As IStudy, _
                ByVal pSessionFuture As IFuture)
Const ProcName As String = "customInitialise"
On Error GoTo Err

mStudyFoundation.ReplayUsingStudyValues = True
mNumberOfValuesToCache = numberOfValuesToCache
customProcessParameters Parameters

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub customNotify( _
                ev As StudyValueEventData)
Const ProcName As String = "customNotify"
On Error GoTo Err

Static sPrevVal As Variant

If Not mStudyFoundation.IsNewBar Then
    If ev.sVal.Value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.Value

mSDStudy.Notify ev
mMAStudy.Notify ev

Dim ma As Variant
ma = mMAStudy.getStudyValue(MovingAverageStudyValueName, 0).Value
If IsEmpty(ma) Or IsEmpty(mSD.SD.Value) Then Exit Sub

notifyValue ma, BBValueCentre, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
notifyValue ma + mDevs * mSD.SD.Value, BBValueTop, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
notifyValue ma - mDevs * mSD.SD.Value, BBValueBottom, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
notifyValue 2 * mDevs * mSD.SD.Value, BBValueSpread, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function customNumberOfBarsRequired() As Long
Dim lStudy As IStudy
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

Set lStudy = mSD
customNumberOfBarsRequired = lStudy.NumberOfBarsRequired

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Const ProcName As String = "customProcessParameters"
On Error GoTo Err

Dim maType As String

Dim param As Parameter
For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(BBParamCentreBandWidth)
        mCentreBandWidth = param.Value
    Case UCase$(BBParamDeviations)
        mDevs = param.Value
    Case UCase$(BBParamEdgeBandWidth)
        mEdgeBandWidth = param.Value
    Case UCase$(BBParamPeriods)
        mPeriods = param.Value
    Case UCase$(BBParamSlopeThreshold)
        mSlopeThreshold = param.Value
    Case UCase$(BBParamMAType)
        maType = param.Value
    End Select
Next

Set mMAStudy = gCreateMA(mStudyManager, maType, mPeriods, 0)

Dim lparams As Parameters
Set lparams = GStandardDeviation.defaultParameters
lparams.SetParameterValue ParamPeriods, mPeriods

Dim valueNames(0) As String
valueNames(0) = "in"

Dim lStudy As IStudy
Set lStudy = mSD
lStudy.Initialise mStudyManager, _
                    GenerateGUIDString, _
                    lparams, _
                    mNumberOfValuesToCache, _
                    valueNames, _
                    Nothing, _
                    Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GBollingerBands.StudyDefinition

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Bottom(Optional ByVal ref As Long) As SValue
Dim sv As SValue
Const ProcName As String = "Bottom"
On Error GoTo Err

Bottom = mMAStudy.getStudyValue(MovingAverageStudyValueName, ref)
sv = mSD.SD(ref)
If IsEmpty(Bottom.Value) Or IsEmpty(sv.Value) Then Exit Property
Bottom.Value = Bottom.Value - mDevs * sv.Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BottomFalling(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "BottomFalling"
On Error GoTo Err

BottomFalling = BottomSlope(ref, periods)
If IsEmpty(BottomFalling.Value) Then Exit Property
BottomFalling.Value = (BottomFalling.Value <= (-1 * mSlopeThreshold))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BottomFlat(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "BottomFlat"
On Error GoTo Err

BottomFlat = BottomSlope(ref, periods)
If IsEmpty(BottomFlat.Value) Then Exit Property
BottomFlat.Value = (BottomFlat.Value < mSlopeThreshold And _
                    BottomFlat.Value > (-1 * mSlopeThreshold))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BottomRising(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "BottomRising"
On Error GoTo Err

BottomRising = BottomSlope(ref, periods)
If IsEmpty(BottomRising.Value) Then Exit Property
BottomRising.Value = (BottomRising.Value >= mSlopeThreshold)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BottomSlope(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As SValue
Dim sv As SValue
Const ProcName As String = "BottomSlope"
On Error GoTo Err

BottomSlope = Bottom(ref)
sv = Bottom(ref - periods + 1)
If IsEmpty(BottomSlope.Value) Or IsEmpty(sv.Value) Then Exit Property
BottomSlope.Value = (BottomSlope.Value - sv.Value) / periods

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Centre(Optional ByVal ref As Long) As SValue
Const ProcName As String = "Centre"
On Error GoTo Err

Centre = mMAStudy.getStudyValue(MovingAverageStudyValueName, ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CentreFalling(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "CentreFalling"
On Error GoTo Err

CentreFalling = CentreSlope(ref, periods)
If IsEmpty(CentreFalling.Value) Then Exit Property
CentreFalling.Value = (CentreFalling.Value <= (-1 * mSlopeThreshold))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CentreFlat(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "CentreFlat"
On Error GoTo Err

CentreFlat = CentreSlope(ref, periods)
If IsEmpty(CentreFlat.Value) Then Exit Property
CentreFlat.Value = (CentreFlat.Value < mSlopeThreshold And _
    CentreFlat.Value > (-1 * mSlopeThreshold))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CentreRising(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "CentreRising"
On Error GoTo Err

CentreRising = CentreSlope(ref, periods)
If IsEmpty(CentreRising.Value) Then Exit Property
CentreRising.Value = (CentreRising.Value >= mSlopeThreshold)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CentreSlope(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As SValue
Dim sv As SValue
Const ProcName As String = "CentreSlope"
On Error GoTo Err

CentreSlope = Centre(ref)
sv = Centre(ref - periods + 1)
If IsEmpty(CentreSlope.Value) Or IsEmpty(sv.Value) Then Exit Property
CentreSlope.Value = (CentreSlope.Value - sv.Value) / periods

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get InBottomBand(ByVal Value As Double, _
                                Optional ByVal ref As Long) As SValue
Const ProcName As String = "inBottomBand"
On Error GoTo Err

InBottomBand = Bottom(ref)
If IsEmpty(InBottomBand.Value) Then Exit Property
InBottomBand.Value = (Value <= (InBottomBand.Value + mEdgeBandWidth))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get InCentreBand(ByVal Value As Double, _
                                Optional ByVal ref As Long) As SValue
Dim bandUpper As Double
Dim bandLower As Double

Const ProcName As String = "inCentreBand"
On Error GoTo Err

InCentreBand = Centre(ref)
If IsEmpty(InCentreBand.Value) Then Exit Property
bandUpper = InCentreBand.Value + (mCentreBandWidth / 2)
bandLower = bandUpper - mCentreBandWidth

InCentreBand.Value = (Value <= bandUpper _
                    And _
                    Value >= bandLower)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get InTopBand(ByVal Value As Double, _
                            Optional ByVal ref As Long) As SValue
Const ProcName As String = "inTopBand"
On Error GoTo Err

InTopBand = Top(ref)
If IsEmpty(InTopBand.Value) Then Exit Property
InTopBand.Value = (Value >= (InTopBand.Value - mEdgeBandWidth))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RangeIntersectsCentreBand( _
                    ByVal lowValue As Double, _
                    ByVal highValue As Double, _
                    Optional ByVal ref As Long) As SValue
Dim bandUpper As Double
Dim bandLower As Double

Const ProcName As String = "RangeIntersectsCentreBand"
On Error GoTo Err

RangeIntersectsCentreBand = Centre(ref)
If IsEmpty(RangeIntersectsCentreBand.Value) Then Exit Property

bandUpper = RangeIntersectsCentreBand.Value + (mCentreBandWidth / 2)
bandLower = bandUpper - mCentreBandWidth

If highValue >= bandLower And highValue <= bandUpper Then
    RangeIntersectsCentreBand.Value = True
ElseIf lowValue >= bandLower And lowValue <= bandUpper Then
    RangeIntersectsCentreBand.Value = True
ElseIf lowValue < bandLower And highValue > bandUpper Then
    RangeIntersectsCentreBand.Value = True
Else
    RangeIntersectsCentreBand.Value = False
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Spread(Optional ByVal ref As Long) As SValue
Dim sv As SValue
Const ProcName As String = "Spread"
On Error GoTo Err

Spread = Top(ref)
sv = Bottom(ref)
If IsEmpty(Spread.Value) Or IsEmpty(sv.Value) Then Exit Property
Spread.Value = Spread.Value - sv.Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Top(Optional ByVal ref As Long) As SValue
Dim sv As SValue
Const ProcName As String = "top"
On Error GoTo Err

Top = mMAStudy.getStudyValue(MovingAverageStudyValueName, ref)
sv = mSD.SD(ref)
If IsEmpty(Top.Value) Or IsEmpty(sv.Value) Then Exit Property
Top.Value = Top.Value + mDevs * sv.Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TopFalling(Optional ByVal ref As Long, _
                                Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "topFalling"
On Error GoTo Err

TopFalling = TopSlope(ref, periods)
If IsEmpty(TopFalling.Value) Then Exit Property
TopFalling.Value = (TopFalling.Value <= (-1 * mSlopeThreshold))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TopFlat(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "topFlat"
On Error GoTo Err

TopFlat = TopSlope(ref, periods)
If IsEmpty(TopFlat.Value) Then Exit Property
TopFlat.Value = (TopFlat.Value < mSlopeThreshold And _
    TopFlat.Value > (-1 * mSlopeThreshold))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TopRising(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As SValue
Const ProcName As String = "topRising"
On Error GoTo Err

TopRising = TopSlope(ref, periods)
If IsEmpty(TopRising.Value) Then Exit Property
TopRising.Value = (TopRising.Value >= mSlopeThreshold)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TopSlope(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 1) As SValue
Dim sv As SValue
Const ProcName As String = "topSlope"
On Error GoTo Err

TopSlope = Top(ref)
sv = Top(ref - periods + 1)
If IsEmpty(TopSlope.Value) Or IsEmpty(sv.Value) Then Exit Property
TopSlope.Value = (TopSlope.Value - sv.Value) / periods

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
    getStudyValue = Me.Bottom(ref)
Case UCase$(BBValueCentre)
    getStudyValue = Me.Centre(ref)
Case UCase$(BBValueSpread)
    getStudyValue = Me.Spread(ref)
Case UCase$(BBValueTop)
    getStudyValue = Me.Top(ref)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub notifyValue( _
                ByVal Value As Variant, _
                ByVal valueName As String, _
                ByVal BarNumber As Long, _
                ByVal BarStartTime As Date, _
                ByVal Timestamp As Date)
Dim ev As StudyValueEventData
Const ProcName As String = "notifyValue"
On Error GoTo Err

If Not IsEmpty(Value) Then
    ev.sVal.Value = Value
    ev.valueName = valueName
    Set ev.Source = Me
    ev.sVal.BarNumber = BarNumber
    ev.sVal.BarStartTime = BarStartTime
    ev.sVal.Timestamp = Timestamp
    mStudyFoundation.notifyValue ev
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


Private Function numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "numberOfCachedValues"
On Error GoTo Err

numberOfCachedValues = mMAStudy.numberOfCachedValues("")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' From this point on the code is identical for all studies - do not change!!!!
'@================================================================================


'@================================================================================
' IStudy Interface Members
'@================================================================================

Private Function IStudy_AddStudy( _
                ByVal Study As IStudy, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController

Const ProcName As String = "IStudy_AddStudy"
On Error GoTo Err

Set IStudy_AddStudy = mStudyFoundation.AddStudy( _
                            Study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudy_AddStudyValueListener( _
                ByVal listener As IStudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "IStudy_AddStudyValueListener"
On Error GoTo Err

Set IStudy_AddStudyValueListener = mStudyFoundation.AddStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_BaseStudy() As IStudy
Const ProcName As String = "IStudy_BaseStudy"
On Error GoTo Err

Set IStudy_BaseStudy = Me

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IStudy_GetStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
'IStudy_GetStudyValue = mStudyFoundation.getStudyValue(ValueName, ref)
Const ProcName As String = "IStudy_GetStudyValue"
On Error GoTo Err

IStudy_GetStudyValue = getStudyValue(valueName, ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudy_GetValueTicksize( _
                ByVal valueName As String) As Double

Const ProcName As String = "IStudy_GetValueTicksize"
On Error GoTo Err



Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_id() As String
Const ProcName As String = "IStudy_id"
On Error GoTo Err

IStudy_id = mStudyFoundation.Id

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStudy_Initialise( _
                ByVal StudyManager As StudyManager, _
                ByVal Id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As IStudy, _
                ByVal pSessionFuture As IFuture)
Const ProcName As String = "IStudy_Initialise"
On Error GoTo Err

Set mStudyManager = StudyManager
Set mStudyFoundation = New StudyFoundation
mStudyFoundation.Initialise StudyManager, _
                            Me, _
                            Id, _
                            Parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudy
                            

customInitialise Id, _
                Parameters, _
                numberOfValuesToCache, _
                inputValueNames, _
                underlyingStudy, _
                pSessionFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_InstanceName() As String
Const ProcName As String = "IStudy_InstanceName"
On Error GoTo Err

IStudy_InstanceName = mStudyFoundation.InstanceName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudy_InstancePath() As String
Const ProcName As String = "IStudy_InstancePath"
On Error GoTo Err

IStudy_InstancePath = mStudyFoundation.InstancePath

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStudy_Notify( _
                ev As StudyValueEventData)
Const ProcName As String = "IStudy_Notify"
On Error GoTo Err

mStudyFoundation.NotifyInput ev
customNotify ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_NumberOfBarsRequired() As Long
Const ProcName As String = "IStudy_NumberOfBarsRequired"
On Error GoTo Err

IStudy_NumberOfBarsRequired = mStudyFoundation.NumberOfBarsRequired(customNumberOfBarsRequired)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IStudy_NumberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "IStudy_NumberOfCachedValues"
On Error GoTo Err

IStudy_NumberOfCachedValues = numberOfCachedValues(valueName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_Parameters() As Parameters
Const ProcName As String = "IStudy_Parameters"
On Error GoTo Err

Set IStudy_Parameters = mStudyFoundation.Parameters

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStudy_RemoveStudyValueListener( _
                ByVal listener As IStudyValueListener)
Const ProcName As String = "IStudy_RemoveStudyValueListener"
On Error GoTo Err

mStudyFoundation.RemoveStudyValueListener listener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStudy_SetInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)

Const ProcName As String = "IStudy_SetInputTicksize"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_StudyDefinition() As StudyDefinition
Const ProcName As String = "IStudy_StudyDefinition"
On Error GoTo Err

Set IStudy_StudyDefinition = customStudyDefinition

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property



