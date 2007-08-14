VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstMomentumBars"
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

Implements Study
Implements BarStudy
Implements CollectionChangeListener

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

Private WithEvents mSession As session
Attribute mSession.VB_VarHelpID = -1

Private mMovementPerBar As Long

Private mBars As Bars
Private mBar As bar

Private mNumberOfValuesToCache As Long

Private mBarNumber As Long

' accumulated volume at end of previous bar
Private mPrevBarVolume As Long

' the latest volume figure received
Private mLatestVolume As Long

Private mLatestTimestamp As Date

Private mVolumeCache As StudyValueCache
Private mTotalVolumeCache As StudyValueCache

Private mPriceTickSize As Double
Private mVolumeTickSize As Double

Private mNewBarsPending As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Debug.Print "Constant Time Bars study terminated"
End Sub

'================================================================================
' BarStudy Interface Members
'================================================================================

Private Property Get BarStudy_bars() As Bars
Set BarStudy_bars = mBars
End Property

Private Property Get BarStudy_closeValueName() As String
BarStudy_closeValueName = BarValueClose
End Property

Private Property Get BarStudy_highValueName() As String
BarStudy_highValueName = BarValueHigh
End Property

Private Property Get BarStudy_lowValueName() As String
BarStudy_lowValueName = BarValueLow
End Property

Private Property Get BarStudy_openValueName() As String
BarStudy_openValueName = BarValueOpen
End Property

Private Sub BarStudy_switchUnderlyingStudy( _
                ByVal newUnderlyingStudy As InputStudy)
                
If newUnderlyingStudy.session.currentSessionStartTime > mSession.currentSessionStartTime Then
    ' we are now in a different session so reset volume variables
    mPrevBarVolume = 0
End If

Set mSession = newUnderlyingStudy.session
mBars.changeSession mSession

' because we are switching to a new source, the volume figures may not be consistent
' with those from the previous source - for example, if the previous source was historical
' data, it may not have run from the start of the session so the aggregate volume will
' be lower than the tick data source.
'
' To compensate for this, we have to adjust mPrevBarVolume when we receive the first
' volume tick from the new source, so set it to indicate this
mPrevBarVolume = -1
End Sub

'================================================================================
' CollectionChangeListener Interface Members
'================================================================================

Private Sub CollectionChangeListener_Change(ev As CollectionChangeEvent)
If ev.changeType = CollItemAdded Then
    If Not mBar Is Nothing Then notifyBarValues mLatestTimestamp
    Set mBar = mBars.bar(0)
    mBarNumber = mBar.barNumber
    notifyBarValues mLatestTimestamp
    If mPrevBarVolume <> -1 Then mPrevBarVolume = mLatestVolume
    mVolumeCache.addValue 0, mBar.barNumber, mBar.DateTime, mBar.DateTime
    mTotalVolumeCache.addValue 0, mBar.barNumber, mBar.DateTime, mBar.DateTime
End If
End Sub

'================================================================================
' mSession Event Handlers
'================================================================================

Private Sub mSession_SessionStarted( _
                ev As SessionEvent)
mLatestVolume = 0
mPrevBarVolume = 0
Set mBar = Nothing
End Sub

'================================================================================
' Custom methods that must be implemented
'================================================================================

Private Sub customInitialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As session)

Set mSession = pSession

mStudyFoundation.ReplayUsingStudyValues = True
customProcessParameters Parameters
Set mVolumeCache = mStudyFoundation.FindValueCache(BarValueVolume)
Set mTotalVolumeCache = mStudyFoundation.FindValueCache(ConstVolBarsValueTotalVolume)
End Sub

Private Function customGetStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As StudyValue
Dim bar As bar
Dim sv As StudyValue

Set bar = mBars.bar(ref)

sv.barNumber = bar.barNumber
sv.barStartTime = bar.DateTime
sv.timestamp = bar.DateTime

Select Case UCase$(valueName)
Case UCase$(ConstMomentumBarsValueBar)
    Set sv.value = bar
Case UCase$(BarValueOpen)
        sv.value = bar.openValue
Case UCase$(BarValueHigh)
        sv.value = bar.highValue
Case UCase$(BarValueLow)
        sv.value = bar.lowValue
Case UCase$(BarValueClose)
        sv.value = bar.closeValue
Case UCase$(BarValueTickVolume)
        sv.value = bar.tickVolume
Case UCase$(BarValueVolume)
        sv = mVolumeCache.getStudyValue(ref)
Case UCase$(ConstMomentumBarsValueTotalVolume)
        sv = mTotalVolumeCache.getStudyValue(ref)
Case UCase$(BarValueTickVolume)
        sv.value = bar.tickVolume
Case UCase$(BarValueHL2)
        sv.value = (bar.highValue + bar.lowValue) / 2
Case UCase$(BarValueHLC3)
        sv.value = (bar.highValue + bar.lowValue + bar.closeValue) / 3
Case UCase$(BarValueOHLC4)
        sv.value = (bar.openValue + bar.highValue + bar.lowValue + bar.closeValue) / 4
End Select

customGetStudyValue = sv
End Function

Private Sub customNotify( _
                ev As StudyValueEvent)
Dim evOut As StudyValueEvent
Dim i As Long

If IsEmpty(ev.value) Then Exit Sub

mLatestTimestamp = ev.timestamp

If ev.valueName = ConstMomentumBarsInputPrice Then
    mBars.notify ev
    If Not mBar Is Nothing Then notifyBarValues ev.timestamp

ElseIf ev.valueName = ConstMomentumBarsInputTotalVolume Then
    ' we assume that a volume notification belongs to the same bar as the
    ' preceding price notification, so we don't check the timestamp
    mLatestVolume = ev.value
    If mPrevBarVolume < 0 Then
        ' this is the first volume notification after switching to a new
        ' source
        mPrevBarVolume = ev.value - mVolumeCache.getValue(0)
    End If
    'Debug.Assert mLatestVolume >= mPrevBarVolume
    mVolumeCache.updateValue mLatestVolume - mPrevBarVolume
    mTotalVolumeCache.updateValue mLatestVolume
    notifyVolume ev.timestamp
    notifyTotalVolume ev.timestamp
ElseIf ev.valueName = ConstMomentumBarsInputTickVolume Then
    ' we assume that a tickVolume notification belongs to the same bar as the
    ' preceding price notification, so we don't check the timestamp
    mBars.incrementTickVolume ev.value
    notifyTickVolume ev.timestamp
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "CommonStudiesLib" & "." & "ConstVolBars" & ":" & "customNotify", _
            "Invalid input value name: " & ev.valueName

End If

End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = 0
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(ConstMomentumBarsParamTicksPerBar)
        mMovementPerBar = param.value
    End Select
Next

Set mBars = CreateBars(mMovementPerBar, _
                        TimePeriodTickMovement, _
                        mSession, _
                        mPriceTickSize, _
                        mNumberOfValuesToCache)
mBars.addCollectionChangeListener Me
End Sub

Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GConstVolumeBars.StudyDefinition
End Function

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get Bars() As Bars
Set Bars = mBars
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub notifyBarValues( _
                ByVal timestamp As Date)
Dim evOut As StudyValueEvent

evOut.barNumber = mBarNumber
evOut.barStartTime = mBar.DateTime
Set evOut.Source = Me
evOut.timestamp = timestamp

If mBar.barChanged Then
    Set evOut.value = mBar
    evOut.valueName = ConstVolBarsValueBar
    mStudyFoundation.notifyValue evOut
End If

If mBar.openChanged Then
    evOut.value = mBar.openValue
    evOut.valueName = BarValueOpen
    mStudyFoundation.notifyValue evOut
End If

If mBar.highChanged Then
    evOut.value = mBar.highValue
    evOut.valueName = BarValueHigh
    mStudyFoundation.notifyValue evOut
End If

If mBar.lowChanged Then
    evOut.value = mBar.lowValue
    evOut.valueName = BarValueLow
    mStudyFoundation.notifyValue evOut
End If

If mBar.closeChanged Then
    evOut.value = mBar.closeValue
    evOut.valueName = BarValueClose
    mStudyFoundation.notifyValue evOut
End If

If mBar.barChanged Then
    evOut.value = mBar.tickVolume
    evOut.valueName = BarValueTickVolume
    mStudyFoundation.notifyValue evOut
End If

If mBar.highChanged Or mBar.lowChanged Then
    evOut.value = (mBar.highValue + mBar.lowValue) / 2
    evOut.valueName = BarValueHL2
    mStudyFoundation.notifyValue evOut
End If

If mBar.highChanged Or mBar.lowChanged Or mBar.closeChanged Then
    evOut.value = (mBar.highValue + mBar.lowValue + mBar.closeValue) / 3
    evOut.valueName = BarValueHLC3
    mStudyFoundation.notifyValue evOut
End If

If mBar.openChanged Or mBar.highChanged Or mBar.lowChanged Or mBar.closeChanged Then
    evOut.value = (mBar.openValue + mBar.highValue + mBar.lowValue + mBar.closeValue) / 4
    evOut.valueName = BarValueOHLC4
    mStudyFoundation.notifyValue evOut
End If

End Sub

Private Sub notifyTickVolume( _
                ByVal timestamp As Date)
Dim evOut As StudyValueEvent

evOut.barNumber = mBarNumber
evOut.barStartTime = mBar.DateTime
Set evOut.Source = Me
evOut.timestamp = timestamp
evOut.valueName = BarValueTickVolume
evOut.value = mBar.tickVolume
mStudyFoundation.notifyValue evOut
End Sub

Private Sub notifyTotalVolume( _
                ByVal timestamp As Date)
Dim evOut As StudyValueEvent

evOut.barNumber = mBarNumber
evOut.barStartTime = mBar.DateTime
Set evOut.Source = Me
evOut.timestamp = timestamp
evOut.valueName = ConstVolBarsValueTotalVolume
evOut.value = mTotalVolumeCache.getValue(0)
mStudyFoundation.notifyValue evOut
End Sub

Private Sub notifyVolume( _
                ByVal timestamp As Date)
Dim evOut As StudyValueEvent

evOut.barNumber = mBarNumber
evOut.barStartTime = mBar.DateTime
Set evOut.Source = Me
evOut.timestamp = timestamp
evOut.valueName = BarValueVolume
evOut.value = mVolumeCache.getValue(0)
mStudyFoundation.notifyValue evOut
End Sub

Private Function numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
numberOfCachedValues = mBars.Count
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
                Optional ByVal taskData As Variant) As TaskController

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
                Optional ByVal taskData As Variant) As TaskController
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
'Study_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)
Study_getStudyValue = customGetStudyValue(valueName, ref)
End Function

Private Function Study_getValueTicksize( _
                ByVal valueName As String) As Double
Select Case UCase$(valueName)
Case UCase$(BarValueOpen), _
    UCase$(BarValueHigh), _
    UCase$(BarValueLow), _
    UCase$(BarValueClose)
    Study_getValueTicksize = mPriceTickSize
Case UCase$(BarValueVolume)
    Study_getValueTicksize = mVolumeTickSize
Case UCase$(BarValueTickVolume)
    Study_getValueTicksize = 1
End Select
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
If valueName = "" Or valueName = ConstVolBarsValueBar Then
    Study_numberOfCachedValues = mBars.Count
Else
    Study_numberOfCachedValues = numberOfCachedValues(valueName)
End If
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
If UCase$(inputName) = UCase$(ConstVolBarsInputPrice) Then
    mPriceTickSize = tickSize
ElseIf UCase$(inputName) = UCase$(ConstVolBarsInputTotalVolume) Then
    mVolumeTickSize = tickSize
End If
End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Set Study_studyDefinition = customStudyDefinition
End Property













