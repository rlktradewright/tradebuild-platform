VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstTimeBars"
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
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements Study
Implements BarStudy

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

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

Private mSession As session
Attribute mSession.VB_VarHelpID = -1

Private mBarLength As Long

Private mTimeUnits As TimePeriodUnits

Private mBars As Bars
Private mBar As bar

Private mNumberOfValuesToCache As Long

Private mBarNumber As Long

Private mVolParser As VolumeParser

Private mCurrBarVolume As Long

Private mVolumeCache As StudyValueCache

Private mPriceTickSize As Double
Private mVolumeTickSize As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "Constant Time Bars study terminated"
End Sub

'@================================================================================
' BarStudy Interface Members
'@================================================================================

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
Set mSession = newUnderlyingStudy.session
mVolParser.changeSession mSession
mBars.changeSession mSession
End Sub

'@================================================================================
' Custom methods that must be implemented
'@================================================================================

Private Sub customInitialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As session)

Set mSession = pSession
Set mVolParser = CreateVolumeParser(pSession)

mStudyFoundation.ReplayUsingStudyValues = True
customProcessParameters Parameters
Set mVolumeCache = mStudyFoundation.FindValueCache(BarValueVolume)
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
Case UCase$(ConstTimeBarsValueBar)
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

If ev.valueName = ConstTimeBarsInputPrice Then
    mBars.notify ev
    If mBars.currentBarNumber <> mBarNumber Then
        ' a new bar has started
        Set mBar = mBars.bar(0)
        mBarNumber = mBars.currentBarNumber
        mCurrBarVolume = 0
        mVolumeCache.addValue 0, mBarNumber, ev.timestamp, mBar.DateTime
    End If
    notifyBarValues ev.timestamp
ElseIf ev.valueName = ConstTimeBarsInputTotalVolume Then
    ' we assume that a volume notification belongs to the same bar as the
    ' preceding price notification, so we don't check the timestamp
    
    If mVolumeCache.Count <> 0 Then
        mCurrBarVolume = mCurrBarVolume + mVolParser.notify(ev.value)
        mVolumeCache.updateValue mCurrBarVolume
        notifyVolume ev.timestamp
    End If
    
ElseIf ev.valueName = ConstTimeBarsInputTickVolume Then
    ' we assume that a tickVolume notification belongs to the same bar as the
    ' preceding price notification, so we don't check the timestamp
    
    ' We also assume that a tickvolume notification represents the actual
    ' current tickvolume for the bar, and therefore replaces the bar's
    ' tickvolume rather than incrementing it
    mBars.incrementTickVolume ev.value - mBar.tickVolume
    notifyTickVolume ev.timestamp
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "CommonStudiesLib" & "." & "ConstTimeBars" & ":" & "customNotify", _
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
    Case UCase$(ConstTimeBarsParamBarLength)
        mBarLength = param.value
    Case UCase$(ConstTimeBarsParamTimeUnits)
        mTimeUnits = TimePeriodUnitsFromString(param.value)
    End Select
Next

Set mBars = CreateBars(mBarLength, _
                                mTimeUnits, _
                                mSession, _
                                mPriceTickSize, _
                                mNumberOfValuesToCache)
End Sub

Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GConstTimeBars.StudyDefinition
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Bars() As Bars
Set Bars = mBars
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

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

Private Sub notifyVolume( _
                ByVal timestamp As Date)
Dim evOut As StudyValueEvent

evOut.barNumber = mBarNumber
evOut.barStartTime = mBar.DateTime
Set evOut.Source = Me
evOut.timestamp = timestamp
evOut.valueName = BarValueVolume
evOut.value = mCurrBarVolume
mStudyFoundation.notifyValue evOut
End Sub

Private Function numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
numberOfCachedValues = mBars.Count
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
If valueName = "" Or valueName = ConstTimeBarsValueBar Then
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
If UCase$(inputName) = UCase$(ConstTimeBarsInputPrice) Then
    mPriceTickSize = tickSize
ElseIf UCase$(inputName) = UCase$(ConstTimeBarsInputTotalVolume) Then
    mVolumeTickSize = tickSize
End If
End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Set Study_studyDefinition = customStudyDefinition
End Property









