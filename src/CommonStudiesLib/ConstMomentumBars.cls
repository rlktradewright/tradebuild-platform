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

Private Const ProjectName                As String = "CmnStudiesLib26"
Private Const ModuleName                As String = "ConstMomentumBars"

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

Private mSession As Session
Attribute mSession.VB_VarHelpID = -1

Private mMovementPerBar As Long

Private mBarsBuilder As BarsBuilder
Private WithEvents mBars As Bars
Attribute mBars.VB_VarHelpID = -1
Private mBar As bar

Private mNumberOfValuesToCache As Long

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
Set mSession = newUnderlyingStudy.Session
mBarsBuilder.changeSession mSession
End Sub

'@================================================================================
' mBars Event Handlers
'@================================================================================

Private Sub mBars_CollectionChanged(ev As TWUtilities30.CollectionChangeEvent)
If ev.changeType = CollItemAdded Then
    If Not mBar Is Nothing Then notifyBarValues mBar.timestamp
    Set mBar = mBars.bar(0)
    notifyBarValues mBar.timestamp
End If
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
                ByVal pSession As Session)

Set mSession = pSession

mStudyFoundation.ReplayUsingStudyValues = True
customProcessParameters Parameters
End Sub

Private Function customGetStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Dim bar As bar
Dim sv As SValue

Set bar = mBars.bar(ref)

sv.barNumber = bar.barNumber
sv.barStartTime = bar.timestamp
sv.timestamp = bar.timestamp

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
        sv.value = bar.volume
Case UCase$(BarValueOpenInterest)
        sv.value = bar.openInterest
Case UCase$(BarValueHL2)
        sv.value = bar.HL2
Case UCase$(BarValueHLC3)
        sv.value = bar.HLC3
Case UCase$(BarValueOHLC4)
        sv.value = bar.OHLC4
End Select

customGetStudyValue = sv
End Function

Private Sub customNotify( _
                ev As StudyValueEvent)

If IsEmpty(ev.sVal.value) Then Exit Sub

If ev.valueName = ConstMomentumBarsInputPrice Then
    mBarsBuilder.notifyValue ev.sVal
    notifyBarValues ev.sVal.timestamp
ElseIf ev.valueName = ConstMomentumBarsInputTotalVolume Then
    mBarsBuilder.notifyVolume ev.sVal
    notifyBarValues ev.sVal.timestamp
ElseIf ev.valueName = ConstMomentumBarsInputTickVolume Then
    mBarsBuilder.incrementTickVolume ev.sVal.value
    notifyBarValues ev.sVal.timestamp
ElseIf ev.valueName = ConstMomentumBarsInputOpenInterest Then
    mBarsBuilder.notifyOpenInterest ev.sVal
    notifyBarValues ev.sVal.timestamp
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "customNotify", _
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

Set mBarsBuilder = CreateBarsBuilder(GetTimePeriod(mMovementPerBar, TimePeriodTickMovement), _
                                    mSession, _
                                    mPriceTickSize, _
                                    mNumberOfValuesToCache)
Set mBars = mBarsBuilder.Bars
End Sub

Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GConstMomentumBars.StudyDefinition
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

If mBar Is Nothing Then Exit Sub

evOut.sVal.barNumber = mBars.currentBarNumber
evOut.sVal.barStartTime = mBar.timestamp
Set evOut.Source = Me
evOut.sVal.timestamp = timestamp

If mBar.barChanged Then
    Set evOut.sVal.value = mBar
    evOut.valueName = ConstMomentumBarsValueBar
    mStudyFoundation.notifyValue evOut
End If

If mBar.openChanged Then
    evOut.sVal.value = mBar.openValue
    evOut.valueName = BarValueOpen
    mStudyFoundation.notifyValue evOut
End If

If mBar.highChanged Then
    evOut.sVal.value = mBar.highValue
    evOut.valueName = BarValueHigh
    mStudyFoundation.notifyValue evOut
End If

If mBar.lowChanged Then
    evOut.sVal.value = mBar.lowValue
    evOut.valueName = BarValueLow
    mStudyFoundation.notifyValue evOut
End If

If mBar.closeChanged Then
    evOut.sVal.value = mBar.closeValue
    evOut.valueName = BarValueClose
    mStudyFoundation.notifyValue evOut
End If

If mBar.volumeChanged Then
    evOut.sVal.value = mBar.volume
    evOut.valueName = BarValueVolume
    mStudyFoundation.notifyValue evOut
End If

If mBar.openInterestChanged Then
    evOut.sVal.value = mBar.openInterest
    evOut.valueName = BarValueOpenInterest
    mStudyFoundation.notifyValue evOut
End If

If mBar.barChanged Then
    evOut.sVal.value = mBar.tickVolume
    evOut.valueName = BarValueTickVolume
    mStudyFoundation.notifyValue evOut
End If

If mBar.highChanged Or mBar.lowChanged Then
    evOut.sVal.value = mBar.HL2
    evOut.valueName = BarValueHL2
    mStudyFoundation.notifyValue evOut
End If

If mBar.highChanged Or mBar.lowChanged Or mBar.closeChanged Then
    evOut.sVal.value = mBar.HLC3
    evOut.valueName = BarValueHLC3
    mStudyFoundation.notifyValue evOut
End If

If mBar.openChanged Or mBar.highChanged Or mBar.lowChanged Or mBar.closeChanged Then
    evOut.sVal.value = mBar.OHLC4
    evOut.valueName = BarValueOHLC4
    mStudyFoundation.notifyValue evOut
End If

mBarsBuilder.resetBarChangeIndicators
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
                ByVal ref As Long) As SValue
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
                ByVal pSession As Session)
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
If valueName = "" Or valueName = ConstMomentumBarsValueBar Then
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
If UCase$(inputName) = UCase$(ConstMomentumBarsInputPrice) Then
    mPriceTickSize = tickSize
ElseIf UCase$(inputName) = UCase$(ConstMomentumBarsInputTotalVolume) Then
    mVolumeTickSize = tickSize
End If
End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Set Study_studyDefinition = customStudyDefinition
End Property













