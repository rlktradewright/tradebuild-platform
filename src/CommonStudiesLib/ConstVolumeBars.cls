VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstVolumeBars"
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

Private Const ModuleName                As String = "ConstVolumeBars"

Private Const NoPrice                   As Double = MaxDouble

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

Private WithEvents mSession As Session
Attribute mSession.VB_VarHelpID = -1

Private mVolumePerBar As Long

Private mBarsBuilder As BarsBuilder
Private WithEvents mBars As Bars
Attribute mBars.VB_VarHelpID = -1
Private WithEvents mBar As bar
Attribute mBar.VB_VarHelpID = -1

Private mNumberOfValuesToCache As Long

Private mPriceTickSize As Double
Private mVolumeTickSize As Double

Private mTimestamp As Date

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
Const ProcName As String = "BarStudy_bars"
On Error GoTo Err

Set BarStudy_bars = mBars

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BarStudy_closeValueName() As String
Const ProcName As String = "BarStudy_closeValueName"
On Error GoTo Err

BarStudy_closeValueName = BarValueClose

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BarStudy_highValueName() As String
Const ProcName As String = "BarStudy_highValueName"
On Error GoTo Err

BarStudy_highValueName = BarValueHigh

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BarStudy_lowValueName() As String
Const ProcName As String = "BarStudy_lowValueName"
On Error GoTo Err

BarStudy_lowValueName = BarValueLow

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BarStudy_openValueName() As String
Const ProcName As String = "BarStudy_openValueName"
On Error GoTo Err

BarStudy_openValueName = BarValueOpen

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub BarStudy_switchUnderlyingStudy( _
                ByVal newUnderlyingStudy As InputStudy)
Const ProcName As String = "BarStudy_switchUnderlyingStudy"
On Error GoTo Err

Set mSession = newUnderlyingStudy.Session
mBarsBuilder.ChangeSession mSession

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' mBar Event Handlers
'@================================================================================

Private Sub mBar_Changed(ev As TWUtilities30.ChangeEvent)
Const ProcName As String = "mBar_Changed"
On Error GoTo Err

notifyBarValues

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' mBars Event Handlers
'@================================================================================

Private Sub mBars_CollectionChanged(ev As TWUtilities30.CollectionChangeEvent)
Const ProcName As String = "mBars_CollectionChanged"
On Error GoTo Err

If ev.changeType = CollItemAdded Then
    Set mBar = mBars.bar(0)
    notifyBarValues
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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

Const ProcName As String = "customInitialise"
On Error GoTo Err

Set mSession = pSession

mStudyFoundation.ReplayUsingStudyValues = True
customProcessParameters Parameters

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function customGetStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Dim bar As bar
Dim sv As SValue

Const ProcName As String = "customGetStudyValue"
On Error GoTo Err

Set bar = mBars.bar(ref)

sv.barNumber = bar.barNumber
sv.barStartTime = bar.timestamp
sv.timestamp = bar.timestamp

Select Case UCase$(valueName)
Case UCase$(ConstVolBarsValueBar)
    Set sv.value = bar
Case UCase$(BarValueOpen)
        sv.value = bar.OpenValue
Case UCase$(BarValueHigh)
        sv.value = bar.highValue
Case UCase$(BarValueLow)
        sv.value = bar.lowValue
Case UCase$(BarValueClose), UCase$(DefaultStudyValueName)
        sv.value = bar.CloseValue
Case UCase$(BarValueTickVolume)
        sv.value = bar.TickVolume
Case UCase$(BarValueVolume)
        sv.value = bar.volume
Case UCase$(BarValueOpenInterest)
        sv.value = bar.OpenInterest
Case UCase$(BarValueHL2)
        sv.value = bar.HL2
Case UCase$(BarValueHLC3)
        sv.value = bar.HLC3
Case UCase$(BarValueOHLC4)
        sv.value = bar.OHLC4
End Select

customGetStudyValue = sv

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub customNotify( _
                ev As StudyValueEvent)

Const ProcName As String = "customNotify"
On Error GoTo Err

If IsEmpty(ev.sVal.value) Then Exit Sub

mTimestamp = ev.sVal.timestamp

ev.valueName = UCase$(ev.valueName)
If ev.valueName = ConstMomentumBarsInputPriceUcase Then
    mBarsBuilder.notifyValue ev.sVal
ElseIf ev.valueName = ConstMomentumBarsInputTotalVolumeUcase Then
    mBarsBuilder.NotifyVolume ev.sVal
ElseIf ev.valueName = ConstMomentumBarsInputTickVolumeUcase Then
    mBarsBuilder.IncrementTickVolume ev.sVal.value
ElseIf ev.valueName = ConstVolBarsInputOpenInterestUcase Then
    mBarsBuilder.NotifyOpenInterest ev.sVal
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid input value name: " & ev.valueName
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = 0

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

Const ProcName As String = "customProcessParameters"
On Error GoTo Err

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(ConstVolBarsParamVolPerBar)
        mVolumePerBar = param.value
        If mVolumePerBar <= 0 Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Volume per bar parameter must be > 0"
        End If
    End Select
Next

Set mBarsBuilder = CreateBarsBuilder(GetTimePeriod(mVolumePerBar, TimePeriodVolume), _
                                    mSession, _
                                    mPriceTickSize, _
                                    mNumberOfValuesToCache)
Set mBars = mBarsBuilder.Bars

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GConstVolumeBars.StudyDefinition

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

Public Property Get Bars() As Bars
Const ProcName As String = "Bars"
On Error GoTo Err

Set Bars = mBars

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

Private Sub notifyBarValues()
Dim evOut As StudyValueEvent

Const ProcName As String = "notifyBarValues"
On Error GoTo Err

If mBar Is Nothing Then Exit Sub

evOut.sVal.barNumber = mBar.barNumber
evOut.sVal.barStartTime = mBar.timestamp
Set evOut.Source = Me
evOut.sVal.timestamp = mTimestamp

If mBar.BarChanged Then
    Set evOut.sVal.value = mBar
    evOut.valueName = ConstVolBarsValueBar
    mStudyFoundation.notifyValue evOut
End If

If mBar.OpenChanged Then
    evOut.sVal.value = mBar.OpenValue
    evOut.valueName = BarValueOpen
    mStudyFoundation.notifyValue evOut
End If

If mBar.HighChanged Then
    evOut.sVal.value = mBar.highValue
    evOut.valueName = BarValueHigh
    mStudyFoundation.notifyValue evOut
End If

If mBar.LowChanged Then
    evOut.sVal.value = mBar.lowValue
    evOut.valueName = BarValueLow
    mStudyFoundation.notifyValue evOut
End If

If mBar.CloseChanged Then
    evOut.sVal.value = mBar.CloseValue
    evOut.valueName = BarValueClose
    mStudyFoundation.notifyValue evOut
End If

If mBar.VolumeChanged Then
    evOut.sVal.value = mBar.volume
    evOut.valueName = BarValueVolume
    mStudyFoundation.notifyValue evOut
End If

If mBar.OpenInterestChanged Then
    evOut.sVal.value = mBar.OpenInterest
    evOut.valueName = BarValueOpenInterest
    mStudyFoundation.notifyValue evOut
End If

If mBar.BarChanged Then
    evOut.sVal.value = mBar.TickVolume
    evOut.valueName = BarValueTickVolume
    mStudyFoundation.notifyValue evOut
End If

If mBar.HighChanged Or mBar.LowChanged Then
    evOut.sVal.value = mBar.HL2
    evOut.valueName = BarValueHL2
    mStudyFoundation.notifyValue evOut
End If

If mBar.HighChanged Or mBar.LowChanged Or mBar.CloseChanged Then
    evOut.sVal.value = mBar.HLC3
    evOut.valueName = BarValueHLC3
    mStudyFoundation.notifyValue evOut
End If

If mBar.OpenChanged Or mBar.HighChanged Or mBar.LowChanged Or mBar.CloseChanged Then
    evOut.sVal.value = mBar.OHLC4
    evOut.valueName = BarValueOHLC4
    mStudyFoundation.notifyValue evOut
End If

mBarsBuilder.ResetBarChangeIndicators

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Function numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "numberOfCachedValues"
On Error GoTo Err

numberOfCachedValues = mBars.Count

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
'Study_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)
Const ProcName As String = "Study_getStudyValue"
On Error GoTo Err

Study_getStudyValue = customGetStudyValue(valueName, ref)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function Study_getValueTicksize( _
                ByVal valueName As String) As Double
Const ProcName As String = "Study_getValueTicksize"
On Error GoTo Err

Select Case UCase$(valueName)
Case UCase$(BarValueOpen), _
    UCase$(BarValueHigh), _
    UCase$(BarValueLow), _
    UCase$(BarValueClose), _
    UCase$(DefaultStudyValueName)
    Study_getValueTicksize = mPriceTickSize
Case UCase$(BarValueVolume)
    Study_getValueTicksize = mVolumeTickSize
Case UCase$(BarValueTickVolume)
    Study_getValueTicksize = 1
End Select

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

If valueName = "" Or valueName = ConstVolBarsValueBar Then
    Study_numberOfCachedValues = mBars.Count
Else
    Study_numberOfCachedValues = numberOfCachedValues(valueName)
End If

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

If UCase$(inputName) = UCase$(ConstVolBarsInputPrice) Then
    mPriceTickSize = tickSize
ElseIf UCase$(inputName) = UCase$(ConstVolBarsInputTotalVolume) Then
    mVolumeTickSize = tickSize
End If

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











