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

Implements IStudy
Implements IBarStudy

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

Private mSessionFuture As IFuture

Private WithEvents mBarStudySupport As BarStudySupport
Attribute mBarStudySupport.VB_VarHelpID = -1

Private WithEvents mBars As Bars
Attribute mBars.VB_VarHelpID = -1
Private WithEvents mBar As BarUtils27.Bar
Attribute mBar.VB_VarHelpID = -1

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
' IBarStudy Interface Members
'@================================================================================

Private Property Get IBarStudy_BarsFuture() As IFuture
Const ProcName As String = "IBarStudy_BarsFuture"
On Error GoTo Err

Set IBarStudy_BarsFuture = mBarStudySupport.BarsFuture

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IBarStudy_CloseValueName() As String
Const ProcName As String = "IBarStudy_CloseValueName"
On Error GoTo Err

IBarStudy_CloseValueName = BarStudyValueClose

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IBarStudy_HighValueName() As String
Const ProcName As String = "IBarStudy_HighValueName"
On Error GoTo Err

IBarStudy_HighValueName = BarStudyValueHigh

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IBarStudy_InitialBarFetchListener() As IBarFetchListener
Const ProcName As String = "IBarStudy_InitialBarFetchListener"
On Error GoTo Err

Set IBarStudy_InitialBarFetchListener = mBarStudySupport.InitialBarFetchListener

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Let IBarStudy_InitialBarsFuture(ByVal Value As IFuture)
Const ProcName As String = "IBarStudy_InitialBarsFuture"
On Error GoTo Err

mBarStudySupport.InitialBarsFuture = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IBarStudy_LowValueName() As String
Const ProcName As String = "IBarStudy_LowValueName"
On Error GoTo Err

IBarStudy_LowValueName = BarStudyValueLow

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IBarStudy_OpenValueName() As String
Const ProcName As String = "IBarStudy_OpenValueName"
On Error GoTo Err

IBarStudy_OpenValueName = BarStudyValueOpen

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' mBar Event Handlers
'@================================================================================

Private Sub mBar_Changed(ev As ChangeEventData)
Const ProcName As String = "mBar_Changed"
On Error GoTo Err

notifyBarValues

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mBars Event Handlers
'@================================================================================

Private Sub mBars_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mBars_CollectionChanged"
On Error GoTo Err

If ev.ChangeType = CollItemAdded Then
    Set mBar = mBars.Bar(0)
    mTimestamp = mBar.Timestamp
    notifyBarValues
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mBarStudySupport Event Handlers
'@================================================================================

Private Sub mBarStudySupport_BarsAvailable(ByVal pBars As Bars)
Const ProcName As String = "mBarStudySupport_BarsAvailable"
On Error GoTo Err

Set mBars = pBars

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Custom methods that must be implemented
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

Set mSessionFuture = pSessionFuture

mStudyFoundation.ReplayUsingStudyValues = True
customProcessParameters Parameters

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function customGetStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Dim Bar As BarUtils27.Bar
Dim sv As SValue

Const ProcName As String = "customGetStudyValue"
On Error GoTo Err

Set Bar = mBars.Bar(ref)

sv.BarNumber = Bar.BarNumber
sv.BarStartTime = Bar.Timestamp
sv.Timestamp = Bar.Timestamp

Select Case UCase$(valueName)
Case UCase$(ConstVolBarsValueBar)
    Set sv.Value = Bar
Case UCase$(BarStudyValueOpen)
        sv.Value = Bar.OpenValue
Case UCase$(BarStudyValueHigh)
        sv.Value = Bar.highValue
Case UCase$(BarStudyValueLow)
        sv.Value = Bar.lowValue
Case UCase$(BarStudyValueClose), UCase$(DefaultStudyValueName)
        sv.Value = Bar.CloseValue
Case UCase$(BarStudyValueTickVolume)
        sv.Value = Bar.TickVolume
Case UCase$(BarStudyValueVolume)
        sv.Value = Bar.volume
Case UCase$(BarStudyValueOpenInterest)
        sv.Value = Bar.OpenInterest
Case UCase$(BarStudyValueHL2)
        sv.Value = Bar.HL2
Case UCase$(BarStudyValueHLC3)
        sv.Value = Bar.HLC3
Case UCase$(BarStudyValueOHLC4)
        sv.Value = Bar.OHLC4
End Select

customGetStudyValue = sv

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub customNotify( _
                ByRef ev As StudyValueEventData)

Const ProcName As String = "customNotify"
On Error GoTo Err

If IsEmpty(ev.sVal.Value) Then Exit Sub

mTimestamp = ev.sVal.Timestamp

ev.valueName = UCase$(ev.valueName)
If ev.valueName = ConstMomentumBarsInputPriceUcase Then
    mBarStudySupport.BarsBuilder.notifyValue ev.sVal
ElseIf ev.valueName = ConstMomentumBarsInputTotalVolumeUcase Then
    mBarStudySupport.BarsBuilder.NotifyVolume ev.sVal
ElseIf ev.valueName = ConstMomentumBarsInputTickVolumeUcase Then
    mBarStudySupport.BarsBuilder.IncrementTickVolume ev.sVal.Value
ElseIf ev.valueName = ConstVolBarsInputOpenInterestUcase Then
    mBarStudySupport.BarsBuilder.NotifyOpenInterest ev.sVal
Else
    AssertArgument False, "Invalid input Value name: " & ev.valueName
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

Const ProcName As String = "customProcessParameters"
On Error GoTo Err

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(ConstVolumeBarsParamVolPerBar)
        Dim lVolumePerBar As Long
        lVolumePerBar = param.Value
        AssertArgument lVolumePerBar > 0, "Volume per bar parameter must be > 0"
    End Select
Next

Set mBarStudySupport = CreateBarStudySupport(GetTimePeriod(lVolumePerBar, TimePeriodVolume), mSessionFuture, mPriceTickSize)
If mBarStudySupport.BarsFuture.IsAvailable Then Set mBars = mBarStudySupport.BarsFuture.Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GConstVolumeBars.StudyDefinition

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

Public Property Get BarsFuture() As IFuture
Const ProcName As String = "BarsFuture"
On Error GoTo Err

Set BarsFuture = mBarStudySupport.BarsFuture

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

Private Sub notifyBarValues()
Dim evOut As StudyValueEventData

Const ProcName As String = "notifyBarValues"
On Error GoTo Err

If mBar Is Nothing Then Exit Sub

evOut.sVal.BarNumber = mBar.BarNumber
evOut.sVal.BarStartTime = mBar.Timestamp
Set evOut.Source = Me
evOut.sVal.Timestamp = mTimestamp

If mBar.BarChanged Then
    Set evOut.sVal.Value = mBar
    evOut.valueName = ConstVolBarsValueBar
    mStudyFoundation.notifyValue evOut
End If

If mBar.OpenChanged Then
    evOut.sVal.Value = mBar.OpenValue
    evOut.valueName = BarStudyValueOpen
    mStudyFoundation.notifyValue evOut
End If

If mBar.HighChanged Then
    evOut.sVal.Value = mBar.highValue
    evOut.valueName = BarStudyValueHigh
    mStudyFoundation.notifyValue evOut
End If

If mBar.LowChanged Then
    evOut.sVal.Value = mBar.lowValue
    evOut.valueName = BarStudyValueLow
    mStudyFoundation.notifyValue evOut
End If

If mBar.CloseChanged Then
    evOut.sVal.Value = mBar.CloseValue
    evOut.valueName = BarStudyValueClose
    mStudyFoundation.notifyValue evOut
End If

If mBar.VolumeChanged Then
    evOut.sVal.Value = mBar.volume
    evOut.valueName = BarStudyValueVolume
    mStudyFoundation.notifyValue evOut
End If

If mBar.OpenInterestChanged Then
    evOut.sVal.Value = mBar.OpenInterest
    evOut.valueName = BarStudyValueOpenInterest
    mStudyFoundation.notifyValue evOut
End If

If mBar.BarChanged Then
    evOut.sVal.Value = mBar.TickVolume
    evOut.valueName = BarStudyValueTickVolume
    mStudyFoundation.notifyValue evOut
End If

If mBar.HighChanged Or mBar.LowChanged Then
    evOut.sVal.Value = mBar.HL2
    evOut.valueName = BarStudyValueHL2
    mStudyFoundation.notifyValue evOut
End If

If mBar.HighChanged Or mBar.LowChanged Or mBar.CloseChanged Then
    evOut.sVal.Value = mBar.HLC3
    evOut.valueName = BarStudyValueHLC3
    mStudyFoundation.notifyValue evOut
End If

If mBar.OpenChanged Or mBar.HighChanged Or mBar.LowChanged Or mBar.CloseChanged Then
    evOut.sVal.Value = mBar.OHLC4
    evOut.valueName = BarStudyValueOHLC4
    mStudyFoundation.notifyValue evOut
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

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
'IStudy_GetStudyValue = mStudyFoundation.getStudyValue(valueName, ref)
Const ProcName As String = "IStudy_GetStudyValue"
On Error GoTo Err

IStudy_GetStudyValue = customGetStudyValue(valueName, ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudy_GetValueTicksize( _
                ByVal valueName As String) As Double
Const ProcName As String = "IStudy_GetValueTicksize"
On Error GoTo Err

Select Case UCase$(valueName)
Case UCase$(BarStudyValueOpen), _
    UCase$(BarStudyValueHigh), _
    UCase$(BarStudyValueLow), _
    UCase$(BarStudyValueClose), _
    UCase$(DefaultStudyValueName)
    IStudy_GetValueTicksize = mPriceTickSize
Case UCase$(BarStudyValueVolume)
    IStudy_GetValueTicksize = mVolumeTickSize
Case UCase$(BarStudyValueTickVolume)
    IStudy_GetValueTicksize = 1
End Select

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
                ByRef ev As StudyValueEventData)
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

If valueName = "" Or valueName = ConstVolBarsValueBar Then
    If mBars Is Nothing Then
        IStudy_NumberOfCachedValues = 0
    Else
        IStudy_NumberOfCachedValues = mBars.Count
    End If
Else
    IStudy_NumberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)
End If

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

If UCase$(inputName) = UCase$(ConstVolBarsInputPrice) Then
    mPriceTickSize = tickSize
ElseIf UCase$(inputName) = UCase$(ConstVolBarsInputTotalVolume) Then
    mVolumeTickSize = tickSize
End If

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











