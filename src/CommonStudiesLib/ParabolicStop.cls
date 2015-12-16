VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParabolicStop"
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

Private Const ModuleName                As String = "ParabolicStop"

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

Private mInitialStopForBar As Double
Private mCurrentStop As Double
Private mPsCache As ValueCache
Private mExtremePoint As Double
Private mSignificantPoint As Double

Private mAccelerationFactor As Double
Private mStartFactor As Double
Private mIncrement As Double
Private mMaxFactor As Double

Private mLong As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mCurrentStop = MaxDouble
mLong = True    ' arbitrary initial setting
End Sub

Private Sub Class_Terminate()
Debug.Print "Parabolic Stop terminated"
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

customProcessParameters Parameters
Set mPsCache = mStudyFoundation.FindValueCache(PsValuePs)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub customNotify( _
                ByRef ev As StudyValueEventData)
Dim newValue As Double
Dim evOut As StudyValueEventData
Static accelAdjustedThisbar As Boolean
Static sPrevVal As Variant

Const ProcName As String = "customNotify"
On Error GoTo Err

If IsEmpty(ev.sVal.Value) Then Exit Sub

If Not mStudyFoundation.IsNewBar Then
    If ev.sVal.Value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.Value

newValue = CDbl(ev.sVal.Value)

If mCurrentStop = MaxDouble Then
    mCurrentStop = newValue * 0.999
    mExtremePoint = newValue * 1.01
    mInitialStopForBar = mCurrentStop
End If

If mStudyFoundation.IsNewBar Then
    accelAdjustedThisbar = False
    mInitialStopForBar = mCurrentStop
    mPsCache.AddValue mCurrentStop, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
    If mAccelerationFactor = 0 Then
        mAccelerationFactor = mStartFactor
    End If
    If mLong Then
        mCurrentStop = mInitialStopForBar + mAccelerationFactor * (mExtremePoint - mInitialStopForBar)
    Else
        mCurrentStop = mInitialStopForBar - mAccelerationFactor * (mInitialStopForBar - mExtremePoint)
    End If
    mPsCache.UpdateValue mCurrentStop, ev.sVal.BarNumber, ev.sVal.Timestamp
End If

If mPsCache.Count < 2 Then Exit Sub
        
If mLong Then
    If newValue > mCurrentStop Then
        If newValue > mExtremePoint Then
            mExtremePoint = newValue
            If Not accelAdjustedThisbar And mAccelerationFactor <> 0 Then
                If Abs(mMaxFactor - mAccelerationFactor) > 0.001 Then
                    mAccelerationFactor = mAccelerationFactor + mIncrement
                    accelAdjustedThisbar = True
                End If
            End If
            mCurrentStop = mInitialStopForBar + mAccelerationFactor * (mExtremePoint - mInitialStopForBar)
            mPsCache.UpdateValue mCurrentStop, ev.sVal.BarNumber, ev.sVal.Timestamp
        End If
        If newValue > mSignificantPoint Then mSignificantPoint = newValue
    ElseIf newValue > mPsCache.GetValue(-1) Then
        mCurrentStop = mPsCache.GetValue(-1)
    Else
        mLong = False
        mCurrentStop = mSignificantPoint
        mInitialStopForBar = mCurrentStop
        mPsCache.UpdateValue mCurrentStop, ev.sVal.BarNumber, ev.sVal.Timestamp
        mAccelerationFactor = 0
        mExtremePoint = newValue
        mSignificantPoint = newValue
    End If
Else
    If newValue < mCurrentStop Then
        If newValue < mExtremePoint Then
            mExtremePoint = newValue
            If Not accelAdjustedThisbar And mAccelerationFactor <> 0 Then
                If Abs(mMaxFactor - mAccelerationFactor) > 0.001 Then
                    mAccelerationFactor = mAccelerationFactor + mIncrement
                    accelAdjustedThisbar = True
                End If
            End If
            mCurrentStop = mInitialStopForBar - mAccelerationFactor * (mInitialStopForBar - mExtremePoint)
            mPsCache.UpdateValue mCurrentStop, ev.sVal.BarNumber, ev.sVal.Timestamp
        End If
        If newValue < mSignificantPoint Then mSignificantPoint = newValue
    ElseIf newValue < mPsCache.GetValue(-1) Then
        mCurrentStop = mPsCache.GetValue(-1)
    Else
        mLong = True
        mCurrentStop = mSignificantPoint
        mInitialStopForBar = mCurrentStop
        mPsCache.UpdateValue mCurrentStop, ev.sVal.BarNumber, ev.sVal.Timestamp
        mAccelerationFactor = 0
        mExtremePoint = newValue
        mSignificantPoint = newValue
    End If
End If

Set evOut.Source = Me
evOut.sVal.BarNumber = ev.sVal.BarNumber
evOut.sVal.BarStartTime = ev.sVal.BarStartTime
evOut.sVal.Timestamp = ev.sVal.Timestamp
evOut.sVal.Value = mCurrentStop
evOut.valueName = PsValuePs
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = 100 ' difficult to say!

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
    Case UCase$(PsParamStartFactor)
        mStartFactor = param.Value
    Case UCase$(PsParamIncrement)
        mIncrement = param.Value
    Case UCase$(PsParamMaxFactor)
        mMaxFactor = param.Value
    End Select
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GParabolicStop.StudyDefinition

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

Public Property Get ps(Optional ByVal ref As Integer) As SValue
Const ProcName As String = "ps"
On Error GoTo Err

ps = mPsCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Sub setStop(ByVal stopValue As Double, _
                    ByVal isLong As Boolean, _
                    ByVal extremePoint As Double)
Const ProcName As String = "setStop"
On Error GoTo Err

mLong = isLong
mCurrentStop = stopValue
mInitialStopForBar = mCurrentStop
mPsCache.UpdateValue mCurrentStop
mAccelerationFactor = 0
mExtremePoint = extremePoint
mSignificantPoint = extremePoint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


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
Const ProcName As String = "IStudy_GetStudyValue"
On Error GoTo Err

IStudy_GetStudyValue = mStudyFoundation.getStudyValue(valueName, ref)

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

IStudy_NumberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)

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




