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

Implements Study

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
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)

Const ProcName As String = "customInitialise"
On Error GoTo Err

customProcessParameters Parameters
Set mPsCache = mStudyFoundation.FindValueCache(PsValuePs)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Dim newValue As Double
Dim evOut As StudyValueEvent
Static accelAdjustedThisbar As Boolean
Static sPrevVal As Variant

Const ProcName As String = "customNotify"
On Error GoTo Err

If IsEmpty(ev.sVal.value) Then Exit Sub

If Not mStudyFoundation.IsNewBar Then
    If ev.sVal.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.value

newValue = CDbl(ev.sVal.value)

If mCurrentStop = MaxDouble Then
    mCurrentStop = newValue * 0.999
    mExtremePoint = newValue * 1.01
    mInitialStopForBar = mCurrentStop
End If

If mStudyFoundation.IsNewBar Then
    accelAdjustedThisbar = False
    mInitialStopForBar = mCurrentStop
    mPsCache.AddValue mCurrentStop, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
    If mAccelerationFactor = 0 Then
        mAccelerationFactor = mStartFactor
    End If
    If mLong Then
        mCurrentStop = mInitialStopForBar + mAccelerationFactor * (mExtremePoint - mInitialStopForBar)
    Else
        mCurrentStop = mInitialStopForBar - mAccelerationFactor * (mInitialStopForBar - mExtremePoint)
    End If
    mPsCache.UpdateValue mCurrentStop, ev.sVal.barNumber, ev.sVal.timestamp
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
            mPsCache.UpdateValue mCurrentStop, ev.sVal.barNumber, ev.sVal.timestamp
        End If
        If newValue > mSignificantPoint Then mSignificantPoint = newValue
    ElseIf newValue > mPsCache.GetValue(-1) Then
        mCurrentStop = mPsCache.GetValue(-1)
    Else
        mLong = False
        mCurrentStop = mSignificantPoint
        mInitialStopForBar = mCurrentStop
        mPsCache.UpdateValue mCurrentStop, ev.sVal.barNumber, ev.sVal.timestamp
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
            mPsCache.UpdateValue mCurrentStop, ev.sVal.barNumber, ev.sVal.timestamp
        End If
        If newValue < mSignificantPoint Then mSignificantPoint = newValue
    ElseIf newValue < mPsCache.GetValue(-1) Then
        mCurrentStop = mPsCache.GetValue(-1)
    Else
        mLong = True
        mCurrentStop = mSignificantPoint
        mInitialStopForBar = mCurrentStop
        mPsCache.UpdateValue mCurrentStop, ev.sVal.barNumber, ev.sVal.timestamp
        mAccelerationFactor = 0
        mExtremePoint = newValue
        mSignificantPoint = newValue
    End If
End If

Set evOut.Source = Me
evOut.sVal.barNumber = ev.sVal.barNumber
evOut.sVal.barStartTime = ev.sVal.barStartTime
evOut.sVal.timestamp = ev.sVal.timestamp
evOut.sVal.value = mCurrentStop
evOut.valueName = PsValuePs
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = 100 ' difficult to say!

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
    Case UCase$(PsParamStartFactor)
        mStartFactor = param.value
    Case UCase$(PsParamIncrement)
        mIncrement = param.value
    Case UCase$(PsParamMaxFactor)
        mMaxFactor = param.value
    End Select
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GParabolicStop.StudyDefinition

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

Public Property Get ps(Optional ByVal ref As Integer) As SValue
Const ProcName As String = "ps"
On Error GoTo Err

ps = mPsCache.GetSValue(ref)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "Study_getStudyValue"
On Error GoTo Err

Study_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)

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

Study_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)

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




