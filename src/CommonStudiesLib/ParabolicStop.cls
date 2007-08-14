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

Private mInitialStopForBar As Double
Private mCurrentStop As Double
Private mPsCache As StudyValueCache
Private mExtremePoint As Double
Private mSignificantPoint As Double

Private mAccelerationFactor As Double
Private mStartFactor As Double
Private mIncrement As Double
Private mMaxFactor As Double

Private mLong As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mCurrentStop = MaxDouble
mLong = True    ' arbitrary initial setting
End Sub

Private Sub Class_Terminate()
Debug.Print "Parabolic Stop terminated"
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

customProcessParameters Parameters
Set mPsCache = mStudyFoundation.FindValueCache(PsValuePs)
End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Dim newValue As Double
Dim evOut As StudyValueEvent
Static accelAdjustedThisbar As Boolean
Static sPrevVal As Variant

If IsEmpty(ev.value) Then Exit Sub

If Not mStudyFoundation.isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.value

newValue = CDbl(ev.value)

If mCurrentStop = MaxDouble Then
    mCurrentStop = newValue * 0.999
    mExtremePoint = newValue * 1.01
    mInitialStopForBar = mCurrentStop
End If

If mStudyFoundation.isNewBar Then
    accelAdjustedThisbar = False
    mInitialStopForBar = mCurrentStop
    mPsCache.addValue mCurrentStop, ev.barNumber, ev.timestamp, ev.barStartTime
    If mAccelerationFactor = 0 Then
        mAccelerationFactor = mStartFactor
    End If
    If mLong Then
        mCurrentStop = mInitialStopForBar + mAccelerationFactor * (mExtremePoint - mInitialStopForBar)
    Else
        mCurrentStop = mInitialStopForBar - mAccelerationFactor * (mInitialStopForBar - mExtremePoint)
    End If
    mPsCache.updateValue mCurrentStop, ev.barNumber, ev.timestamp
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
            mPsCache.updateValue mCurrentStop, ev.barNumber, ev.timestamp
        End If
        If newValue > mSignificantPoint Then mSignificantPoint = newValue
    ElseIf newValue > mPsCache.getValue(-1) Then
        mCurrentStop = mPsCache.getValue(-1)
    Else
        mLong = False
        mCurrentStop = mSignificantPoint
        mInitialStopForBar = mCurrentStop
        mPsCache.updateValue mCurrentStop, ev.barNumber, ev.timestamp
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
            mPsCache.updateValue mCurrentStop, ev.barNumber, ev.timestamp
        End If
        If newValue < mSignificantPoint Then mSignificantPoint = newValue
    ElseIf newValue < mPsCache.getValue(-1) Then
        mCurrentStop = mPsCache.getValue(-1)
    Else
        mLong = True
        mCurrentStop = mSignificantPoint
        mInitialStopForBar = mCurrentStop
        mPsCache.updateValue mCurrentStop, ev.barNumber, ev.timestamp
        mAccelerationFactor = 0
        mExtremePoint = newValue
        mSignificantPoint = newValue
    End If
End If

Set evOut.Source = Me
evOut.barNumber = ev.barNumber
evOut.barStartTime = ev.barStartTime
evOut.timestamp = ev.timestamp
evOut.value = mCurrentStop
evOut.valueName = PsValuePs
mStudyFoundation.notifyValue evOut

End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = 100 ' difficult to say!
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

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

End Sub

Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GParabolicStop.StudyDefinition
End Function

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get ps(Optional ByVal ref As Integer) As StudyValue
ps = mPsCache.getStudyValue(ref)
End Property

Public Sub setStop(ByVal stopValue As Double, _
                    ByVal isLong As Boolean, _
                    ByVal extremePoint As Double)
mLong = isLong
mCurrentStop = stopValue
mInitialStopForBar = mCurrentStop
mPsCache.updateValue mCurrentStop
mAccelerationFactor = 0
mExtremePoint = extremePoint
mSignificantPoint = extremePoint
End Sub

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


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
Study_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)
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
Study_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)
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




