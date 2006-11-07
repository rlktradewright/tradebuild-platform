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

Implements TradeBuildSP.IStudy

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
Private mPsCache As IStudyValueCache
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
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)
customProcessParameters parameters
Set mPsCache = mStudyFoundation.FindValueCache(PsValuePs)
End Sub

Private Sub customNotify( _
                ev As TradeBuildSP.StudyValueEvent)
Dim newValue As Double
Dim evOut As StudyValueEvent
Static accelAdjustedThisbar As Boolean
Static sPrevVal As Variant

If IsEmpty(ev.value) Then Exit Sub

If Not mStudyFoundation.isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
    sPrevVal = ev.value
End If

newValue = CDbl(ev.value)

If mCurrentStop = MaxDouble Then
    mCurrentStop = newValue * 0.999
    mExtremePoint = newValue * 1.01
    mInitialStopForBar = mCurrentStop
End If

If mStudyFoundation.isNewBar Then
    accelAdjustedThisbar = False
    mInitialStopForBar = mCurrentStop
    mPsCache.addValue mCurrentStop, ev.barnumber, ev.timestamp
    If mAccelerationFactor = 0 Then
        mAccelerationFactor = mStartFactor
    End If
    If mLong Then
        mCurrentStop = mInitialStopForBar + mAccelerationFactor * (mExtremePoint - mInitialStopForBar)
    Else
        mCurrentStop = mInitialStopForBar - mAccelerationFactor * (mInitialStopForBar - mExtremePoint)
    End If
    mPsCache.updateValue mCurrentStop, ev.barnumber, ev.timestamp
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
            mPsCache.updateValue mCurrentStop, ev.barnumber, ev.timestamp
        End If
        If newValue > mSignificantPoint Then mSignificantPoint = newValue
    ElseIf newValue > mPsCache.getValue(-1) Then
        mCurrentStop = mPsCache.getValue(-1)
    Else
        mLong = False
        mCurrentStop = mSignificantPoint
        mInitialStopForBar = mCurrentStop
        mPsCache.updateValue mCurrentStop, ev.barnumber, ev.timestamp
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
            mPsCache.updateValue mCurrentStop, ev.barnumber, ev.timestamp
        End If
        If newValue < mSignificantPoint Then mSignificantPoint = newValue
    ElseIf newValue < mPsCache.getValue(-1) Then
        mCurrentStop = mPsCache.getValue(-1)
    Else
        mLong = True
        mCurrentStop = mSignificantPoint
        mInitialStopForBar = mCurrentStop
        mPsCache.updateValue mCurrentStop, ev.barnumber, ev.timestamp
        mAccelerationFactor = 0
        mExtremePoint = newValue
        mSignificantPoint = newValue
    End If
End If

Set evOut.Source = Me
evOut.barnumber = ev.barnumber
evOut.timestamp = ev.timestamp
evOut.value = mCurrentStop
evOut.valueName = PsValuePs
mStudyFoundation.notifyValue evOut

End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = 100 ' difficult to say!
End Function

Private Sub customProcessParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter

Set param = params.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(PsParamStartFactor)
        mStartFactor = param.value
    Case UCase$(PsParamIncrement)
        mIncrement = param.value
    Case UCase$(PsParamMaxFactor)
        mMaxFactor = param.value
    End Select
    Set param = params.getNextParameter
Loop

End Sub

Private Function customStudyDefinition() As IStudyDefinition
Set customStudyDefinition = GParabolicStop.studyDefinition
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
' IStudy Interface Members
'================================================================================

Private Function IStudy_addStudy( _
                ByVal study As IStudy, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion

Set IStudy_addStudy = mStudyFoundation.addStudy( _
                            study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)
End Function

Private Function IStudy_addStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion
Set IStudy_addStudyValueListener = mStudyFoundation.addStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)
End Function

Private Property Get IStudy_baseStudy() As IStudy
Set IStudy_baseStudy = Me
End Property

Private Function IStudy_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As StudyValue
IStudy_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)
End Function

Private Property Get IStudy_id() As String
IStudy_id = mStudyFoundation.id
End Property

Private Sub IStudy_initialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)
Set mStudyFoundation = New StudyFoundation
mStudyFoundation.initialise Me, _
                            id, _
                            parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudy, _
                            commonServiceConsumer, _
                            studyServiceConsumer
                            

customInitialise commonServiceConsumer, _
                studyServiceConsumer, _
                id, _
                parameters, _
                numberOfValuesToCache, _
                inputValueNames, _
                underlyingStudy
End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = mStudyFoundation.instanceName
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = mStudyFoundation.instancePath
End Property

Private Sub IStudy_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
mStudyFoundation.NotifyInput ev
customNotify ev
End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long
IStudy_numberOfBarsRequired = mStudyFoundation.numberOfBarsRequired(customNumberOfBarsRequired)
End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal valueName As String = "") As Long
IStudy_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)
End Function

Private Property Get IStudy_parameters() As IParameters
Set IStudy_parameters = mStudyFoundation.parameters
End Property

Private Sub IStudy_removeStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
mStudyFoundation.removeStudyValueListener listener
End Sub

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = customStudyDefinition
End Property




