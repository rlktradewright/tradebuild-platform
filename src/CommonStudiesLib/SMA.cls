VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SMA"
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


Private mPeriods As Long

Private mSlopeThreshold As Double

Private mInputValueCache As StudyValueCache

Private mMaCache As StudyValueCache

Private mSum As Double
Private mMA As Variant

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Debug.Print "Simple Moving Average Study terminated"
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
Set mMaCache = mStudyFoundation.FindValueCache(MovingAverageStudyValueName)
End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Static sPrevVal As Variant
Dim evOut As StudyValueEvent

If Not mStudyFoundation.isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
End If

If mStudyFoundation.isNewBar Then
    If mInputValueCache.Count < mPeriods Then
        If Not IsEmpty(ev.value) Then
            mSum = mSum + ev.value
            mInputValueCache.addValue ev.value, ev.barNumber, ev.timestamp, ev.barStartTime
        End If
        If mInputValueCache.Count = mPeriods Then
            mMA = mSum / mPeriods
        End If
    Else
        If Not IsEmpty(ev.value) Then
            mSum = mSum - mInputValueCache.getValue(1) + ev.value
            mMA = mSum / mPeriods
            mInputValueCache.addValue ev.value, ev.barNumber, ev.timestamp, ev.barStartTime
        End If
    End If
    sPrevVal = ev.value
Else
    If Not IsEmpty(ev.value) Then
        If Not IsEmpty(sPrevVal) Then
            mSum = mSum - sPrevVal
            mInputValueCache.updateValue ev.value, ev.barNumber, ev.timestamp
        Else
            ' this can only happen when there has not yet been a non-empty
            ' input value in this bar
            mInputValueCache.addValue ev.value, ev.barNumber, ev.timestamp, ev.barStartTime
        End If
        mSum = mSum + ev.value
        sPrevVal = ev.value
    End If
    If mInputValueCache.Count >= mPeriods Then
        mMA = mSum / mPeriods
    End If
End If

If mStudyFoundation.isNewBar Then
    mMaCache.addValue mMA, ev.barNumber, ev.timestamp, ev.barStartTime
Else
    mMaCache.updateValue mMA, ev.barNumber, ev.timestamp
End If
            
Set evOut.Source = Me
evOut.barNumber = ev.barNumber
evOut.barStartTime = ev.barStartTime
evOut.timestamp = ev.timestamp
evOut.value = mMA
evOut.valueName = MovingAverageStudyValueName
mStudyFoundation.notifyValue evOut
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = mPeriods + mMaCache.cyclicSize
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(SMAParamPeriods)
        mPeriods = param.value
        Set mInputValueCache = CreateStudyValueCache(mPeriods, "")
    Case UCase$(SMAParamSlopeThreshold)
        mSlopeThreshold = param.value
    End Select
Next

End Sub

Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GSMA.StudyDefinition
End Function

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get falling( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As StudyValue
Dim sv As StudyValue
sv = slope(ref, periods)
sv.value = (sv.value <= (-1 * mSlopeThreshold))
falling = sv
End Property

Public Property Get flat( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As StudyValue
Dim sv As StudyValue
sv = slope(ref, periods)
sv.value = (sv.value < mSlopeThreshold And _
    sv.value > (-1 * mSlopeThreshold))
flat = sv
End Property

Public Property Get ma(Optional ByVal ref As Long) As StudyValue
ma = mMaCache.getStudyValue(ref)
End Property

Public Property Get rising( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As StudyValue
Dim sv As StudyValue
sv = slope(ref, periods)
sv.value = (sv.value >= mSlopeThreshold)
rising = sv
End Property

Public Property Get slope( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As StudyValue
Dim sv As StudyValue
sv = ma(ref)
sv.value = (sv.value - ma(ref - periods + 1).value) / periods
slope = sv
End Property

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







