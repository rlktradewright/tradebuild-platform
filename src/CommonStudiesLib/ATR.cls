VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ATR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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

Private mMAStudy As Study


Private mTR As Double
Private mATR As Variant
Private mATRCache As StudyValueCache

Private mPrevClose As Variant   ' this is Variant type to cater for cases where
                                ' a bar is empty (eg the first bar has no previous bar)
Private mCurrHigh As Double
Private mCurrLow As Double
Private mCurrClose As Double

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Debug.Print "ATR study terminated"
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

customProcessParameters Parameters
Set mATRCache = mStudyFoundation.FindValueCache(AtrValueATR)
End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Dim newval As Double
Dim evMa As StudyValueEvent
Dim evOut As StudyValueEvent
Static sPrevVal As Variant

If Not mStudyFoundation.isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.value

newval = CDbl(ev.value)

evMa = ev

If mStudyFoundation.isNewBar Then
    mPrevClose = mCurrClose
    mCurrHigh = newval
    mCurrLow = newval
    mCurrClose = newval
    
    If mStudyFoundation.localBarNumber <> 1 Then
        mTR = trueRange(newval)
        
        evMa.value = mTR
        mMAStudy.Notify evMa
        mATR = mMAStudy.getStudyValue(MovingAverageStudyValueName, 0).value
    End If
    mATRCache.addValue mATR, ev.barNumber, ev.timestamp, ev.barStartTime
Else
    If newval > mCurrHigh Then
        mCurrHigh = newval
    End If
    If newval < mCurrLow Then
        mCurrLow = newval
    End If
    mCurrClose = newval
    
    If mStudyFoundation.localBarNumber <> 1 Then
        mTR = trueRange(newval)
        
        evMa.value = mTR
        mMAStudy.Notify evMa
        mATR = mMAStudy.getStudyValue(MovingAverageStudyValueName, 0).value
        mATRCache.updateValue mATR, ev.barNumber, ev.timestamp
    End If
End If

Set evOut.Source = Me
evOut.barNumber = ev.barNumber
evOut.barStartTime = ev.barStartTime
evOut.timestamp = ev.timestamp
evOut.value = mATR
evOut.valueName = AtrValueATR
mStudyFoundation.notifyValue evOut
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = mPeriods + 1 + mATRCache.cyclicSize
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter
Dim maType As String

Set param = params.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(AtrParamPeriods)
        mPeriods = param.value
    Case UCase$(AtrParamMAType)
        maType = param.value
    End Select
    Set param = params.getNextParameter
Loop

Set mMAStudy = gCreateMA(maType, _
                    mPeriods, _
                    1)
End Sub

Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GATR.StudyDefinition
End Function

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get ATR(Optional ByVal ref As Long) As StudyValue
ATR = mATRCache.getStudyValue(ref)
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Function trueRange(ByVal value As Variant) As Variant
Dim thisBarRange As Double
Dim prevCloseToHigh As Double
Dim prevCloseToLow As Double

If IsEmpty(value) Or IsEmpty(mPrevClose) Then Exit Function

thisBarRange = mCurrHigh - mCurrLow
prevCloseToHigh = mCurrHigh - mPrevClose
prevCloseToLow = mPrevClose - mCurrLow

trueRange = IIf(thisBarRange >= prevCloseToHigh, thisBarRange, prevCloseToHigh)
trueRange = IIf(trueRange >= prevCloseToLow, trueRange, prevCloseToLow)
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
                Optional ByVal taskData As Variant) As Tasks.TaskCompletion

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
                Optional ByVal taskData As Variant) As Tasks.TaskCompletion
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
mStudyFoundation.NotifyInput ev
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


