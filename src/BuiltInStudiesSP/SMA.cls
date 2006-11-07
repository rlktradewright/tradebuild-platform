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

Implements TradeBuildSP.IStudy
Implements IMovingAverageStudy

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

Private mStudyServiceConsumer As IStudyServiceConsumer
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mPeriods As Long

Private mSlopeThreshold As Double

Private mInputValueCache As TradeBuildSP.IStudyValueCache

Private mMaCache As TradeBuildSP.IStudyValueCache

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
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)
Set mCommonServiceConsumer = commonServiceConsumer
Set mStudyServiceConsumer = studyServiceConsumer
customProcessParameters parameters
Set mMaCache = mStudyFoundation.FindValueCache(SMAValueSMA)
End Sub

Private Sub customNotify( _
                ev As TradeBuildSP.StudyValueEvent)
Static sPrevVal As Variant
Dim evOut As StudyValueEvent

If Not mStudyFoundation.isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
End If

If mStudyFoundation.isNewBar Then
    If mInputValueCache.Count < mPeriods Then
        If Not IsEmpty(ev.value) Then
            mSum = mSum + ev.value
            mInputValueCache.addValue ev.value, ev.barnumber, ev.timestamp
        End If
        If mInputValueCache.Count = mPeriods Then
            mMA = mSum / mPeriods
        End If
    Else
        If Not IsEmpty(ev.value) Then
            mSum = mSum - mInputValueCache.getValue(1) + ev.value
            mMA = mSum / mPeriods
            mInputValueCache.addValue ev.value, ev.barnumber, ev.timestamp
        End If
    End If
    sPrevVal = ev.value
Else
    If Not IsEmpty(ev.value) Then
        If Not IsEmpty(sPrevVal) Then
            mSum = mSum - sPrevVal
            mInputValueCache.updateValue ev.value, ev.barnumber, ev.timestamp
        Else
            ' this can only happen when there has not yet been a non-empty
            ' input value in this bar
            mInputValueCache.addValue ev.value, ev.barnumber, ev.timestamp
        End If
        mSum = mSum + ev.value
        sPrevVal = ev.value
    End If
    If mInputValueCache.Count >= mPeriods Then
        mMA = mSum / mPeriods
    End If
End If

If mStudyFoundation.isNewBar Then
    mMaCache.addValue mMA, ev.barnumber, ev.timestamp
Else
    mMaCache.updateValue mMA, ev.barnumber, ev.timestamp
End If
            
Set evOut.Source = Me
evOut.barnumber = ev.barnumber
evOut.timestamp = ev.timestamp
evOut.value = mMA
evOut.valueName = SMAValueSMA
mStudyFoundation.notifyValue evOut
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = mPeriods + mMaCache.cyclicSize
End Function

Private Sub customProcessParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter

Set param = params.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(SMAParamPeriods)
        mPeriods = param.value
        Set mInputValueCache = mCommonServiceConsumer.NewStudyValueCache(mPeriods, "")
    Case UCase$(SMAParamSlopeThreshold)
        mSlopeThreshold = param.value
    End Select
    Set param = params.getNextParameter
Loop

End Sub

Private Function customStudyDefinition() As IStudyDefinition
Set customStudyDefinition = GSMA.studyDefinition
End Function

'================================================================================
' IMovingAverageStudy Interface Members
'================================================================================

Public Property Get IMovingAverageStudy_ma( _
                Optional ByVal ref As Long) As StudyValue
IMovingAverageStudy_ma = ma(ref)
End Property

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







