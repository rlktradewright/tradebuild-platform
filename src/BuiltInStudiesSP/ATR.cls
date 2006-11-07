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

Private mStudyServiceConsumer As IStudyServiceConsumer
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mPeriods As Long

Private mMA As IMovingAverageStudy
Private mMAStudy As IStudy


Private mTR As Double
Private mATR As Variant
Private mATRCache As TradeBuildSP.IStudyValueCache

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
Set mATRCache = mStudyFoundation.FindValueCache(AtrValueATR)
End Sub

Private Sub customNotify( _
                ev As TradeBuildSP.StudyValueEvent)
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
        mATR = mMA.ma.value
    End If
    mATRCache.addValue mATR, ev.barnumber, ev.timestamp
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
        mATR = mMA.ma.value
        mATRCache.updateValue mATR, ev.barnumber, ev.timestamp
    End If
End If

Set evOut.Source = Me
evOut.barnumber = ev.barnumber
evOut.timestamp = ev.timestamp
evOut.value = mATR
evOut.valueName = AtrValueATR
mStudyFoundation.notifyValue evOut
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = mPeriods + 1 + mATRCache.cyclicSize
End Function

Private Sub customProcessParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter
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

Set mMA = gCreateMA(maType, _
                    mCommonServiceConsumer, _
                    mStudyServiceConsumer, _
                    mPeriods, _
                    1)
Set mMAStudy = mMA

End Sub

Private Function customStudyDefinition() As IStudyDefinition
Set customStudyDefinition = GATR.studyDefinition
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


