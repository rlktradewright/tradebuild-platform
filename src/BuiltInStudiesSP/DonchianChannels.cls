VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DonchianChannels"
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

Private mHighCache As IStudyValueCache
Private mLowcache As IStudyValueCache
Private mCurrHigh As Double
Private mCurrLow As Double

Private mUpper As Double
Private mLower As Double

Private mUpperCache As IStudyValueCache
Private mLowerCache As IStudyValueCache

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Debug.Print "Donchian Channels study terminated"
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
Set mLowerCache = mStudyFoundation.FindValueCache(DoncValueLower)
Set mUpperCache = mStudyFoundation.FindValueCache(DoncValueUpper)
End Sub

Private Sub customNotify( _
                ev As TradeBuildSP.StudyValueEvent)
Static sPrevVal As Variant
Static numValues As Long
Dim i As Long

If IsEmpty(ev.value) Then Exit Sub

If Not mStudyFoundation.isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.value

If mStudyFoundation.isNewBar Then
    numValues = numValues + 1
    mCurrHigh = ev.value
    mCurrLow = ev.value
    mHighCache.addValue mCurrHigh, ev.barnumber, ev.timestamp
    mLowcache.addValue mCurrLow, ev.barnumber, ev.timestamp
    If numValues >= mPeriods Then
        mUpper = DummyHigh
        For i = 1 To mPeriods - 1
            If mHighCache.getValue(i) > mUpper Then
                mUpper = mHighCache.getValue(i)
            End If
        Next
        mUpperCache.addValue mUpper, ev.barnumber, ev.timestamp
        notifyValue mUpper, DoncValueUpper, ev.barnumber, ev.timestamp
        
        mLower = DummyLow
        For i = 1 To mPeriods - 1
            If mLowcache.getValue(i) < mLower Then
                mLower = mLowcache.getValue(i)
            End If
        Next
        mLowerCache.addValue mLower, ev.barnumber, ev.timestamp
        notifyValue mLower, DoncValueLower, ev.barnumber, ev.timestamp
    Else
        mUpperCache.addValue Empty, ev.barnumber, ev.timestamp
        notifyValue Empty, DoncValueUpper, ev.barnumber, ev.timestamp
        mLowerCache.addValue Empty, ev.barnumber, ev.timestamp
        notifyValue Empty, DoncValueLower, ev.barnumber, ev.timestamp
    End If
End If

If ev.value > mCurrHigh Then
    mCurrHigh = ev.value
    mHighCache.updateValue mCurrHigh, ev.barnumber, ev.timestamp
End If
If ev.value < mCurrLow Then
    mCurrLow = ev.value
    mLowcache.updateValue mCurrLow, ev.barnumber, ev.timestamp
End If
    
If numValues < mPeriods Then Exit Sub

If ev.value > mUpper Then
    mUpper = ev.value
    mUpperCache.updateValue mUpper, ev.barnumber, ev.timestamp
    notifyValue mUpper, DoncValueUpper, ev.barnumber, ev.timestamp
End If
If ev.value < mLower Then
    mLower = ev.value
    mLowerCache.updateValue mLower, ev.barnumber, ev.timestamp
    notifyValue mLower, DoncValueLower, ev.barnumber, ev.timestamp
End If
        
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = mPeriods + mHighCache.cyclicSize
End Function

Private Sub customProcessParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter

Set param = params.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(DoncParamPeriods)
        mPeriods = param.value
        Set mHighCache = mCommonServiceConsumer.NewStudyValueCache(mPeriods, "highs")
        Set mLowcache = mCommonServiceConsumer.NewStudyValueCache(mPeriods, "lows")
    End Select
    Set param = params.getNextParameter
Loop

End Sub

Private Function customStudyDefinition() As IStudyDefinition
Set customStudyDefinition = GDonchianChannels.studyDefinition
End Function

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get lower(Optional ByVal ref As Long) As StudyValue
lower = mLowerCache.getStudyValue(ref)
End Property

Public Property Get upper(Optional ByVal ref As Long) As StudyValue
upper = mUpperCache.getStudyValue(ref)
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub notifyValue( _
                ByVal value As Variant, _
                ByVal valueName As String, _
                ByVal barnumber As Long, _
                ByVal timestamp As Date)
Dim ev As StudyValueEvent
If Not IsEmpty(value) Then
    ev.value = value
    ev.valueName = valueName
    Set ev.Source = Me
    ev.barnumber = barnumber
    ev.timestamp = timestamp
    mStudyFoundation.notifyValue ev
End If
End Sub


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








