VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StandardDeviation"
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
Implements TradeBuildSP.ITaskCompletionListener

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

Private mStudyServiceConsumer As IStudyServiceConsumer
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mIdent As String

Private mParams As IParameters
Private mStudies As IStudies
Private mListeners As IStudyValueListeners

Private mInputValueName As String
Private mUnderlyingStudy As TradeBuildSP.IStudy

Private mPeriods As Long

Private mSD As Variant

Private mSum As Double
Private mSumSquares As Double

Private mSDCache As TradeBuildSP.IStudyValueCache
Private mInputValueCache As TradeBuildSP.IStudyValueCache

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mParams = GStandardDeviation.defaultParameters
End Sub

Private Sub Class_Terminate()
Debug.Print "StandardDeviation study terminated"
End Sub

'================================================================================
' IStudy Interface Members
'================================================================================

Private Function IStudy_addStudy( _
                ByVal study As IStudy, _
                ByVal ValueName As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion


If ValueName = SDValueStandardDeviation Then
    Set IStudy_addStudy = mSDCache.replayValues(study, Me, numUnderlyingValuesToUse, taskName, taskData)
    IStudy_addStudy.addTaskCompletionListener Me
End If

End Function

Private Sub IStudy_addStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByRef valueNames() As String)
mListeners.Add listener, valueNames
End Sub

Private Property Get IStudy_baseStudy() As IStudy
Set IStudy_baseStudy = Me
End Property

Private Property Let IStudy_defaultParameters(ByVal RHS As TradeBuildSP.IParameters)
GStandardDeviation.defaultParameters = RHS
End Property

Private Property Get IStudy_defaultParameters() As TradeBuildSP.IParameters
Set IStudy_defaultParameters = GStandardDeviation.defaultParameters
End Property

Private Function IStudy_getStudyValue( _
                ByVal ValueName As String, _
                ByVal ref As Long, _
                ByRef params() As Variant) As StudyValue
Select Case UCase$(ValueName)
Case UCase$(DefaultStudyValueName), UCase$(SDValueStandardDeviation)
    IStudy_getStudyValue = SD(ref)
End Select

End Function

Private Property Get IStudy_id() As String
IStudy_id = mIdent
End Property

Private Sub IStudy_initialise( _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByVal inputValueName As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)
mIdent = id
mInputValueName = inputValueName
Set mUnderlyingStudy = underlyingStudy
processParameters parameters
Set mSDCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, SDValueStandardDeviation)

End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = SdShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams) & ")"
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = SdShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams) & ")." & _
                mUnderlyingStudy.instancePath
End Property

Private Sub IStudy_Notify(ev As TradeBuildSP.StudyValueEvent)
Dim isNewBar  As Boolean
Static sPrevVal As Variant
Dim newValue As Double
Dim valueToRemove As Double

isNewBar = newBar(ev.barnumber)

If Not isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
End If

newValue = CDbl(ev.value)

If isNewBar Then
    If mInputValueCache.Count < mPeriods Then
        If Not IsEmpty(ev.value) Then
            mInputValueCache.addValue newValue, ev.barnumber, ev.timestamp
            mSum = mSum + newValue
            mSumSquares = mSumSquares + newValue * newValue
        End If
        If mInputValueCache.Count = mPeriods Then
            mSD = Sqr((mPeriods * mSumSquares - mSum * mSum) / (mPeriods * mPeriods))
        End If
    Else
        If Not IsEmpty(ev.value) Then
            valueToRemove = mInputValueCache.getValue(1)
            mInputValueCache.addValue newValue, ev.barnumber, ev.timestamp
            mSum = mSum - valueToRemove + newValue
            mSumSquares = mSumSquares - valueToRemove * valueToRemove + newValue * newValue
            mSD = Sqr((mPeriods * mSumSquares - mSum * mSum) / (mPeriods * mPeriods))
        End If
    End If
    sPrevVal = ev.value
Else
    If Not IsEmpty(ev.value) Then
        If Not IsEmpty(sPrevVal) Then
            valueToRemove = sPrevVal
            mInputValueCache.updateValue newValue, ev.barnumber, ev.timestamp
            mSum = mSum - valueToRemove
            mSumSquares = mSumSquares - valueToRemove * valueToRemove
        Else
            ' this can only happen when there has not yet been a non-empty
            ' input value in this bar
            mInputValueCache.addValue newValue, ev.barnumber, ev.timestamp
        End If
        mSum = mSum + newValue
        mSumSquares = mSumSquares + newValue * newValue
        sPrevVal = ev.value
    End If
    If mInputValueCache.Count >= mPeriods Then
        mSD = Sqr((mPeriods * mSumSquares - mSum * mSum) / (mPeriods * mPeriods))
    End If
End If

If isNewBar Then
    mSDCache.addValue mSD, ev.barnumber, ev.timestamp
Else
    mSDCache.updateValue mSD, ev.barnumber, ev.timestamp
End If
            
notifyValue mSD, SDValueStandardDeviation, ev.barnumber, ev.timestamp
End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long
Dim thisNeeds As Long
Dim superiorStudiesNeeds
If mSDCache.cyclicSize = 0 Then
    IStudy_numberOfBarsRequired = 0
Else
    thisNeeds = mPeriods + mSDCache.cyclicSize
    superiorStudiesNeeds = mStudies.numberOfBarsRequired
    IStudy_numberOfBarsRequired = IIf(thisNeeds > superiorStudiesNeeds, _
                                    thisNeeds, _
                                    superiorStudiesNeeds)
End If
End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal ValueName As String = "") As Long
IStudy_numberOfCachedValues = mSDCache.Count
End Function

Private Property Get IStudy_parameters() As IParameters
Set IStudy_parameters = mParams.Clone
End Property

Private Sub IStudy_removeStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
mListeners.Remove listener
End Sub

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = GStandardDeviation.studyDefinition
End Property

'================================================================================
' ITaskCompletionListener Interface Members
'================================================================================

Private Sub ITaskCompletionListener_taskCompleted(ev As TradeBuildSP.TaskCompletionEvent)

End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get SD(Optional ByVal ref As Long) As StudyValue
SD = mSDCache.getStudyValue(ref)
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub initialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                Optional ByVal numberOfValuesToCache As Long, _
                Optional ByVal params As IParameters)

Set mCommonServiceConsumer = commonServiceConsumer
Set mStudyServiceConsumer = studyServiceConsumer
Set mStudies = mCommonServiceConsumer.NewStudies
Set mListeners = mCommonServiceConsumer.NewStudyValueListeners

Set mSDCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, SDValueStandardDeviation)

If Not params Is Nothing Then processParameters params

End Sub

'================================================================================
' Helper Functions
'================================================================================



Private Function newBar(barnumber As Long) As Boolean
Static previousBarNumber As Long
If barnumber <> previousBarNumber Then
    previousBarNumber = barnumber
    newBar = True
Else
    newBar = False
End If
End Function

Private Sub notifyValue( _
                ByVal value As Variant, _
                ByVal ValueName As String, _
                ByVal barnumber As Long, _
                ByVal timestamp As Date)
Dim ev As StudyValueEvent
If Not IsEmpty(value) Then
    ev.value = value
    ev.ValueName = ValueName
    Set ev.Source = Me
    ev.barnumber = barnumber
    ev.timestamp = timestamp
    mStudies.Notify ev
    mListeners.Notify ev
End If
End Sub

Private Sub processParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter

Set mParams = params.Clone

Set param = mParams.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(SDParamPeriods)
        mPeriods = param.value
        Set mInputValueCache = mCommonServiceConsumer.NewStudyValueCache(mPeriods, "")
    End Select
    Set param = mParams.getNextParameter
Loop

End Sub

