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

Private mSD As Variant

Private mSum As Double
Private mSumSquares As Double

Private mSDCache As TradeBuildSP.IStudyValueCache
Private mInputValueCache As TradeBuildSP.IStudyValueCache

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Debug.Print "StandardDeviation study terminated"
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
Set mSDCache = mStudyFoundation.FindValueCache(SDValueStandardDeviation)
End Sub

Private Sub customNotify( _
                ev As TradeBuildSP.StudyValueEvent)
Static sPrevVal As Variant
Dim newValue As Double
Dim valueToRemove As Double

If Not mStudyFoundation.isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
End If

newValue = CDbl(ev.value)

If mStudyFoundation.isNewBar Then
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

If mStudyFoundation.isNewBar Then
    mSDCache.addValue mSD, ev.barnumber, ev.timestamp
Else
    mSDCache.updateValue mSD, ev.barnumber, ev.timestamp
End If
            
notifyValue mSD, SDValueStandardDeviation, ev.barnumber, ev.timestamp
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = mPeriods + mSDCache.cyclicSize
End Function

Private Sub customProcessParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter

Set param = params.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(SDParamPeriods)
        mPeriods = param.value
        Set mInputValueCache = mCommonServiceConsumer.NewStudyValueCache(mPeriods, "")
    End Select
    Set param = params.getNextParameter
Loop

End Sub

Private Function customStudyDefinition() As IStudyDefinition
Set customStudyDefinition = GStandardDeviation.studyDefinition
End Function

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




