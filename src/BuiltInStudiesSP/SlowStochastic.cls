VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlowStochastic"
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

Implements TradeBuildSP.IStudyValueListener
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

Private mKPeriods As Long
Private mKDPeriods As Long
Private mDPeriods As Long

Private mK As Variant
Private mD As Variant

Private mKCache As IStudyValueCache
Private mDCache As IStudyValueCache

Private mStoch As Stochastic
Private mStochStudy As TradeBuildSP.IStudy
Private mMASmoothing As SMA
Private mMASmoothingStudy As TradeBuildSP.IStudy

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Debug.Print "SlowStochastic terminated"
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
Set mKCache = mStudyFoundation.FindValueCache(SStochValueK)
Set mDCache = mStudyFoundation.FindValueCache(SStochValueD)
End Sub

Private Sub customNotify( _
                ev As TradeBuildSP.StudyValueEvent)
' will never get called because mStoch will be called instead
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = mKPeriods + mDPeriods + mKCache.cyclicSize
End Function

Private Sub customProcessParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter
Dim lparams As TradeBuildSP.IParameters
Dim valueNames(0) As String

Set param = params.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(SStochParamKPeriods)
        mKPeriods = param.value
    Case UCase$(SStochParamKDPeriods)
        mKDPeriods = param.value
    Case UCase$(StochParamDPeriods)
        mDPeriods = param.value
        Set mMASmoothing = gCreateMA("SMA", _
                                    mCommonServiceConsumer, _
                                    mStudyServiceConsumer, _
                                    mDPeriods, _
                                    1)
        Set mMASmoothingStudy = mMASmoothing
    End Select
    Set param = params.getNextParameter
Loop

Set mStoch = New Stochastic
Set mStochStudy = mStoch
Set lparams = GStochastic.defaultParameters
lparams.setParameterValue StochParamKPeriods, mKPeriods
lparams.setParameterValue StochParamDPeriods, mKDPeriods
valueNames(0) = "Input"
mStochStudy.initialise mCommonServiceConsumer, _
                        mStudyServiceConsumer, _
                        mCommonServiceConsumer.GenerateGUIDString, _
                        lparams, _
                        1, _
                        valueNames, _
                        Nothing

mStochStudy.addStudyValueListener Me, StochValueD, 1

End Sub

Private Function customStudyDefinition() As IStudyDefinition
Set customStudyDefinition = GSlowStochastic.studyDefinition
End Function

'================================================================================
' IStudyValueListener Interface Members
'================================================================================

Private Sub IStudyValueListener_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
Static sPrevVal As Variant
Dim newEv As TradeBuildSP.StudyValueEvent

mStudyFoundation.NotifyInput ev

If Not mStudyFoundation.isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
    sPrevVal = ev.value
End If

mK = ev.value
mMASmoothingStudy.Notify ev
mD = mMASmoothing.ma.value

If mStudyFoundation.isNewBar Then
    mKCache.addValue mK, ev.barnumber, ev.timestamp
    mDCache.addValue mD, ev.barnumber, ev.timestamp
Else
    mKCache.updateValue mK, ev.barnumber, ev.timestamp
    mDCache.updateValue mD, ev.barnumber, ev.timestamp
End If

newEv.barnumber = ev.barnumber
newEv.timestamp = ev.timestamp
Set newEv.Source = Me
newEv.value = mK
newEv.valueName = SStochValueK
mStudyFoundation.notifyValue newEv

newEv.value = mD
newEv.valueName = SStochValueD
mStudyFoundation.notifyValue newEv
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get percentD(Optional ByVal ref As Integer) As StudyValue
percentD = mKCache.getStudyValue(ref)
End Property

Public Property Get percentK(Optional ByVal ref As Integer) As StudyValue
percentK = mKCache.getStudyValue(ref)
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
'Set IStudy_baseStudy = Me
Set IStudy_baseStudy = mStoch
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




