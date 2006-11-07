VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
' This study class calculates the Relative Strength Index.
'
' There are various ways to calcuate RSI. The one that we use is:
'
'
'               100
' RSI = 100 - ------
'             1 + RS
'
'
'            movAvg(daily gain, n)
' where RS = ---------------------
'            movAvg(daily loss, n)
'
'
' n is the number of periods
'
' We will allow the type of moving average used to be specified as a
' parameter (see studyDefinition property in the GRSI module


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

'-------these variables may be required in some studies---------
Private mStudyServiceConsumer As IStudyServiceConsumer
Private mCommonServiceConsumer As ICommonServiceConsumer
'--------------------------------------------------------

Private mRsiCache As IStudyValueCache

Private mPeriods As Long
Private mMovingAverageType As String

' the following variables refer to moving average study objects which we'll
' get TradeBuild to create for us from the BuiltInStudies service provider.
' Since we don't know the default types of these objects (because that will
' vary depending on whether they are simple, exponential or whatever), we
' will access them only through their IStudy interfaces
Private mMaGains As IStudy
Private mMaLosses As IStudy

' it's easy to calculate the open and close for each bar ourselves, but we'll
' illustrate the use of the Bars class provided by StudiesSupport which
' will do this for us.
Private mBars As StudiesSupport.Bars

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Debug.Print "RSI study terminated"
End Sub

'================================================================================
' Custom procedures that must be implemented
'================================================================================

Private Sub customInitialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)

Set mStudyServiceConsumer = studyServiceConsumer
Set mCommonServiceConsumer = commonServiceConsumer

customProcessParameters parameters

Set mRsiCache = mStudyFoundation.FindValueCache(RsiValueRsi)

' create the Bars object
Set mBars = New StudiesSupport.Bars
' initalise the Bars object. We only need it to cache the latest bar, so
' set the numberOfValuesToCache argument to 1
mBars.initialise mCommonServiceConsumer, 1

End Sub

Private Sub customNotify( _
                ev As TradeBuildSP.StudyValueEvent)
Static sPrevVal As Variant
Dim gainsAvg As TradeBuildSP.StudyValue
Dim lossesAvg As TradeBuildSP.StudyValue
Dim RSI As Variant
Dim evMa As TradeBuildSP.StudyValueEvent

If Not mStudyFoundation.isNewBar Then
    ' if this input value is the same as the previous one in this bar, no
    ' need to do anything as it won't change the study value
    If ev.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.value

' update the bars
mBars.Notify ev

' pass the loss and gain values for the current bar to the relevant moving
' average object

evMa = ev
evMa.value = mBars.gain(0)  ' NB: parameter value of 0 means current bar,
                            ' -1 is previous bar etc
mMaGains.Notify evMa        ' update the gains moving average

evMa.value = mBars.loss(0)
mMaLosses.Notify evMa        ' update the losses moving average

' now get the updated gain and loss moving averages. Note that we don't actually
' need to know the name of the moving average study values, because we can use
' the standard name $DEFAULT (defined in constant DefaultStudyValueName)
' which always refers to a study's default output value.
gainsAvg = mMaGains.getStudyValue(DefaultStudyValueName, 0)
lossesAvg = mMaLosses.getStudyValue(DefaultStudyValueName, 0)

' now we can calculate the RSI. If any of the inputs to this calculation are
' Empty (for example because we haven't yet got far enough to get genuine values
' from the moving averages), then we will store an Empty value in the RSI cache.
' This is so that any user of this study doesn't have to bother to work out whether
' a value for this bar exists or not.
If IsEmpty(gainsAvg.value) Or IsEmpty(lossesAvg.value) Then
    RSI = Empty
Else
    If lossesAvg.value = 0 Then
        RSI = 100
    Else
        RSI = 100 - 100 / (1 + (gainsAvg.value / lossesAvg.value))
    End If
End If

If mStudyFoundation.isNewBar Then
    mRsiCache.addValue RSI, ev.barnumber, ev.timestamp
Else
    mRsiCache.updateValue RSI, ev.barnumber, ev.timestamp
End If

notifyValue RSI, RsiValueRsi, ev.barnumber, ev.timestamp
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = mPeriods
End Function

Private Sub customProcessParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter

' NB: the For Each construct cannot be used with objects of the
' IParameters class

Set param = params.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(RsiParamPeriods)
        mPeriods = param.value
    Case UCase$(RsiParamMovingAverageType)
        mMovingAverageType = param.value
    End Select
    Set param = params.getNextParameter
Loop

' Now we know what we need to create the two moving average study objects. Note
' that we only need these studies to cache the latest moving average value.
Set mMaGains = gCreateMA(mMovingAverageType, _
                        mCommonServiceConsumer, _
                        mStudyServiceConsumer, _
                        mPeriods, _
                        1)
Set mMaLosses = gCreateMA(mMovingAverageType, _
                        mCommonServiceConsumer, _
                        mStudyServiceConsumer, _
                        mPeriods, _
                        1)

End Sub

Private Function customStudyDefinition() As IStudyDefinition
Set customStudyDefinition = GRSI.studyDefinition
End Function

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

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
' From this point on the code is identical for all studies - do not change unless
' you are sure you know what you're doing!!!!
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












