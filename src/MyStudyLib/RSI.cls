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
' There are various ways to calculate RSI. The one that we use is:
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

Implements study

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "RSI"

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


Private mSession As session
Private mRsiCache As ValueCache

Private mPeriods As Long
Private mMovingAverageType As String

' the following variables refer to moving average study objects which we'll
' get TradeBuild to create for us from the BuiltInStudies service provider.
' Since we don't know the default types of these objects (because that will
' vary depending on whether they are simple, exponential or whatever), we
' will access them only through their Study interfaces
Private mMaGains As study
Private mMaLosses As study

' it's easy to calculate the open and close for each bar ourselves, but we'll
' illustrate the use of the Bars class provided by Timeframe Utilities which
' will do this for us. Note that you actually construct a Bars object using
' a BarsBuilder object.
Private mBarsBuilder As BarsBuilder
Private mBars As Bars

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
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As study, _
                ByVal pSession As session)

Set mSession = pSession

customProcessParameters Parameters

Set mRsiCache = mStudyFoundation.FindValueCache(RsiValueRsi)

' create the Bars object. We only need it to cache the latest bar, so
' set the numberOfValuesToCache argument to 1
Set mBarsBuilder = CreateBarsBuilder(GetTimePeriod(0, TimePeriodNone), _
                                mSession, _
                                0, _
                                1)
Set mBars = mBarsBuilder.Bars
End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Static sPrevVal As Variant
Dim gainsAvg As SValue
Dim lossesAvg As SValue
Dim rsi As Variant
Dim evMa As StudyValueEvent

Dim failpoint As Long
On Error GoTo Err

If Not mStudyFoundation.isNewBar Then
    ' if this input value is the same as the previous one in this bar, no
    ' need to do anything as it won't change the study value
    If ev.sVal.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.value

' update the bars
mBarsBuilder.notifyValue ev.sVal

' pass the loss and gain values for the current bar to the relevant moving
' average object

evMa = ev
evMa.sVal.value = mBars.gain(0)  ' NB: parameter value of 0 means current bar,
                            ' -1 is previous bar etc
mMaGains.notify evMa        ' update the gains moving average

evMa.sVal.value = mBars.loss(0)
mMaLosses.notify evMa        ' update the losses moving average

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
    rsi = Empty
Else
    If lossesAvg.value = 0 Then
        rsi = 100
    Else
        rsi = 100 - 100 / (1 + (gainsAvg.value / lossesAvg.value))
    End If
End If

If mStudyFoundation.isNewBar Then
    mRsiCache.addValue rsi, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
Else
    mRsiCache.updateValue rsi, ev.sVal.barNumber, ev.sVal.timestamp
End If

notifyValue rsi, RsiValueRsi, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "customNotify" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = mPeriods
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

' NB: the For Each construct cannot be used with objects of the
' Parameters class

Dim failpoint As Long
On Error GoTo Err

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(RsiParamPeriods)
        mPeriods = param.value
    Case UCase$(RsiParamMovingAverageType)
        mMovingAverageType = param.value
    End Select
Next

' Now we know what we need to create the two moving average study objects. Note
' that we only need these studies to cache the latest moving average value.
Set mMaGains = gCreateMA(mMovingAverageType, _
                        mPeriods, _
                        1)
Set mMaLosses = gCreateMA(mMovingAverageType, _
                        mPeriods, _
                        1)

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "customProcessParameters" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Private Function customStudyDefinition() As studyDefinition
Set customStudyDefinition = GRsi.studyDefinition
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
                ByVal barNumber As Long, _
                ByVal barStartTime As Date, _
                ByVal timestamp As Date)
Dim ev As StudyValueEvent
If Not IsEmpty(value) Then
    ev.sVal.value = value
    ev.valueName = valueName
    Set ev.Source = Me
    ev.sVal.barNumber = barNumber
    ev.sVal.barStartTime = barStartTime
    ev.sVal.timestamp = timestamp
    mStudyFoundation.notifyValue ev
End If
End Sub


'================================================================================
' From this point on the code is identical for all studies - do not change unless
' you are sure you know what you're doing!!!!
'================================================================================


'================================================================================
' Study Interface Members
'================================================================================

Private Function Study_addStudy( _
                ByVal study As study, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController

Set Study_addStudy = mStudyFoundation.addStudy( _
                            study, _
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

Private Property Get Study_baseStudy() As study
Set Study_baseStudy = Me
End Property

Private Function Study_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Study_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)
End Function

Private Function study_getValueTicksize( _
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
                ByVal underlyingStudy As study, _
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

Private Sub study_setInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)

End Sub

Private Property Get Study_studyDefinition() As studyDefinition
Set Study_studyDefinition = customStudyDefinition
End Property










