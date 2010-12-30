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

'@================================================================================
' Description
'@================================================================================
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


'@================================================================================
' Interfaces
'@================================================================================

Implements Study

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "RSI"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

'-------this variable is required in all studies---------
Private mStudyFoundation As StudyFoundation
'--------------------------------------------------------


Private mSession As Session
Private mRsiCache As ValueCache

Private mPeriods As Long
Private mMovingAverageType As String

' the following variables refer to moving average study objects which we'll
' get TradeBuild to create for us from the BuiltInStudies service provider.
' Since we don't know the default types of these objects (because that will
' vary depending on whether they are simple, exponential or whatever), we
' will access them only through their IStudy interfaces
Private mMaGains As Study
Private mMaLosses As Study

' it's easy to calculate the open and close for each bar ourselves, but we'll
' illustrate the use of the Bars class provided by Timeframe Utilities which
' will do this for us. Note that you actually construct a Bars object using
' a BarsBuilder object.
Private mBarsBuilder As BarsBuilder
Private mBars As Bars

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "RSI study terminated"
End Sub

'@================================================================================
' Custom procedures that must be implemented
'@================================================================================

Private Sub customInitialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)

Const ProcName As String = "customInitialise"
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub customNotify( _
                ev As StudyValueEventData)
Static sPrevVal As Variant
Dim gainsAvg As SValue
Dim lossesAvg As SValue
Dim RSI As Variant
Dim evMa As StudyValueEventData

Const ProcName As String = "customNotify"
On Error GoTo Err

If Not mStudyFoundation.IsNewBar Then
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
evMa.sVal.value = mBars.Gain(0)  ' NB: parameter value of 0 means current bar,
                            ' -1 is previous bar etc
mMaGains.Notify evMa        ' update the gains moving average

evMa.sVal.value = mBars.Loss(0)
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

If mStudyFoundation.IsNewBar Then
    mRsiCache.AddValue RSI, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
Else
    mRsiCache.UpdateValue RSI, ev.sVal.barNumber, ev.sVal.timestamp
End If

notifyValue RSI, RsiValueRsi, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = mPeriods

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

' NB: the For Each construct cannot be used with objects of the
' IParameters class

Const ProcName As String = "customProcessParameters"
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GRSI.StudyDefinition

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub notifyValue( _
                ByVal value As Variant, _
                ByVal valueName As String, _
                ByVal barNumber As Long, _
                ByVal barStartTime As Date, _
                ByVal timestamp As Date)
Dim ev As StudyValueEventData
Const ProcName As String = "notifyValue"
On Error GoTo Err

If Not IsEmpty(value) Then
    ev.sVal.value = value
    ev.valueName = valueName
    Set ev.Source = Me
    ev.sVal.barNumber = barNumber
    ev.sVal.barStartTime = barStartTime
    ev.sVal.timestamp = timestamp
    mStudyFoundation.notifyValue ev
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


'@================================================================================
' From this point on the code is identical for all studies - do not change unless
' you are sure you know what you're doing!!!!
'@================================================================================


'@================================================================================
' Study Interface Members
'@================================================================================

Private Function Study_addStudy( _
                ByVal Study As Study, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController

Const ProcName As String = "Study_addStudy"
On Error GoTo Err

Set Study_addStudy = mStudyFoundation.AddStudy( _
                            Study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function Study_addStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "Study_addStudyValueListener"
On Error GoTo Err

Set Study_addStudyValueListener = mStudyFoundation.AddStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Property Get Study_baseStudy() As Study
Const ProcName As String = "Study_baseStudy"
On Error GoTo Err

Set Study_baseStudy = Me

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Function Study_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Const ProcName As String = "Study_getStudyValue"
On Error GoTo Err

Study_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function Study_getValueTicksize( _
                ByVal valueName As String) As Double

Const ProcName As String = "Study_getValueTicksize"
On Error GoTo Err



Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Property Get Study_id() As String
Const ProcName As String = "Study_id"
On Error GoTo Err

Study_id = mStudyFoundation.id

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub Study_initialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)

Const ProcName As String = "Study_initialise"
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get Study_instanceName() As String
Const ProcName As String = "Study_instanceName"
On Error GoTo Err

Study_instanceName = mStudyFoundation.InstanceName

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get Study_instancePath() As String
Const ProcName As String = "Study_instancePath"
On Error GoTo Err

Study_instancePath = mStudyFoundation.InstancePath

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub Study_Notify( _
                ev As StudyValueEventData)
Const ProcName As String = "Study_Notify"
On Error GoTo Err

mStudyFoundation.NotifyInput ev
customNotify ev

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get Study_numberOfBarsRequired() As Long
Const ProcName As String = "Study_numberOfBarsRequired"
On Error GoTo Err

Study_numberOfBarsRequired = mStudyFoundation.NumberOfBarsRequired(customNumberOfBarsRequired)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Function Study_numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "Study_numberOfCachedValues"
On Error GoTo Err

Study_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Property Get Study_parameters() As Parameters
Const ProcName As String = "Study_parameters"
On Error GoTo Err

Set Study_parameters = mStudyFoundation.Parameters

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub Study_removeStudyValueListener( _
                ByVal listener As StudyValueListener)
Const ProcName As String = "Study_removeStudyValueListener"
On Error GoTo Err

mStudyFoundation.RemoveStudyValueListener listener

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub Study_setInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)

Const ProcName As String = "Study_setInputTicksize"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Const ProcName As String = "Study_studyDefinition"
On Error GoTo Err

Set Study_studyDefinition = customStudyDefinition

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property












