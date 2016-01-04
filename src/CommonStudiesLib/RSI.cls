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

Implements IStudy

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

Private mStudyManager As StudyManager

Private mSessionFuture As IFuture
Private mRsiCache As ValueCache

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
' illustrate the use of the Bars class provided by Timeframe Utilities which
' will do this for us. Note that you actually construct a Bars object using
' a BarsBuilder object.
Private mBarsBuilder As BarsBuilder
Private mBars As Bars

Private WithEvents mFutureWaiter As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

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
                ByVal Id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal UnderlyingStudy As IStudy, _
                ByVal pSessionFuture As IFuture)

Const ProcName As String = "customInitialise"
On Error GoTo Err

Set mSessionFuture = pSessionFuture
customProcessParameters Parameters

Set mRsiCache = mStudyFoundation.FindValueCache(RsiValueRsi)

' create the Bars object. We only need it to cache the latest bar, so
' set the numberOfValuesToCache argument to 1
If mSessionFuture.IsAvailable Then
    Set mBarsBuilder = CreateBarsBuilder(GetTimePeriod(0, TimePeriodNone), _
                                    mSessionFuture.Value, _
                                    0, _
                                    1)
    Set mBars = mBarsBuilder.Bars
Else
    Set mFutureWaiter = New FutureWaiter
    mFutureWaiter.Add CreateBarsBuilderFuture(GetTimePeriod(0, TimePeriodNone), _
                                    mSessionFuture, _
                                    0, _
                                    1)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub customNotify( _
                ByRef ev As StudyValueEventData)
Static sPrevVal As Variant
Dim gainsAvg As SValue
Dim lossesAvg As SValue
Dim RSI As Variant
Dim evMa As StudyValueEventData

Const ProcName As String = "customNotify"
On Error GoTo Err

If Not mStudyFoundation.IsNewBar Then
    ' if this input Value is the same as the previous one in this bar, no
    ' need to do anything as it won't change the study Value
    If ev.sVal.Value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.Value

' update the bars
mBarsBuilder.notifyValue ev.sVal

' pass the loss and gain values for the current bar to the relevant moving
' average object

evMa = ev
evMa.sVal.Value = mBars.Gain(0)  ' NB: parameter Value of 0 means current bar,
                            ' -1 is previous bar etc
mMaGains.Notify evMa        ' update the gains moving average

evMa.sVal.Value = mBars.Loss(0)
mMaLosses.Notify evMa        ' update the losses moving average

' now get the updated gain and loss moving averages. Note that we don't actually
' need to know the name of the moving average study values, because we can use
' the standard name $DEFAULT (defined in constant DefaultStudyValueName)
' which always refers to a study's default output Value.
gainsAvg = mMaGains.getStudyValue(DefaultStudyValueName, 0)
lossesAvg = mMaLosses.getStudyValue(DefaultStudyValueName, 0)

' now we can calculate the RSI. If any of the inputs to this calculation are
' Empty (for example because we haven't yet got far enough to get genuine values
' from the moving averages), then we will store an Empty Value in the RSI cache.
' This is so that any user of this study doesn't have to bother to work out whether
' a Value for this bar exists or not.
If IsEmpty(gainsAvg.Value) Or IsEmpty(lossesAvg.Value) Then
    RSI = Empty
Else
    If lossesAvg.Value = 0 Then
        RSI = 100
    Else
        RSI = 100 - 100 / (1 + (gainsAvg.Value / lossesAvg.Value))
    End If
End If

If mStudyFoundation.IsNewBar Then
    mRsiCache.AddValue RSI, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
Else
    mRsiCache.UpdateValue RSI, ev.sVal.BarNumber, ev.sVal.Timestamp
End If

notifyValue RSI, RsiValueRsi, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = mPeriods

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
        mPeriods = param.Value
    Case UCase$(RsiParamMovingAverageType)
        mMovingAverageType = param.Value
    End Select
Next

' Now we know what we need to create the two moving average study objects. Note
' that we only need these studies to cache the latest moving average Value.
Set mMaGains = gCreateMA(mStudyManager, mMovingAverageType, mPeriods, 1)
Set mMaLosses = gCreateMA(mStudyManager, mMovingAverageType, mPeriods, 1)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GRSI.StudyDefinition

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

Set mBarsBuilder = ev.Future.Value
Set mBars = mBarsBuilder.Bars

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

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
                ByVal Value As Variant, _
                ByVal valueName As String, _
                ByVal BarNumber As Long, _
                ByVal BarStartTime As Date, _
                ByVal Timestamp As Date)
Dim ev As StudyValueEventData
Const ProcName As String = "notifyValue"
On Error GoTo Err

If Not IsEmpty(Value) Then
    ev.sVal.Value = Value
    ev.valueName = valueName
    Set ev.Source = Me
    ev.sVal.BarNumber = BarNumber
    ev.sVal.BarStartTime = BarStartTime
    ev.sVal.Timestamp = Timestamp
    mStudyFoundation.notifyValue ev
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' From this point on the code is identical for all studies - do not change unless
' you are sure you know what you're doing!!!!
'@================================================================================


'@================================================================================
' IStudy Interface Members
'@================================================================================

Private Function IStudy_AddStudy( _
                ByVal Study As IStudy, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController

Const ProcName As String = "IStudy_AddStudy"
On Error GoTo Err

Set IStudy_AddStudy = mStudyFoundation.AddStudy( _
                            Study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudy_AddStudyValueListener( _
                ByVal listener As IStudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "IStudy_AddStudyValueListener"
On Error GoTo Err

Set IStudy_AddStudyValueListener = mStudyFoundation.AddStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_BaseStudy() As IStudy
Const ProcName As String = "IStudy_BaseStudy"
On Error GoTo Err

Set IStudy_BaseStudy = Me

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IStudy_GetStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Const ProcName As String = "IStudy_GetStudyValue"
On Error GoTo Err

IStudy_GetStudyValue = mStudyFoundation.getStudyValue(valueName, ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudy_GetValueTicksize( _
                ByVal valueName As String) As Double

Const ProcName As String = "IStudy_GetValueTicksize"
On Error GoTo Err



Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_id() As String
Const ProcName As String = "IStudy_id"
On Error GoTo Err

IStudy_id = mStudyFoundation.Id

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStudy_Initialise( _
                ByVal pStudyFoundation As StudyFoundation)
Const ProcName As String = "IStudy_Initialise"
On Error GoTo Err

Set mStudyFoundation = pStudyFoundation

customInitialise mStudyFoundation.Id, _
                mStudyFoundation.Parameters, _
                mStudyFoundation.numberOfCachedValues, _
                mStudyFoundation.inputValueNames, _
                mStudyFoundation.UnderlyingStudy, _
                mStudyFoundation.SessionFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_InstanceName() As String
Const ProcName As String = "IStudy_InstanceName"
On Error GoTo Err

IStudy_InstanceName = mStudyFoundation.InstanceName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudy_InstancePath() As String
Const ProcName As String = "IStudy_InstancePath"
On Error GoTo Err

IStudy_InstancePath = mStudyFoundation.InstancePath

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudy_LibraryName() As String
IStudy_LibraryName = mStudyFoundation.LibraryName
End Property

Private Property Get IStudy_Name() As String
IStudy_Name = mStudyFoundation.name
End Property

Private Sub IStudy_Notify( _
                ByRef ev As StudyValueEventData)
Const ProcName As String = "IStudy_Notify"
On Error GoTo Err

mStudyFoundation.NotifyInput ev
customNotify ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_NumberOfBarsRequired() As Long
Const ProcName As String = "IStudy_NumberOfBarsRequired"
On Error GoTo Err

IStudy_NumberOfBarsRequired = mStudyFoundation.NumberOfBarsRequired(customNumberOfBarsRequired)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IStudy_NumberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "IStudy_NumberOfCachedValues"
On Error GoTo Err

IStudy_NumberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_Parameters() As Parameters
Const ProcName As String = "IStudy_Parameters"
On Error GoTo Err

Set IStudy_Parameters = mStudyFoundation.Parameters

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStudy_RemoveStudyValueListener( _
                ByVal listener As IStudyValueListener)
Const ProcName As String = "IStudy_RemoveStudyValueListener"
On Error GoTo Err

mStudyFoundation.RemoveStudyValueListener listener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStudy_SetInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)

Const ProcName As String = "IStudy_SetInputTicksize"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_StudyDefinition() As StudyDefinition
Const ProcName As String = "IStudy_StudyDefinition"
On Error GoTo Err

Set IStudy_StudyDefinition = customStudyDefinition

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudy_UnderlyingStudy() As IStudy
Const ProcName As String = "IStudy_UnderlyingStudy"
On Error GoTo Err

Set IStudy_UnderlyingStudy = mStudyFoundation.UnderlyingStudy

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property












