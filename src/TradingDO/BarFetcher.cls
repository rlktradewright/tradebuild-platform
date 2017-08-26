VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarFetcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements ITask

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum AsyncStates
    None
    Started
    Fetching
    ProcessingBars
    ProcessingTicks
    AppendingBars
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "BarFetcher"

'@================================================================================
' Member variables
'@================================================================================

Private mSpecifier                          As BarDataSpecifier

Private mInstrument                         As instrument

Private mBarLengthMinutes                   As Long
Private mNumberOfBars                       As Long
Private mFromDate                           As Date
Private mToDate                             As Date

Private WithEvents mConnector               As Connector
Attribute mConnector.VB_VarHelpID = -1
Private WithEvents mConnection              As ADODB.connection
Attribute mConnection.VB_VarHelpID = -1
Private WithEvents mRs                      As ADODB.Recordset
Attribute mRs.VB_VarHelpID = -1

Private mDB                                 As TradingDB

Private mAsynchronous                       As Boolean

Private mSessionBuilder                     As SessionBuilder
Private mSession                            As Session
Attribute mSession.VB_VarHelpID = -1

Private mCustomSessionBuilder               As SessionBuilder
Private mCustomSession                      As Session

Private mAccumulatedVolume                  As Long

Private mNotifiedFirstVolume                As Boolean

Private mTaskContext                        As TaskContext

Private mAsyncState                         As AsyncStates

Private mInterimBarsBuilder                 As BarsBuilder
Private WithEvents mTargetBarsBuilder       As BarsBuilder
Attribute mTargetBarsBuilder.VB_VarHelpID = -1
Private mCurrentBar                         As Bar

Private WithEvents mAppendBarsTaskController    As TaskController
Attribute mAppendBarsTaskController.VB_VarHelpID = -1

Private mCookie                             As Variant
Private mBarListener                        As IBarFetchListener

Private mInterimBarsCollection              As New EnumerableCollection

Private mBarsFutureBuilder                  As New FutureBuilder

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "BarFetcher terminated"
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()
Const ProcName As String = "ITask_Cancel"
On Error GoTo Err

doCancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    doCancel
    Exit Sub
End If

Select Case mAsyncState
Case Started
    gLogger.Log "BarFetcher started for " & mInstrument.ShortName, ProcName, ModuleName, LogLevelDetail
    
    connectToDBAsync
    
    mAsyncState = Fetching
    mTaskContext.Suspend -1
    
Case Fetching
    If useTickData Then
        Set mInterimBarsBuilder = CreateBarsBuilder(mSpecifier.BarTimePeriod, _
                                                    mSession, _
                                                    mInstrument.TickSize, _
                                                    mSpecifier.maxNumberOfBars)
        generateTickRequestDetails mSpecifier.fromDate, mSpecifier.toDate
        Set mRs = doFetchTicks(mFromDate, mToDate, True)
        mAsyncState = ProcessingTicks
    Else
        generateBarRequestDetails mSpecifier.fromDate, mSpecifier.toDate
        Set mRs = doFetchBars(mSpecifier.barType, mBarLengthMinutes, mNumberOfBars, mFromDate, mToDate, True)
        mAsyncState = ProcessingBars
    End If

    mTaskContext.Suspend -1
Case ProcessingBars
    If Not mRs.BOF Then
        If filterByTime(barTimestamp) Then processBar barTimestamp, openValue, highValue, lowValue, closeValue, volume, tickVolume
        mRs.MovePrevious
    Else
        returnAsyncResults
    End If

Case ProcessingTicks
    If Not mRs.EOF Then
        processTickData
        mRs.MoveNext
    Else
        mInterimBarsCollection.Add mInterimBarsBuilder.Bars
            
        If needMoreTicks Then
            generateNextTickDataRequest
            Set mRs = doFetchTicks(mFromDate, mToDate, True)
            mTaskContext.Suspend -1    ' wait for more ticks to be fetched
        Else
            mAsyncState = AppendingBars
        End If
    End If
Case AppendingBars
    If mInterimBarsCollection.Count = 0 Then
        returnAsyncResults
    Else
        Dim lBars As Bars
        If mSpecifier.fromDate = 0 Then
            Set lBars = mInterimBarsCollection(mInterimBarsCollection.Count)
        Else
            Set lBars = mInterimBarsCollection(1)
        End If
        mInterimBarsCollection.Remove lBars
        
        Set mAppendBarsTaskController = mTargetBarsBuilder.AppendBarsAsync(lBars)
        mTaskContext.Suspend -1    ' wait for the append to Finish
    End If
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal value As TaskContext)
Set mTaskContext = value
End Property

Private Property Get ITask_taskName() As String
ITask_taskName = mTaskContext.Name
End Property

'@===============================================================================
' mAppendBarsTaskController Event Handlers
'@===============================================================================

Private Sub mAppendBarsTaskController_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mAppendBarsTaskController_Completed"
On Error GoTo Err

If ev.Cancelled Then
    mTaskContext.Finish Empty, True
    fireFetchCancelled
ElseIf ev.ErrorNumber <> 0 Then
    mTaskContext.Error ev.ErrorNumber, ev.ErrorMessage, ev.ErrorSource
    fireFetchFailed ev.ErrorNumber, ev.ErrorMessage, ev.ErrorSource
Else
    mTaskContext.Continue
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@===============================================================================
' mConnection Event Handlers
'@===============================================================================

Private Sub mConnection_ExecuteComplete( _
                ByVal RecordsAffected As Long, _
                ByVal pError As ADODB.Error, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pCommand As ADODB.Command, _
                ByVal pRecordset As ADODB.Recordset, _
                ByVal pConnection As ADODB.connection)
Const ProcName As String = "mConnection_ExecuteComplete"
On Error GoTo Err

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    
    If pError.Number = 3712 Then
        ' this error is "Operation has been cancelled by the user"
        ' This should no longer happen, but we'll make the check just in case
        connectToDB
        Exit Sub
    End If
    
    Dim s As String
    s = "Fetch error " & vbCrLf & _
        gGenerateErrorMessage(pError) & _
        gGenerateConnectionErrorMessages(pConnection)
        
    If mAsynchronous Then
        If gIsStateSet(mConnection.State, adStateOpen) Then
            Debug.Print "BarFetcher: close connection(error)"
            mConnection.Close
        End If
        mTaskContext.Error pError.Number, _
                            ProjectName & "." & ModuleName & ":" & "mConnection_ExecuteComplete" & vbCrLf & s
        fireFetchFailed pError.Number, s, ""
    Else
        Err.Raise pError.Number, , s
    End If
    Exit Sub
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@===============================================================================
' mConnector Event Handlers
'@===============================================================================

Private Sub mConnector_Connected( _
                ByVal pConnection As ADODB.connection)
Const ProcName As String = "mConnector_Connected"
On Error GoTo Err

Set mConnection = pConnection
If mAsynchronous Then mTaskContext.Continue

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mConnector_ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
Const ProcName As String = "mConnector_ConnectFailed"
On Error GoTo Err

If mAsynchronous Then
    gLogger.Log "BarFetcher connection error for " & mInstrument.ShortName, ProcName, ModuleName, LogLevelSevere
    mTaskContext.Error errorCode, _
                        ProjectName & "." & ModuleName & ":" & "mConnector_ConnectFailed" & vbCrLf & errorDesc
    fireFetchFailed errorCode, errorDesc, ""
Else
    Err.Raise errorCode, , errorDesc
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@===============================================================================
' mRs Event Handlers
'@===============================================================================

Private Sub mRs_FetchComplete( _
                ByVal pError As ADODB.Error, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pRecordset As ADODB.Recordset)
Const ProcName As String = "mRs_FetchComplete"
On Error GoTo Err

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    Dim s As String
    s = "Fetch error " & vbCrLf & _
        gGenerateErrorMessage(pError) & _
        gGenerateConnectionErrorMessages(mConnection)
    
    Debug.Print "BarFetcher: close connection(error)"
    If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
    If mAsynchronous Then
        mTaskContext.Error pError.Number, _
                            ProjectName & "." & ModuleName & ":" & "mRs_FetchComplete" & vbCrLf & s
        fireFetchFailed pError.Number, s, ""
    Else
        Err.Raise pError.Number, , s
    End If
    Exit Sub
End If

If mAsynchronous Then
    If Not useTickData Then
        gLogger.Log "BarFetcher retrieved bars from database for " & mInstrument.ShortName, ProcName, ModuleName, LogLevelDetail
        If Not mRs.BOF Then mRs.MoveLast    ' because the stored procedure returns them in reverse order
        gLogger.Log "BarFetcher repositioned recordset for " & mInstrument.ShortName, ProcName, ModuleName, LogLevelDetail
    End If
    mTaskContext.Continue   ' start calling ITask_run again
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mTargetBarsBuilder Event Handlers
'@================================================================================

Private Sub mTargetBarsBuilder_BarAdded(ByVal pBar As Bar)
Const ProcName As String = "mTargetBarsBuilder_BarAdded"
On Error GoTo Err

Set mCurrentBar = pBar
fireNotifyBar mCurrentBar

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get BarsFuture() As IFuture
Const ProcName As String = "BarsFuture"
On Error GoTo Err

Set BarsFuture = mBarsFutureBuilder.Future

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function FetchBars( _
                ByVal Specifier As BarDataSpecifier, _
                ByVal pListener As IBarFetchListener) As Bars
Const ProcName As String = "FetchBars"
On Error GoTo Err

mAsynchronous = False
Set mBarListener = pListener

initialiseFetch Specifier

connectToDB

If Not useTickData Then
    generateBarRequestDetails mSpecifier.fromDate, mSpecifier.toDate
    Set mRs = doFetchBars(mSpecifier.barType, mBarLengthMinutes, mNumberOfBars, mFromDate, mToDate, False)
    processBars
Else
    Set mInterimBarsBuilder = CreateBarsBuilder(mSpecifier.BarTimePeriod, _
                                                mSession, _
                                                mInstrument.TickSize, _
                                                mSpecifier.maxNumberOfBars)
    generateTickRequestDetails mSpecifier.fromDate, mSpecifier.toDate
    
    Set mRs = doFetchTicks(mFromDate, mToDate, False)
    
    Do While Not mRs.EOF
        processTicks
        mInterimBarsCollection.Add mInterimBarsBuilder.Bars
        
        If Not needMoreTicks Then Exit Do
        
        generateNextTickDataRequest
        Set mRs = doFetchTicks(mFromDate, mToDate, False)
    Loop
        
    Do While mInterimBarsCollection.Count <> 0
        Dim lBars As Bars
        If mSpecifier.fromDate = 0 Then
            Set lBars = mInterimBarsCollection(mInterimBarsCollection.Count)
        Else
            Set lBars = mInterimBarsCollection(1)
        End If
        mInterimBarsCollection.Remove lBars
        
        mTargetBarsBuilder.AppendBars lBars
    Loop
    
End If

Debug.Print "BarFetcher: close connection"
If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close

Set FetchBars = mTargetBarsBuilder.Bars

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Friend Function FetchBarsAsync( _
                ByVal Specifier As BarDataSpecifier, _
                ByVal pListener As IBarFetchListener, _
                ByVal cookie As Variant, _
                ByVal taskName As String) As TaskController
Const ProcName As String = "FetchBarsAsync"
On Error GoTo Err

mAsyncState = Started
mAsynchronous = True
Set mBarListener = pListener
gSetVariant mCookie, cookie

initialiseFetch Specifier

If taskName = "" Then taskName = "BarFetcher/" & mInstrument.ShortName & "(" & mSpecifier.BarTimePeriod.ToString & ")" & FormatTimestamp(mFromDate, TimestampNoMillisecs) & "-" & FormatTimestamp(mToDate, TimestampNoMillisecs) & "(" & gGetSequenceNumber & ")"
Set FetchBarsAsync = StartTask(Me, PriorityNormal, taskName, cookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal pInstrument As instrument)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mDB = db
Set mInstrument = pInstrument

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function barTimestamp() As Date
Const ProcName As String = "barTimestamp"
On Error GoTo Err

barTimestamp = mRs("DateTime")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function closeValue() As Double
Const ProcName As String = "closeValue"
On Error GoTo Err

closeValue = mRs("closeprice")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub connectToDB()
Const ProcName As String = "connectToDB"
On Error GoTo Err

Debug.Print "BarFetcher: Connect to DB"
If mConnector Is Nothing Then Set mConnector = mDB.GetConnector
mConnector.Connect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub connectToDBAsync()
Const ProcName As String = "connectToDBAsync"
On Error GoTo Err

Debug.Print "BarFetcher: Connect to DB async"
If mConnector Is Nothing Then Set mConnector = mDB.GetConnector
mConnector.ConnectAsync

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createNewCustomSession()
Const ProcName As String = "createNewCustomSession"
On Error GoTo Err

If mSpecifier.customSessionEndTime <> 0 Or mSpecifier.customSessionStartTime <> 0 Then
    Set mCustomSessionBuilder = CreateSessionBuilder(mSpecifier.customSessionStartTime, mSpecifier.customSessionEndTime, GetTimeZone(mInstrument.TimeZoneName))
    Set mCustomSession = mCustomSessionBuilder.Session
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createNewSession()
Const ProcName As String = "createNewSession"
On Error GoTo Err

Set mSessionBuilder = CreateSessionBuilder(mInstrument.SessionStartTime, mInstrument.SessionEndTime, GetTimeZone(mInstrument.TimeZoneName))
Set mSession = mSessionBuilder.Session

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doCancel()
Const ProcName As String = "doCancel"
On Error GoTo Err

If Not mRs Is Nothing Then
    If Not gIsStateSet(mRs.State, adStateOpen) Then mRs.Cancel
End If

If Not mConnection Is Nothing Then
    If gIsStateSet(mConnection.State, adStateOpen) Then
        mConnection.Close
    Else
        mConnection.Cancel
    End If
End If

mTaskContext.Finish Empty, True
fireFetchCancelled

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function doFetchBars( _
                ByVal barType As Long, _
                ByVal barLength As Long, _
                ByVal NumberOfBars As Long, _
                ByVal fromTime As Date, _
                ByVal toTime As Date, _
                ByVal pAsync As Boolean) As Recordset
Const ProcName As String = "doFetchBars"
On Error GoTo Err

Set doFetchBars = Execute(setupFetchBarsCommand(barType, barLength, NumberOfBars, fromTime, toTime), pAsync)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function doFetchTicks( _
                ByVal fromTime As Date, _
                ByVal toTime As Date, _
                ByVal pAsync As Boolean) As Recordset
Const ProcName As String = "doFetchTicks"
On Error GoTo Err

Set doFetchTicks = Execute(setupFetchTicksCommand(fromTime, toTime), pAsync)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function Execute(ByVal pCmd As Command, ByVal pAsync As Boolean) As Recordset
Const ProcName As String = "Execute"
On Error GoTo Err

'gLogger.Log "BarFetcher initiating fetch for " & mInstrument.ShortName, ProcName, ModuleName

pCmd.ActiveConnection = mConnection

Dim lRs As New ADODB.Recordset
lRs.LockType = adLockReadOnly
lRs.cursorlocation = adUseClient

If pAsync Then
    lRs.Properties("Initial Fetch Size") = 0
    lRs.Open pCmd, _
                , _
                CursorTypeEnum.adOpenStatic, _
                , _
                CommandTypeEnum.adCmdStoredProc + _
                ExecuteOptionEnum.adAsyncFetchNonBlocking + _
                ExecuteOptionEnum.adAsyncExecute
Else
    lRs.Open pCmd, _
                , _
                CursorTypeEnum.adOpenForwardOnly, _
                , _
                CommandTypeEnum.adCmdStoredProc
End If

Set Execute = lRs

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function filterByTime( _
                ByVal timestamp As Date) As Boolean
Const ProcName As String = "filterByTime"
On Error GoTo Err

If timestamp < mFromDate Or timestamp >= mToDate Then Exit Function

Dim lSessEv As SessionEventData
lSessEv = mSessionBuilder.SetSessionCurrentTime(timestamp)
If lSessEv.ChangeType = SessionChangeStart Then
    mAccumulatedVolume = 0
End If

If Not mCustomSessionBuilder Is Nothing Then
    mCustomSessionBuilder.SetSessionCurrentTime timestamp
    If Not mCustomSession.IsTimeInSession(timestamp) Then Exit Function
End If

If mSpecifier.includeBarsOutsideSession Or _
    mSession.IsTimeInSession(timestamp) _
Then
    filterByTime = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub fireFetchCancelled()
Const ProcName As String = "fireFetchCancelled"
On Error GoTo Err

If mBarListener Is Nothing Then Exit Sub

mBarListener.FetchCancelled mCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireFetchCompleted()
Const ProcName As String = "fireFetchCompleted"
On Error GoTo Err

If mBarListener Is Nothing Then Exit Sub

mBarListener.FetchCompleted mCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireFetchFailed(ByVal pErrorNumber As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
Const ProcName As String = "fireFetchFailed"
On Error GoTo Err

If mBarListener Is Nothing Then Exit Sub

mBarListener.FetchFailed mCookie, pErrorNumber, pErrorMessage, pErrorSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireNotifyBar(ByVal pBar As Bar)
Const ProcName As String = "fireNotifyBar"
On Error GoTo Err

If pBar Is Nothing Or mBarListener Is Nothing Then Exit Sub

mBarListener.NotifyBar mCookie, pBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub generateNextTickDataRequest()
Const ProcName As String = "generateNextTickDataRequest"
On Error GoTo Err

If mSpecifier.fromDate = 0 Then
    generateTickRequestDetails 0, mFromDate
Else
    generateTickRequestDetails mToDate, 0
End If

mNotifiedFirstVolume = False
createNewSession
createNewCustomSession

Set mInterimBarsBuilder = CreateBarsBuilder(mSpecifier.BarTimePeriod, _
                                            mSession, _
                                            mInstrument.TickSize, _
                                            mSpecifier.maxNumberOfBars)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub generateBarRequestDetails( _
                        ByVal fromDate As Date, _
                        ByVal toDate As Date)
Const ProcName As String = "generateBarRequestDetails"
On Error GoTo Err

If toDate <> 0 Then
    mToDate = toDate
    
    If fromDate = 0 Then
        ' calculate the earliest possible date for the supplied number of bars (assuming
        ' there are no gaps in the data)
        fromDate = OffsetBarStartTime(IIf(mToDate = MaxDateValue, Now, mToDate), _
                                    mSpecifier.BarTimePeriod, _
                                    -1 * mSpecifier.maxNumberOfBars, _
                                    mInstrument.SessionStartTime, _
                                    mInstrument.SessionEndTime)
    Else
        mFromDate = fromDate
    End If
Else
    mFromDate = fromDate
    mToDate = OffsetBarStartTime(fromDate, _
                                mSpecifier.BarTimePeriod, _
                                mSpecifier.maxNumberOfBars, _
                                mInstrument.SessionStartTime, _
                                mInstrument.SessionEndTime)
End If

Dim lNumberOfTargetBars As Long
lNumberOfTargetBars = MaxNumberOfBarsInTimespan(mSpecifier.BarTimePeriod, fromDate, mToDate, mInstrument.SessionStartTime, mInstrument.SessionEndTime)

Select Case mSpecifier.BarTimePeriod.Units
Case TimePeriodUnits.TimePeriodDay
    mBarLengthMinutes = 60
    mNumberOfBars = lNumberOfTargetBars * 24 * mSpecifier.BarTimePeriod.Length
Case TimePeriodUnits.TimePeriodHour
    mBarLengthMinutes = 60
    mNumberOfBars = lNumberOfTargetBars * mSpecifier.BarTimePeriod.Length
Case TimePeriodUnits.TimePeriodMinute
    If (mSpecifier.BarTimePeriod.Length Mod 60) = 0 Then
        mBarLengthMinutes = 60
        mNumberOfBars = lNumberOfTargetBars * (mSpecifier.BarTimePeriod.Length / 60)
    ElseIf (mSpecifier.BarTimePeriod.Length Mod 15) = 0 Then
        mBarLengthMinutes = 15
        mNumberOfBars = lNumberOfTargetBars * (mSpecifier.BarTimePeriod.Length / 15)
    ElseIf (mSpecifier.BarTimePeriod.Length Mod 5) = 0 Then
        mBarLengthMinutes = 5
        mNumberOfBars = lNumberOfTargetBars * (mSpecifier.BarTimePeriod.Length / 5)
    Else
        mBarLengthMinutes = 1
        mNumberOfBars = lNumberOfTargetBars * mSpecifier.BarTimePeriod.Length
    End If
Case TimePeriodUnits.TimePeriodMonth
    mBarLengthMinutes = 60
    mNumberOfBars = lNumberOfTargetBars * 24 * 22 * mSpecifier.BarTimePeriod.Length
Case TimePeriodUnits.TimePeriodWeek
    mBarLengthMinutes = 60
    mNumberOfBars = lNumberOfTargetBars * 24 * 5 * mSpecifier.BarTimePeriod.Length
Case TimePeriodUnits.TimePeriodYear
    mBarLengthMinutes = 60
    mNumberOfBars = lNumberOfTargetBars * 24 * 260 * mSpecifier.BarTimePeriod.Length
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub generateTickRequestDetails( _
                        ByVal fromDate As Date, _
                        ByVal toDate As Date)
Const ProcName As String = "generateTickRequestDetails"
On Error GoTo Err

If fromDate <> 0 And toDate <> 0 Then
    mFromDate = GetSessionTimes(fromDate, _
                                mInstrument.SessionStartTime, _
                                mInstrument.SessionEndTime).StartTime
    mToDate = toDate
End If

If fromDate <> 0 And toDate = 0 Then
    mFromDate = GetSessionTimes(fromDate, _
                                mInstrument.SessionStartTime, _
                                mInstrument.SessionEndTime).StartTime
    mToDate = GetOffsetSessionTimes(mFromDate, _
                                1, _
                                mInstrument.SessionStartTime, _
                                mInstrument.SessionEndTime).StartTime
End If

If fromDate = 0 And toDate <> 0 Then
    mToDate = toDate
    mFromDate = GetOffsetSessionTimes(IIf(mToDate = MaxDateValue, Now, mToDate), _
                                -1, _
                                mInstrument.SessionStartTime, _
                                mInstrument.SessionEndTime).StartTime
End If

If fromDate = 0 And toDate = 0 Then
    mToDate = Now
    mFromDate = GetOffsetSessionTimes(mToDate, _
                                -1, _
                                mInstrument.SessionStartTime, _
                                mInstrument.SessionEndTime).StartTime
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getNumberOfPrices( _
                ByVal pOpenValue As Double, _
                ByVal pHighValue As Double, _
                ByVal pLowValue As Double, _
                ByVal pCloseValue As Double) As Long
getNumberOfPrices = 1
If pOpenValue <> pHighValue Then getNumberOfPrices = getNumberOfPrices + 1
If pOpenValue <> pLowValue Then getNumberOfPrices = getNumberOfPrices + 1
If pCloseValue <> pHighValue And pCloseValue <> pLowValue Then getNumberOfPrices = getNumberOfPrices + 1
End Function

Private Function highValue() As Double
Const ProcName As String = "highValue"
On Error GoTo Err

highValue = mRs("highprice")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub initialiseFetch(ByVal Specifier As BarDataSpecifier)
Const ProcName As String = "initialiseFetch"
On Error GoTo Err

Set mSpecifier = Specifier
If mSpecifier.maxNumberOfBars = 0 Then
    Dim lSessionStartTime As Date
    lSessionStartTime = IIf(mSpecifier.customSessionStartTime <> 0, mSpecifier.customSessionStartTime, mInstrument.SessionStartTime)
    
    Dim lSessionEndTime As Date
    lSessionEndTime = IIf(mSpecifier.customSessionEndTime <> 0, mSpecifier.customSessionEndTime, mInstrument.SessionEndTime)
    
    mSpecifier.maxNumberOfBars = MaxNumberOfBarsInTimespan( _
                                                mSpecifier.BarTimePeriod, _
                                                mSpecifier.fromDate, _
                                                mSpecifier.toDate, _
                                                lSessionStartTime, _
                                                lSessionEndTime)
End If

createNewSession
createNewCustomSession

Set mTargetBarsBuilder = CreateBarsBuilder(mSpecifier.BarTimePeriod, _
                                            mSession, _
                                            mInstrument.TickSize, _
                                            mSpecifier.maxNumberOfBars)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "logMessage"
gLogger.Log pMsg, pProcName, ModuleName, pLogLevel, pMsgQualifier
End Sub

Private Function lowValue() As Double
Const ProcName As String = "lowValue"
On Error GoTo Err

lowValue = mRs("lowprice")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function needMoreTicks() As Boolean
Const ProcName As String = "needMoreTicks"
On Error GoTo Err

Dim lTotalBars As Long
Dim lBars As Bars
For Each lBars In mInterimBarsCollection
    lTotalBars = lTotalBars + lBars.Count
Next

needMoreTicks = lTotalBars < mSpecifier.maxNumberOfBars And _
            (mSpecifier.fromDate = 0 Or mSpecifier.toDate = 0)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub notifyBarValue( _
                ByVal pValue As Double, _
                ByVal pNumberOfPrices As Long, _
                ByRef pNumberOfPricesNotified As Long, _
                ByVal pTimestamp As Date, _
                ByVal pTickVolumePerNotify As Long, _
                ByVal pTickVolume As Long, _
                ByVal pVolumePerNotify As Long, _
                ByVal pVolume As Long)
Const ProcName As String = "notifyBarValue"
On Error GoTo Err

Dim lValue As SValue
lValue.timestamp = pTimestamp

lValue.value = pValue
mTargetBarsBuilder.NotifyValue lValue
pNumberOfPricesNotified = pNumberOfPricesNotified + 1

mTargetBarsBuilder.IncrementTickVolume IIf(pNumberOfPricesNotified < pNumberOfPrices, pTickVolumePerNotify, pTickVolume)
lValue.value = mAccumulatedVolume + IIf(pNumberOfPricesNotified < pNumberOfPrices, pNumberOfPricesNotified * pVolumePerNotify, pVolume)
mTargetBarsBuilder.NotifyVolume lValue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function openInterest() As Long
Const ProcName As String = "openInterest"
On Error GoTo Err

openInterest = mRs("openInterest")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function openValue() As Double
Const ProcName As String = "openValue"
On Error GoTo Err

openValue = mRs("openprice")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processBar( _
    ByVal pTimestamp As Date, _
    ByVal pOpenValue As Double, _
    ByVal pHighValue As Double, _
    ByVal pLowValue As Double, _
    ByVal pCloseValue As Double, _
    ByVal pVolume As Long, _
    ByVal pTickVolume As Long)
Const ProcName As String = "processBar"
On Error GoTo Err

Dim lNumberOfPrices As Long
lNumberOfPrices = getNumberOfPrices(pOpenValue, pHighValue, pLowValue, pCloseValue)

Dim lNumberOfPricesNotified As Long

Dim lTickVolumePerNotify As Long
lTickVolumePerNotify = Int((pTickVolume - lNumberOfPrices) / lNumberOfPrices)
If lTickVolumePerNotify < 0 Then lTickVolumePerNotify = 0

Dim lVolumePerNotify As Long
lVolumePerNotify = Int((pVolume - lNumberOfPrices) / lNumberOfPrices)
If lVolumePerNotify < 0 Then lVolumePerNotify = 0

notifyBarValue pOpenValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume

If pCloseValue = pHighValue Then
    If pLowValue = pOpenValue And pLowValue = pHighValue Then
    ElseIf pLowValue = pOpenValue Then
        notifyBarValue pHighValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    Else
        notifyBarValue pLowValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
        notifyBarValue pHighValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    End If
ElseIf pCloseValue = pLowValue Then
    If pHighValue = pOpenValue And pHighValue = pLowValue Then
    ElseIf pHighValue = pOpenValue Then
        notifyBarValue pLowValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    Else
        notifyBarValue pHighValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
        notifyBarValue pLowValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    End If
Else
    If pLowValue <> pOpenValue Then notifyBarValue pLowValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    If pHighValue <> pOpenValue Then notifyBarValue pHighValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    notifyBarValue pCloseValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
End If

mAccumulatedVolume = mAccumulatedVolume + pVolume

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processBars()
Const ProcName As String = "processBars"
On Error GoTo Err

If Not mRs.BOF Then
    mRs.MoveLast    ' because the stored procedure returns them in reverse order
End If

Do While Not mRs.BOF
    If filterByTime(barTimestamp) Then processBar barTimestamp, openValue, highValue, lowValue, closeValue, volume, tickVolume
    mRs.MovePrevious
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTickData()
Const ProcName As String = "processTickData"
On Error GoTo Err

If Not filterByTime(mRs("datetime")) Then Exit Sub

Dim Data() As Byte
Data = mRs("data")

Dim tdd As ITickDataDecoder
Set tdd = CreateTickDecoder(mRs("datetime"), _
                            mRs("tickSize"), _
                            mRs("baseprice"), _
                            Data, _
                            mRs("version"))

Dim tick As GenericTick
Do While tdd.GetNextTick(tick)
    Dim val As SValue
    val.timestamp = tick.timestamp
    Select Case tick.TickType
    Case TickTypeBid
        If mSpecifier.barType = BarTypeBid Then
            val.value = tick.price
            mInterimBarsBuilder.NotifyValue val
        End If
    Case TickTypeAsk
        If mSpecifier.barType = BarTypeAsk Then
            val.value = tick.price
            mInterimBarsBuilder.NotifyValue val
        End If
    Case TickTypeTrade
        If mSpecifier.barType = BarTypeTrade Then
            val.value = tick.price
            mInterimBarsBuilder.NotifyValue val
        End If
    Case TickTypeVolume
        If mSpecifier.barType = BarTypeTrade Then
            If Not mNotifiedFirstVolume Then
                val.value = 1
                mInterimBarsBuilder.NotifyVolume val
                mNotifiedFirstVolume = True
            End If
            val.value = tick.size + 1
            mInterimBarsBuilder.NotifyVolume val
        End If
    Case TickTypeOpenInterest
        If mSpecifier.barType = BarTypeTrade Then
            val.value = tick.size
            mInterimBarsBuilder.NotifyOpenInterest val
        End If
    End Select
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTicks()
Const ProcName As String = "processTicks"
On Error GoTo Err

Do While Not mRs.EOF
    processTickData
    mRs.MoveNext
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub returnAsyncResults()
Const ProcName As String = "returnAsyncResults"
On Error GoTo Err

Debug.Print "BarFetcher: close connection"
If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close

gLogger.Log "BarFetcher fetched " & mTargetBarsBuilder.Bars.Count & " bars for " & mInstrument.ShortName, ProcName, ModuleName
setBarsFutureValue

mTaskContext.Finish mTargetBarsBuilder.Bars, False

fireFetchCompleted

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setBarsFutureValue()
Const ProcName As String = "setBarsFutureValue"
On Error GoTo Err

If Not mBarsFutureBuilder.Future.IsPending Then Exit Sub

mBarsFutureBuilder.value = mTargetBarsBuilder.Bars
mBarsFutureBuilder.Complete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setupFetchBarsCommand( _
                        ByVal barType As Long, _
                        ByVal barLength As Long, _
                        ByVal NumberOfBars As Long, _
                        ByVal fromTime As Date, _
                        ByVal toTime As Date) As Command
Const ProcName As String = "setupFetchBarsCommand"
On Error GoTo Err

fromTime = gRoundTimeToSecond(fromTime)
toTime = gRoundTimeToSecond(toTime)

Dim lCmd As New ADODB.Command
lCmd.CommandType = adCmdStoredProc

lCmd.CommandText = "FetchBarData"

Dim param As ADODB.Parameter
' @InstrumentID
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mInstrument.Id)
lCmd.Parameters.Append param

' @BarType
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            barType)
lCmd.Parameters.Append param

' @BarLength
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            barLength)
lCmd.Parameters.Append param

' @NumberRequired
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            NumberOfBars)
lCmd.Parameters.Append param

If fromTime < #1/1/1900# Then fromTime = #1/1/1900# ' don't exceed range of SmallDateTime
If toTime < #1/1/1900# Then toTime = #1/1/1900# ' don't exceed range of SmallDateTime
' @From
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            fromTime)
lCmd.Parameters.Append param

' @To
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            toTime)
lCmd.Parameters.Append param

Set setupFetchBarsCommand = lCmd

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupFetchTicksCommand( _
                ByVal fromTime As Date, _
                ByVal toTime As Date) As Command
Const ProcName As String = "setupFetchTicksCommand"
On Error GoTo Err

fromTime = gRoundTimeToSecond(fromTime)
toTime = gRoundTimeToSecond(toTime)

Dim lCmd As New ADODB.Command
lCmd.CommandType = adCmdStoredProc

lCmd.CommandText = "FetchTickData"

Dim param As ADODB.Parameter

' @InstrumentID
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mInstrument.Id)
lCmd.Parameters.Append param

' @From
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            fromTime)
lCmd.Parameters.Append param

' @To
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            toTime)
lCmd.Parameters.Append param

Set setupFetchTicksCommand = lCmd

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function tickVolume() As Long
Const ProcName As String = "tickVolume"
On Error GoTo Err

tickVolume = mRs("tickVolume")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function useTickData() As Boolean
useTickData = mSpecifier.BarTimePeriod.Units = TimePeriodVolume Or _
    mSpecifier.BarTimePeriod.Units = TimePeriodTickMovement Or _
    mSpecifier.BarTimePeriod.Units = TimePeriodTickVolume Or _
    mSpecifier.BarTimePeriod.Units = TimePeriodSecond
End Function

Private Function volume() As Long
Const ProcName As String = "volume"
On Error GoTo Err

volume = mRs("Volume")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



