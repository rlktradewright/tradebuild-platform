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

Implements Task

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum AsyncStates
    None
    Started
    Connecting
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

Private Const ProjectName                   As String = "TradingDO26"
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

Private mRequestTicks                       As Boolean

Private WithEvents mConnector               As Connector
Attribute mConnector.VB_VarHelpID = -1
Private WithEvents mConnection              As ADODB.connection
Attribute mConnection.VB_VarHelpID = -1
Private WithEvents mRs                      As ADODB.Recordset
Attribute mRs.VB_VarHelpID = -1

Private mDB                                 As TradingDB
Private mCmd                                As ADODB.Command

Private mAsynchronous                       As Boolean

Private mSessionBuilder                     As SessionBuilder
Private WithEvents mSession                 As Session
Attribute mSession.VB_VarHelpID = -1

Private mCustomSessionBuilder               As SessionBuilder
Private mCustomSession                      As Session

Private mAccumulatedVolume                  As Long

Private mNotifiedFirstVolume                As Boolean

Private mTaskContext                        As TaskContext

Private mAsyncState                         As AsyncStates

Private mBuilder                            As BarsBuilder
Private mFetchBarsBuilder                   As BarsBuilder

Private WithEvents mAppendBarsTaskController    As TaskController
Attribute mAppendBarsTaskController.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "BarFetcher terminated"
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_cancel()

If Not mRs Is Nothing Then
    If mRs.State <> adStateOpen Then mRs.cancel
End If

If Not mCmd Is Nothing Then mCmd.cancel

If Not mConnection Is Nothing Then
    If mConnection.State = adStateOpen Then
        mConnection.Close
    Else
        mConnection.cancel
    End If
End If

mTaskContext.finish Empty, True
End Sub

Private Sub Task_run()
Static gotFirstTicks As Boolean

Select Case mAsyncState
Case Started

    gLogger.Log LogLevelDetail, "BarFetcher started for " & mInstrument.shortName
    
    If mTaskContext.cancelPending Then
        mTaskContext.finish Empty, True
        Exit Sub
    End If
    
    connectToDB
    
    generateRequestDetails mSpecifier.fromDate, mSpecifier.toDate
    
    If Not mRequestTicks Then
        setupFetchBarsCommand mSpecifier.barType, _
                            mBarLengthMinutes, _
                            mNumberOfBars, _
                            mFromDate, _
                            mToDate
    Else
        setupFetchTicksCommand mFromDate, mToDate
    End If
    
    mTaskContext.suspend -1
    
Case Connecting
    gLogger.Log LogLevelNormal, mTaskContext.name & ": Task_run invoked in Connecting state"
    gLogger.Log LogLevelNormal, "Runnable tasks--------------------------------------------"
    gLogger.Log LogLevelNormal, TaskManager.getRunnableTaskSummary
    gLogger.Log LogLevelNormal, "Suspended tasks-------------------------------------------"
    gLogger.Log LogLevelNormal, TaskManager.getSuspendedTaskSummary
    mTaskContext.Error ErrorCodes.ErrIllegalStateException, _
                        "Task_run invoked in Connecting state"
Case Fetching
    gLogger.Log LogLevelNormal, mTaskContext.name & ": Task_run invoked in Fetching state"
    gLogger.Log LogLevelNormal, "Runnable tasks--------------------------------------------"
    gLogger.Log LogLevelNormal, TaskManager.getRunnableTaskSummary
    gLogger.Log LogLevelNormal, "Suspended tasks-------------------------------------------"
    gLogger.Log LogLevelNormal, TaskManager.getSuspendedTaskSummary
    mTaskContext.Error ErrorCodes.ErrIllegalStateException, _
                        "Task_run invoked in Fetching state"
Case ProcessingBars
    
    If mTaskContext.cancelPending Then
        mTaskContext.finish Empty, True
        Exit Sub
    End If
    
    If Not mRs.BOF Then
        If filterByTime(mRs("DateTime")) Then processBar
        mRs.MovePrevious
    Else
        Debug.Print "BarFetcher: close connection"
        If mConnection.State = adStateOpen Then mConnection.Close
        gLogger.Log LogLevelDetail, "BarFetcher fetched " & builder.Bars.count & " bars for " & mInstrument.shortName
        mTaskContext.finish builder.Bars, False
    End If

Case ProcessingTicks
    
    If mTaskContext.cancelPending Then
        mTaskContext.finish Empty, True
        Exit Sub
    End If
    
    If Not noTicks Then
        processTickData
        mRs.MoveNext
    Else
    
        If Not gotFirstTicks Then
            gotFirstTicks = True
            If needMoreTicks Then
                connectToDB
                generateNextRequest
                setupFetchTicksCommand mFromDate, mToDate
                mTaskContext.suspend -1    ' wait for more ticks to be fetched
            Else
                gLogger.Log LogLevelDetail, "BarFetcher fetched " & builder.Bars.count & " bars for " & mInstrument.shortName
                Debug.Print "BarFetcher: close connection"
                If mConnection.State = adStateOpen Then mConnection.Close
                mTaskContext.finish mFetchBarsBuilder.Bars, False
            End If
        Else
            If mSpecifier.fromDate = 0 Then
                Set mAppendBarsTaskController = builder.appendBarsAsync(mFetchBarsBuilder.Bars)
                Set mFetchBarsBuilder = builder
            Else
                Set mAppendBarsTaskController = mFetchBarsBuilder.appendBarsAsync(builder.Bars)
            End If
            mTaskContext.suspend -1    ' wait for the append to finish
        End If
    
    End If
Case AppendingBars
    mTaskContext.Error ErrorCodes.ErrIllegalStateException, _
                        "Task_run invoked in AppendingBars state"
End Select
End Sub

Private Property Let Task_TaskContext(ByVal value As TWUtilities30.TaskContext)
Set mTaskContext = value
End Property

Private Property Get Task_taskName() As String
Task_taskName = mTaskContext.name
End Property

'@===============================================================================
' mAppendBarsTaskController Event Handlers
'@===============================================================================

Private Sub mAppendBarsTaskController_Completed(ev As TWUtilities30.TaskCompletionEvent)
If ev.cancelled Then
    mTaskContext.finish Empty, True
ElseIf ev.errorNumber <> 0 Then
    mTaskContext.Error ev.errorNumber, ev.errorMessage, ev.errorSource
Else
    If needMoreTicks Then
        connectToDB
        generateNextRequest
        setupFetchTicksCommand mFromDate, mToDate
        mAsyncState = Fetching
    Else
        Debug.Print "BarFetcher: close connection"
        If mConnection.State = adStateOpen Then mConnection.Close
        mTaskContext.finish mFetchBarsBuilder.Bars, False
    End If
End If
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
Dim s As String

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    
    If pError.Number = 3712 Then
        ' this error is "Operation has been cancelled by the user"
        ' This occurs occasionally when running async, but I've no idea why
        connectToDB
        Exit Sub
    End If
    
    s = "Fetch error " & vbCrLf & _
        gGenerateErrorMessage(pError) & _
        gGenerateConnectionErrorMessages(pConnection)
        
    If mAsynchronous Then
        If mConnection.State = adStateOpen Then
            Debug.Print "BarFetcher: close connection(error)"
            mConnection.Close
        End If
        mTaskContext.Error pError.Number, _
                            ProjectName & "." & ModuleName & ":" & "mConnection_ExecuteComplete" & vbCrLf & s
    Else
        Err.Raise pError.Number, _
                ProjectName & "." & ModuleName & ":" & "mConnection_ExecuteComplete", _
                s
    End If
    Exit Sub
End If
End Sub

'@===============================================================================
' mConnector Event Handlers
'@===============================================================================

Private Sub mConnector_Connected( _
                ByVal pConnection As ADODB.connection)
Set mConnection = pConnection
If mAsynchronous Then
    Execute
    mAsyncState = Fetching
End If
End Sub

Private Sub mConnector_ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
If mAsynchronous Then
    gLogger.Log LogLevelDetail, "BarFetcher connection error for " & mInstrument.shortName
    mTaskContext.Error errorCode, _
                        ProjectName & "." & ModuleName & ":" & "mConnector_ConnectFailed" & vbCrLf & errorDesc
Else
    Err.Raise errorCode, _
                ProjectName & "." & ModuleName & ":" & "mConnector_ConnectFailed", _
                errorDesc
End If
End Sub

'@===============================================================================
' mRs Event Handlers
'@===============================================================================

Private Sub mRs_FetchComplete( _
                ByVal pError As ADODB.Error, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pRecordset As ADODB.Recordset)
Dim s As String
Dim failpoint As Long
On Error GoTo Err

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    s = "Fetch error " & vbCrLf & _
        gGenerateErrorMessage(pError) & _
        gGenerateConnectionErrorMessages(mConnection)
    
    Debug.Print "BarFetcher: close connection(error)"
    If mConnection.State = adStateOpen Then mConnection.Close
    If mAsynchronous Then
        mTaskContext.Error pError.Number, _
                            ProjectName & "." & ModuleName & ":" & "mRs_FetchComplete" & vbCrLf & s
    Else
        Err.Raise pError.Number, _
                ProjectName & "." & ModuleName & ":" & "mRs_FetchComplete", _
                s
    End If
    Exit Sub
End If

'Set mrs.ActiveConnection = Nothing ' ADO docs say this is not allowed for a record set
                                    ' that uses a command object - but it works most
                                    ' of the time !
If mAsynchronous Then
    If mRequestTicks Then
        If noTicks Then
            gLogger.Log LogLevelDetail, "BarFetcher fetched " & builder.Bars.count & " bars for " & mInstrument.shortName
            Debug.Print "BarFetcher: close connection"
            If mConnection.State = adStateOpen Then mConnection.Close
            mTaskContext.finish mFetchBarsBuilder.Bars, False
        Else
            mAsyncState = ProcessingTicks
        End If
    Else
        If Not mRs.BOF Then
            mRs.MoveLast    ' because the stored procedure returns them in reverse order
        End If
        mAsyncState = ProcessingBars
    End If
    mTaskContext.continue   ' start calling Task_run again
End If

Exit Sub

Err:
mTaskContext.Error Err.Number, _
        Err.Description, _
        ProjectName & "." & ModuleName & ":" & "mRs_FetchComplete" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, "")


End Sub

'@================================================================================
' mSession Event Handlers
'@================================================================================

Private Sub mSession_SessionStarted(ev As SessionEvent)
mAccumulatedVolume = 0
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Function fetchBars( _
                ByVal specifier As BarDataSpecifier) As Bars

mAsynchronous = False
Set mSpecifier = specifier

createNewCustomSession

generateRequestDetails mSpecifier.fromDate, mSpecifier.toDate

connectToDB

Set mFetchBarsBuilder = builder

If Not mRequestTicks Then
    doFetchBars mSpecifier.barType, _
                mBarLengthMinutes, _
                mNumberOfBars, _
                mFromDate, _
                mToDate
                
    Set fetchBars = processBars
Else
    
    doFetchTicks mFromDate, mToDate
    If noTicks Then
    Else
        processTicks
        
        Do While needMoreTicks
            generateNextRequest
            doFetchBars mSpecifier.barType, _
                        mBarLengthMinutes, _
                        mNumberOfBars, _
                        mFromDate, _
                        mToDate
            processTicks
            If mSpecifier.fromDate = 0 Then
                builder.appendBars mFetchBarsBuilder.Bars
                Set mFetchBarsBuilder = builder
            Else
                mFetchBarsBuilder.appendBars builder.Bars
            End If
        Loop
        
    End If
    Set fetchBars = mFetchBarsBuilder.Bars
End If

Debug.Print "BarFetcher: close connection"
If mConnection.State = adStateOpen Then mConnection.Close

End Function

Friend Function fetchbarsAsync( _
                ByVal specifier As BarDataSpecifier, _
                ByVal cookie As Variant, _
                ByVal taskName As String) As TaskController
mAsyncState = Started

mAsynchronous = True
Set mSpecifier = specifier

createNewCustomSession

Set mFetchBarsBuilder = builder

If taskName = "" Then taskName = "BarFetcher/" & mInstrument.shortName & "(" & mSpecifier.barLength & TimePeriodUnitsToString(mSpecifier.barLengthUnits) & ")" & FormatTimestamp(mFromDate, TimestampNoMillisecs) & "-" & FormatTimestamp(mToDate, TimestampNoMillisecs) & "(" & gGetSequenceNumber & ")"

Set fetchbarsAsync = StartTask(Me, _
                            PriorityNormal, _
                            taskName, _
                            cookie)
End Function

Friend Sub initialise( _
                ByVal db As TradingDB, _
                ByVal pInstrument As instrument)
Set mDB = db
Set mInstrument = pInstrument
createNewSession

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function builder() As BarsBuilder
If mBuilder Is Nothing Then Set mBuilder = CreateBarsBuilder(mSpecifier.barLength, _
                                                            mSpecifier.barLengthUnits, _
                                                            mSession, _
                                                            mInstrument.tickSize, _
                                                            mSpecifier.maxNumberOfBars)
Set builder = mBuilder
End Function

Private Function closeValue() As Double
closeValue = mRs("closeprice")
End Function

Private Sub connectToDB()
Debug.Print "BarFetcher: Connect to DB"
If mConnector Is Nothing Then Set mConnector = mDB.getConnector(, IIf(mAsynchronous, ConnectOptionEnum.adAsyncConnect, 0))
mConnector.connect
If mAsynchronous Then mAsyncState = Connecting
End Sub

Private Sub createNewCustomSession()
If mSpecifier.customSessionEndTime <> 0 Or mSpecifier.customSessionStartTime <> 0 Then
    Set mCustomSessionBuilder = New SessionBuilder
    mCustomSessionBuilder.sessionStartTime = mSpecifier.customSessionStartTime
    mCustomSessionBuilder.sessionEndTime = mSpecifier.customSessionEndTime
    mCustomSessionBuilder.TimeZone = GetTimeZone(mInstrument.TimeZoneName)
    Set mCustomSession = mCustomSessionBuilder.Session
End If
End Sub

Private Sub createNewSession()
Set mSessionBuilder = New SessionBuilder
mSessionBuilder.sessionStartTime = mInstrument.sessionStartTime
mSessionBuilder.sessionEndTime = mInstrument.sessionEndTime
mSessionBuilder.TimeZone = GetTimeZone(mInstrument.TimeZoneName)
Set mSession = mSessionBuilder.Session
End Sub

Private Sub doFetchBars( _
                        ByVal barType As Long, _
                        ByVal barLength As Long, _
                        ByVal NumberOfBars As Long, _
                        ByVal fromTime As Date, _
                        ByVal toTime As Date)

setupFetchBarsCommand barType, barLength, NumberOfBars, fromTime, toTime

If Not mAsynchronous Then Execute
End Sub

Private Sub doFetchTicks( _
                ByVal fromTime As Date, _
                ByVal toTime As Date)
setupFetchTicksCommand fromTime, toTime

If Not mAsynchronous Then Execute
End Sub

Private Sub Execute()
gLogger.Log LogLevelDetail, "BarFetcher initiating fetch for " & mInstrument.shortName

mCmd.ActiveConnection = mConnection

Set mRs = New ADODB.Recordset
mRs.LockType = adLockReadOnly
mRs.cursorlocation = adUseClient

If mAsynchronous Then
    mRs.Properties("Initial Fetch Size") = 0
    mRs.Open mCmd, _
                , _
                CursorTypeEnum.adOpenStatic, _
                , _
                CommandTypeEnum.adCmdStoredProc + _
                ExecuteOptionEnum.adAsyncFetchNonBlocking + _
                ExecuteOptionEnum.adAsyncExecute
Else
    mRs.Open mCmd, _
                , _
                CursorTypeEnum.adOpenForwardOnly, _
                , _
                CommandTypeEnum.adCmdStoredProc
End If
End Sub

Private Function filterByTime( _
                ByVal timestamp As Date) As Boolean
mSessionBuilder.setSessionCurrentTime timestamp
If Not mCustomSessionBuilder Is Nothing Then
    mCustomSessionBuilder.setSessionCurrentTime timestamp
    If Not mCustomSession.isTimeInSession(timestamp) Then Exit Function
End If

If mSpecifier.includeBarsOutsideSession Or _
    mSession.isTimeInSession(timestamp) _
Then
    filterByTime = True
End If
End Function

Private Sub generateNextRequest()
If mSpecifier.fromDate = 0 Then
    generateRequestDetails 0, mFromDate
Else
    generateRequestDetails mToDate, 0
End If

Set mBuilder = Nothing
mNotifiedFirstVolume = False
createNewSession
End Sub

Private Sub generateRequestDetails( _
                        ByVal fromDate As Date, _
                        ByVal toDate As Date)
Dim st As SessionTimes

Select Case mSpecifier.barLengthUnits
Case TimePeriodUnits.TimePeriodDay
    mBarLengthMinutes = 60
    mNumberOfBars = mSpecifier.maxNumberOfBars * 24 * mSpecifier.barLength
Case TimePeriodUnits.TimePeriodHour
    mBarLengthMinutes = 60
    mNumberOfBars = mSpecifier.maxNumberOfBars * mSpecifier.barLength
Case TimePeriodUnits.TimePeriodMinute
    If (mSpecifier.barLength Mod 60) = 0 Then
        mBarLengthMinutes = 60
        mNumberOfBars = mSpecifier.maxNumberOfBars * (mSpecifier.barLength / 60)
    ElseIf (mSpecifier.barLength Mod 15) = 0 Then
        mBarLengthMinutes = 15
        mNumberOfBars = mSpecifier.maxNumberOfBars * (mSpecifier.barLength / 15)
    ElseIf (mSpecifier.barLength Mod 5) = 0 Then
        mBarLengthMinutes = 5
        mNumberOfBars = mSpecifier.maxNumberOfBars * (mSpecifier.barLength / 5)
    Else
        mBarLengthMinutes = 1
        mNumberOfBars = mSpecifier.maxNumberOfBars * mSpecifier.barLength
    End If
Case TimePeriodUnits.TimePeriodMonth
    mBarLengthMinutes = 60
    mNumberOfBars = mSpecifier.maxNumberOfBars * 24 * 22 * mSpecifier.barLength
Case TimePeriodUnits.TimePeriodSecond
    mRequestTicks = True
Case TimePeriodUnits.TimePeriodWeek
    mBarLengthMinutes = 60
    mNumberOfBars = mSpecifier.maxNumberOfBars * 24 * 5 * mSpecifier.barLength
Case TimePeriodUnits.TimePeriodYear
    mBarLengthMinutes = 60
    mNumberOfBars = mSpecifier.maxNumberOfBars * 24 * 260 * mSpecifier.barLength
Case TimePeriodUnits.TimePeriodVolume, _
        TimePeriodUnits.TimePeriodTickMovement, _
        TimePeriodUnits.TimePeriodTickVolume
    mRequestTicks = True
End Select


If mSpecifier.barLengthUnits = TimePeriodVolume Or _
    mSpecifier.barLengthUnits = TimePeriodTickMovement Or _
    mSpecifier.barLengthUnits = TimePeriodTickVolume _
Then
    
    If fromDate <> 0 And toDate <> 0 Then
        st = GetSessionTimes(fromDate, _
                                    mInstrument.sessionStartTime, _
                                    mInstrument.sessionEndTime)
        mFromDate = st.startTime
        mToDate = toDate
        mNumberOfBars = MaxLong
    ElseIf fromDate <> 0 Then
        st = GetSessionTimes(fromDate, _
                                    mInstrument.sessionStartTime, _
                                    mInstrument.sessionEndTime)
        mFromDate = st.startTime
        st = GetOffsetSessionTimes(mFromDate, _
                                    1, _
                                    mInstrument.sessionStartTime, _
                                    mInstrument.sessionEndTime)
        mToDate = st.startTime
        mNumberOfBars = MaxLong
    ElseIf toDate <> 0 Then
        mToDate = toDate
        st = GetOffsetSessionTimes(IIf(mToDate = MaxDateValue, Now, mToDate), _
                                    -1, _
                                    mInstrument.sessionStartTime, _
                                    mInstrument.sessionEndTime)
        mFromDate = st.startTime
        mNumberOfBars = MaxLong
    Else
        mToDate = Now
        st = GetOffsetSessionTimes(mToDate, _
                                    -1, _
                                    mInstrument.sessionStartTime, _
                                    mInstrument.sessionEndTime)
        mFromDate = st.startTime
        mNumberOfBars = MaxLong
    End If
Else
    If toDate <> 0 Then
        mToDate = toDate
        
        If fromDate = 0 Then
            ' calculate the earliest possible date for the supplied number of bars (assuming
            ' there are no gaps in the data)
            mFromDate = OffsetBarStartTime(IIf(mToDate = MaxDateValue, Now, mToDate), _
                                        mSpecifier.barLength, _
                                        mSpecifier.barLengthUnits, _
                                        -1 * mSpecifier.maxNumberOfBars, _
                                        mInstrument.sessionStartTime, _
                                        mInstrument.sessionEndTime)
        Else
            mFromDate = fromDate
        End If
    Else
        mFromDate = fromDate
        mToDate = OffsetBarStartTime(fromDate, _
                                    mSpecifier.barLength, _
                                    mSpecifier.barLengthUnits, _
                                    mSpecifier.maxNumberOfBars, _
                                    mInstrument.sessionStartTime, _
                                    mInstrument.sessionEndTime)
    End If
End If
End Sub

Private Function highValue() As Double
highValue = mRs("highprice")
End Function

Private Function lowValue() As Double
lowValue = mRs("lowprice")
End Function

Private Function needMoreTicks() As Boolean
needMoreTicks = builder.Bars.count < mSpecifier.maxNumberOfBars And _
            (mSpecifier.fromDate = 0 Or mSpecifier.toDate = 0)
End Function

Private Function noTicks() As Boolean
noTicks = mRs.EOF
End Function

Private Function openInterest() As Long
openInterest = mRs("openInterest")
End Function

Private Function openValue() As Double
openValue = mRs("openprice")
End Function

Private Sub processBar()
                
builder.notifyValue toSValue(openValue)

If Not mNotifiedFirstVolume Then
    ' send an initial volume of 1, because the first volume notification
    ' is used by the bar studies to set the accum volume at end of
    ' previous bar
    builder.notifyVolume toSValue(1)
    mAccumulatedVolume = 1
    mNotifiedFirstVolume = True
End If

If closeValue >= openValue Then
    builder.notifyValue toSValue(lowValue)
    builder.incrementTickVolume Int(tickVolume / 3) - 2
    builder.notifyVolume toSValue(mAccumulatedVolume + Int(volume / 3))
    
    builder.notifyValue toSValue(lowValue)
    builder.notifyValue toSValue(highValue)
    builder.incrementTickVolume Int(tickVolume / 3) - 2
    builder.notifyVolume toSValue(mAccumulatedVolume + 2 * Int(volume / 3))

    builder.notifyValue toSValue(highValue)
    builder.notifyValue toSValue(closeValue)
    builder.incrementTickVolume tickVolume - 2 * Int(tickVolume / 3) - 2
    builder.notifyVolume toSValue(mAccumulatedVolume + volume)
Else
    builder.notifyValue toSValue(highValue)
    builder.incrementTickVolume Int(tickVolume / 3) - 2
    builder.notifyVolume toSValue(mAccumulatedVolume + Int(volume / 3))
    
    builder.notifyValue toSValue(highValue)
    builder.notifyValue toSValue(lowValue)
    builder.incrementTickVolume Int(tickVolume / 3) - 2
    builder.notifyVolume toSValue(mAccumulatedVolume + 2 * Int(volume / 3))

    builder.notifyValue toSValue(lowValue)
    builder.notifyValue toSValue(closeValue)
    builder.incrementTickVolume tickVolume - 2 * Int(tickVolume / 3) - 2
    builder.notifyVolume toSValue(mAccumulatedVolume + volume)
End If

mAccumulatedVolume = mAccumulatedVolume + volume
End Sub

Private Function processBars() As Bars
If Not mRs.BOF Then
    mRs.MoveLast    ' because the stored procedure returns them in reverse order
End If

Do While Not mRs.BOF
    If filterByTime(mRs("DateTime")) Then processBar
    mRs.MovePrevious
Loop

Set processBars = builder.Bars
End Function

Private Sub processTickData()
Dim val As SValue
Dim tdd As TickDataDecoder
Dim tick As GenericTick
Dim Data() As Byte

If Not filterByTime(mRs("datetime")) Then Exit Sub
    
Data = mRs("data")
Set tdd = CreateTickDecoder(mRs("datetime"), _
                            mInstrument.tickSize, _
                            mRs("baseprice"), _
                            Data, _
                            mRs("version"))
Do While tdd.getNextTick(tick)
    val.timestamp = tick.timestamp
    Select Case tick.tickType
    Case TickTypeBid
        If mSpecifier.barType = BarTypeBid Then
            val.value = tick.price
            builder.notifyValue val
        End If
    Case TickTypeAsk
        If mSpecifier.barType = BarTypeAsk Then
            val.value = tick.price
            builder.notifyValue val
        End If
    Case TickTypeTrade
        If mSpecifier.barType = BarTypeTrade Then
            val.value = tick.price
            builder.notifyValue val
        End If
    Case TickTypeVolume
        If mSpecifier.barType = BarTypeTrade Then
            If Not mNotifiedFirstVolume Then
                val.value = 1
                builder.notifyVolume val
                mNotifiedFirstVolume = True
            End If
            val.value = tick.size + 1
            builder.notifyVolume val
        End If
    Case TickTypeOpenInterest
        If mSpecifier.barType = BarTypeTrade Then
            val.value = tick.size
            builder.notifyOpenInterest val
        End If
    End Select
Loop
End Sub

Private Sub processTicks()
Do While Not mRs.EOF
    processTickData
    mRs.MoveNext
Loop
End Sub

Private Sub setupFetchBarsCommand( _
                        ByVal barType As Long, _
                        ByVal barLength As Long, _
                        ByVal NumberOfBars As Long, _
                        ByVal fromTime As Date, _
                        ByVal toTime As Date)

Dim param As ADODB.Parameter
Dim fromLocal As Date
Dim toLocal As Date

fromTime = gRoundTimeToSecond(fromTime)
toTime = gRoundTimeToSecond(toTime)

Set mCmd = New ADODB.Command
mCmd.CommandType = adCmdStoredProc

mCmd.CommandText = "FetchBarData"

' @InstrumentID
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mInstrument.id)
mCmd.Parameters.Append param

' @BarType
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            barType)
mCmd.Parameters.Append param

' @BarLength
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            barLength)
mCmd.Parameters.Append param

' @NumberRequired
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            NumberOfBars)
mCmd.Parameters.Append param

If fromTime < #1/1/1900# Then fromTime = #1/1/1900# ' don't exceed range of SmallDateTime
If toTime < #1/1/1900# Then toTime = #1/1/1900# ' don't exceed range of SmallDateTime
' @From
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            fromTime)
mCmd.Parameters.Append param

' @To
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            toTime)
mCmd.Parameters.Append param

End Sub

Private Sub setupFetchTicksCommand( _
                ByVal fromTime As Date, _
                ByVal toTime As Date)
Dim param As ADODB.Parameter

fromTime = gRoundTimeToSecond(fromTime)
toTime = gRoundTimeToSecond(toTime)

Set mCmd = New ADODB.Command
mCmd.CommandType = adCmdStoredProc

mCmd.CommandText = "FetchTickData"

' @InstrumentID
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mInstrument.id)
mCmd.Parameters.Append param

' @From
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            fromTime)
mCmd.Parameters.Append param

' @To
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            toTime)
mCmd.Parameters.Append param

End Sub

Private Function tickVolume() As Long
tickVolume = mRs("tickVolume")
End Function

Private Function toSValue( _
                ByVal value As Variant) As SValue
toSValue.value = value
toSValue.timestamp = mRs("datetime")
End Function

Private Function volume() As Long
volume = mRs("Volume")
End Function



