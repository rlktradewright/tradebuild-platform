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

Private Sub Task_Cancel()

Const ProcName As String = "Task_Cancel"
Dim failpoint As String
On Error GoTo Err

If Not mRs Is Nothing Then
    If Not gIsStateSet(mRs.State, adStateOpen) Then mRs.Cancel
End If

If Not mCmd Is Nothing Then mCmd.Cancel

If Not mConnection Is Nothing Then
    If gIsStateSet(mConnection.State, adStateOpen) Then
        mConnection.Close
    Else
        mConnection.Cancel
    End If
End If

mTaskContext.Finish Empty, True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub Task_run()
Static gotFirstTicks As Boolean

Const ProcName As String = "Task_run"
Dim failpoint As String
On Error GoTo Err

Select Case mAsyncState
Case Started

    gLogger.Log LogLevelDetail, "BarFetcher started for " & mInstrument.ShortName
    
    If mTaskContext.CancelPending Then
        mTaskContext.Finish Empty, True
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
    
    mTaskContext.Suspend -1
    
Case Connecting
    logMessage pMsg:=mTaskContext.Name & ": Task_run invoked in Connecting state", pProcName:="Task_run"
    gLogger.Log LogLevelNormal, "Runnable tasks--------------------------------------------"
    gLogger.Log LogLevelNormal, TaskManager.GetRunnableTaskSummary
    gLogger.Log LogLevelNormal, "Suspended tasks-------------------------------------------"
    gLogger.Log LogLevelNormal, TaskManager.GetSuspendedTaskSummary
    mTaskContext.Error ErrorCodes.ErrIllegalStateException, _
                        "Task_run invoked in Connecting state"
Case Fetching
    logMessage pMsg:=mTaskContext.Name & ": Task_run invoked in Fetching state", pProcName:="Task_run"
    gLogger.Log LogLevelNormal, "Runnable tasks--------------------------------------------"
    gLogger.Log LogLevelNormal, TaskManager.GetRunnableTaskSummary
    gLogger.Log LogLevelNormal, "Suspended tasks-------------------------------------------"
    gLogger.Log LogLevelNormal, TaskManager.GetSuspendedTaskSummary
    mTaskContext.Error ErrorCodes.ErrIllegalStateException, _
                        "Task_run invoked in Fetching state"
Case ProcessingBars
    
    If mTaskContext.CancelPending Then
        mTaskContext.Finish Empty, True
        Exit Sub
    End If
    
    If Not mRs.BOF Then
        If filterByTime(mRs("DateTime")) Then processBar
        mRs.MovePrevious
    Else
        Debug.Print "BarFetcher: close connection"
        If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
        gLogger.Log LogLevelDetail, "BarFetcher fetched " & builder.Bars.Count & " bars for " & mInstrument.ShortName
        mTaskContext.Finish builder.Bars, False
    End If

Case ProcessingTicks
    
    If mTaskContext.CancelPending Then
        mTaskContext.Finish Empty, True
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
                mTaskContext.Suspend -1    ' wait for more ticks to be fetched
            Else
                gLogger.Log LogLevelDetail, "BarFetcher fetched " & builder.Bars.Count & " bars for " & mInstrument.ShortName
                Debug.Print "BarFetcher: close connection"
                If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
                mTaskContext.Finish mFetchBarsBuilder.Bars, False
            End If
        Else
            If mSpecifier.fromDate = 0 Then
                Set mAppendBarsTaskController = builder.AppendBarsAsync(mFetchBarsBuilder.Bars)
                Set mFetchBarsBuilder = builder
            Else
                Set mAppendBarsTaskController = mFetchBarsBuilder.AppendBarsAsync(builder.Bars)
            End If
            mTaskContext.Suspend -1    ' wait for the append to Finish
        End If
    
    End If
Case AppendingBars
    mTaskContext.Error ErrorCodes.ErrIllegalStateException, _
                        "Task_run invoked in AppendingBars state"
End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Let Task_TaskContext(ByVal value As TWUtilities30.TaskContext)
Set mTaskContext = value
End Property

Private Property Get Task_taskName() As String
Task_taskName = mTaskContext.Name
End Property

'@===============================================================================
' mAppendBarsTaskController Event Handlers
'@===============================================================================

Private Sub mAppendBarsTaskController_Completed(ev As TWUtilities30.TaskCompletionEvent)
Const ProcName As String = "mAppendBarsTaskController_Completed"
Dim failpoint As String
On Error GoTo Err

If ev.cancelled Then
    mTaskContext.Finish Empty, True
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
        If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
        mTaskContext.Finish mFetchBarsBuilder.Bars, False
    End If
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
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
Dim s As String


On Error GoTo Err

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
        If gIsStateSet(mConnection.State, adStateOpen) Then
            Debug.Print "BarFetcher: close connection(error)"
            mConnection.Close
        End If
        mTaskContext.Error pError.Number, _
                            ProjectName & "." & ModuleName & ":" & "mConnection_ExecuteComplete" & vbCrLf & s
    Else
        Err.Raise pError.Number, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                s
    End If
    Exit Sub
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@===============================================================================
' mConnector Event Handlers
'@===============================================================================

Private Sub mConnector_Connected( _
                ByVal pConnection As ADODB.connection)
Const ProcName As String = "mConnector_Connected"


On Error GoTo Err

Set mConnection = pConnection
If mAsynchronous Then
    Execute
    mAsyncState = Fetching
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mConnector_ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
Const ProcName As String = "mConnector_ConnectFailed"


On Error GoTo Err

If mAsynchronous Then
    gLogger.Log LogLevelDetail, "BarFetcher connection error for " & mInstrument.ShortName
    mTaskContext.Error errorCode, _
                        ProjectName & "." & ModuleName & ":" & "mConnector_ConnectFailed" & vbCrLf & errorDesc
Else
    Err.Raise errorCode, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                errorDesc
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@===============================================================================
' mRs Event Handlers
'@===============================================================================

Private Sub mRs_FetchComplete( _
                ByVal pError As ADODB.Error, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pRecordset As ADODB.Recordset)
Dim s As String

Const ProcName As String = "mRs_FetchComplete"
Dim failpoint As String
On Error GoTo Err

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    s = "Fetch error " & vbCrLf & _
        gGenerateErrorMessage(pError) & _
        gGenerateConnectionErrorMessages(mConnection)
    
    Debug.Print "BarFetcher: close connection(error)"
    If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
    If mAsynchronous Then
        mTaskContext.Error pError.Number, _
                            ProjectName & "." & ModuleName & ":" & "mRs_FetchComplete" & vbCrLf & s
    Else
        Err.Raise pError.Number, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
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
            gLogger.Log LogLevelDetail, "BarFetcher fetched " & builder.Bars.Count & " bars for " & mInstrument.ShortName
            Debug.Print "BarFetcher: close connection"
            If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
            mTaskContext.Finish mFetchBarsBuilder.Bars, False
        Else
            mAsyncState = ProcessingTicks
        End If
    Else
        If Not mRs.BOF Then
            mRs.MoveLast    ' because the stored procedure returns them in reverse order
        End If
        mAsyncState = ProcessingBars
    End If
    mTaskContext.Continue   ' start calling Task_run again
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
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

Friend Function FetchBars( _
                ByVal specifier As BarDataSpecifier) As Bars
Const ProcName As String = "FetchBars"

On Error GoTo Err

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
                
    Set FetchBars = processBars
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
                builder.AppendBars mFetchBarsBuilder.Bars
                Set mFetchBarsBuilder = builder
            Else
                mFetchBarsBuilder.AppendBars builder.Bars
            End If
        Loop
        
    End If
    Set FetchBars = mFetchBarsBuilder.Bars
End If

Debug.Print "BarFetcher: close connection"
If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Friend Function FetchbarsAsync( _
                ByVal specifier As BarDataSpecifier, _
                ByVal cookie As Variant, _
                ByVal taskName As String) As TaskController
Const ProcName As String = "FetchbarsAsync"


On Error GoTo Err

mAsyncState = Started

mAsynchronous = True
Set mSpecifier = specifier

createNewCustomSession

Set mFetchBarsBuilder = builder

If taskName = "" Then taskName = "BarFetcher/" & mInstrument.ShortName & "(" & mSpecifier.barTimePeriod.ToString & ")" & FormatTimestamp(mFromDate, TimestampNoMillisecs) & "-" & FormatTimestamp(mToDate, TimestampNoMillisecs) & "(" & gGetSequenceNumber & ")"

Set FetchbarsAsync = StartTask(Me, _
                            PriorityNormal, _
                            taskName, _
                            cookie)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal pInstrument As instrument)
Const ProcName As String = "Initialise"


On Error GoTo Err

Set mDB = db
Set mInstrument = pInstrument
createNewSession

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function builder() As BarsBuilder
Const ProcName As String = "builder"


On Error GoTo Err

If mBuilder Is Nothing Then Set mBuilder = CreateBarsBuilder(mSpecifier.barTimePeriod, _
                                                            mSession, _
                                                            mInstrument.TickSize, _
                                                            IIf(mSpecifier.maxNumberOfBars = &H7FFFFFFF, 0, mSpecifier.maxNumberOfBars))
Set builder = mBuilder

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function closeValue() As Double
Const ProcName As String = "closeValue"


On Error GoTo Err

closeValue = mRs("closeprice")

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub connectToDB()
Const ProcName As String = "connectToDB"


On Error GoTo Err

Debug.Print "BarFetcher: Connect to DB"
If mConnector Is Nothing Then Set mConnector = mDB.GetConnector(, IIf(mAsynchronous, ConnectOptionEnum.adAsyncConnect, 0))
mConnector.connect
If mAsynchronous Then mAsyncState = Connecting

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub createNewCustomSession()
Const ProcName As String = "createNewCustomSession"


On Error GoTo Err

If mSpecifier.customSessionEndTime <> 0 Or mSpecifier.customSessionStartTime <> 0 Then
    Set mCustomSessionBuilder = New SessionBuilder
    mCustomSessionBuilder.SessionStartTime = mSpecifier.customSessionStartTime
    mCustomSessionBuilder.SessionEndTime = mSpecifier.customSessionEndTime
    mCustomSessionBuilder.TimeZone = GetTimeZone(mInstrument.TimeZoneName)
    Set mCustomSession = mCustomSessionBuilder.Session
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub createNewSession()
Const ProcName As String = "createNewSession"


On Error GoTo Err

Set mSessionBuilder = New SessionBuilder
mSessionBuilder.SessionStartTime = mInstrument.SessionStartTime
mSessionBuilder.SessionEndTime = mInstrument.SessionEndTime
mSessionBuilder.TimeZone = GetTimeZone(mInstrument.TimeZoneName)
Set mSession = mSessionBuilder.Session

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub doFetchBars( _
                        ByVal barType As Long, _
                        ByVal barLength As Long, _
                        ByVal NumberOfBars As Long, _
                        ByVal fromTime As Date, _
                        ByVal toTime As Date)
Const ProcName As String = "doFetchBars"



On Error GoTo Err

setupFetchBarsCommand barType, barLength, NumberOfBars, fromTime, toTime

If Not mAsynchronous Then Execute

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub doFetchTicks( _
                ByVal fromTime As Date, _
                ByVal toTime As Date)
Const ProcName As String = "doFetchTicks"


On Error GoTo Err

setupFetchTicksCommand fromTime, toTime

If Not mAsynchronous Then Execute

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub Execute()
Const ProcName As String = "Execute"


On Error GoTo Err

gLogger.Log LogLevelDetail, "BarFetcher initiating fetch for " & mInstrument.ShortName

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function filterByTime( _
                ByVal timestamp As Date) As Boolean
Const ProcName As String = "filterByTime"


On Error GoTo Err

mSessionBuilder.SetSessionCurrentTime timestamp
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub generateNextRequest()
Const ProcName As String = "generateNextRequest"


On Error GoTo Err

If mSpecifier.fromDate = 0 Then
    generateRequestDetails 0, mFromDate
Else
    generateRequestDetails mToDate, 0
End If

Set mBuilder = Nothing
mNotifiedFirstVolume = False
createNewSession

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub generateRequestDetails( _
                        ByVal fromDate As Date, _
                        ByVal toDate As Date)
Const ProcName As String = "generateRequestDetails"
Dim st As SessionTimes


On Error GoTo Err

Select Case mSpecifier.barTimePeriod.Units
Case TimePeriodUnits.TimePeriodDay
    mBarLengthMinutes = 60
    mNumberOfBars = mSpecifier.maxNumberOfBars * 24 * mSpecifier.barTimePeriod.Length
Case TimePeriodUnits.TimePeriodHour
    mBarLengthMinutes = 60
    mNumberOfBars = mSpecifier.maxNumberOfBars * mSpecifier.barTimePeriod.Length
Case TimePeriodUnits.TimePeriodMinute
    If (mSpecifier.barTimePeriod.Length Mod 60) = 0 Then
        mBarLengthMinutes = 60
        mNumberOfBars = mSpecifier.maxNumberOfBars * (mSpecifier.barTimePeriod.Length / 60)
    ElseIf (mSpecifier.barTimePeriod.Length Mod 15) = 0 Then
        mBarLengthMinutes = 15
        mNumberOfBars = mSpecifier.maxNumberOfBars * (mSpecifier.barTimePeriod.Length / 15)
    ElseIf (mSpecifier.barTimePeriod.Length Mod 5) = 0 Then
        mBarLengthMinutes = 5
        mNumberOfBars = mSpecifier.maxNumberOfBars * (mSpecifier.barTimePeriod.Length / 5)
    Else
        mBarLengthMinutes = 1
        mNumberOfBars = mSpecifier.maxNumberOfBars * mSpecifier.barTimePeriod.Length
    End If
Case TimePeriodUnits.TimePeriodMonth
    mBarLengthMinutes = 60
    mNumberOfBars = mSpecifier.maxNumberOfBars * 24 * 22 * mSpecifier.barTimePeriod.Length
Case TimePeriodUnits.TimePeriodSecond
    mRequestTicks = True
Case TimePeriodUnits.TimePeriodWeek
    mBarLengthMinutes = 60
    mNumberOfBars = mSpecifier.maxNumberOfBars * 24 * 5 * mSpecifier.barTimePeriod.Length
Case TimePeriodUnits.TimePeriodYear
    mBarLengthMinutes = 60
    mNumberOfBars = mSpecifier.maxNumberOfBars * 24 * 260 * mSpecifier.barTimePeriod.Length
Case TimePeriodUnits.TimePeriodVolume, _
        TimePeriodUnits.TimePeriodTickMovement, _
        TimePeriodUnits.TimePeriodTickVolume
    mRequestTicks = True
End Select


If mSpecifier.barTimePeriod.Units = TimePeriodVolume Or _
    mSpecifier.barTimePeriod.Units = TimePeriodTickMovement Or _
    mSpecifier.barTimePeriod.Units = TimePeriodTickVolume _
Then
    
    If fromDate <> 0 And toDate <> 0 Then
        st = GetSessionTimes(fromDate, _
                                    mInstrument.SessionStartTime, _
                                    mInstrument.SessionEndTime)
        mFromDate = st.startTime
        mToDate = toDate
        mNumberOfBars = MaxLong
    ElseIf fromDate <> 0 Then
        st = GetSessionTimes(fromDate, _
                                    mInstrument.SessionStartTime, _
                                    mInstrument.SessionEndTime)
        mFromDate = st.startTime
        st = GetOffsetSessionTimes(mFromDate, _
                                    1, _
                                    mInstrument.SessionStartTime, _
                                    mInstrument.SessionEndTime)
        mToDate = st.startTime
        mNumberOfBars = MaxLong
    ElseIf toDate <> 0 Then
        mToDate = toDate
        st = GetOffsetSessionTimes(IIf(mToDate = MaxDateValue, Now, mToDate), _
                                    -1, _
                                    mInstrument.SessionStartTime, _
                                    mInstrument.SessionEndTime)
        mFromDate = st.startTime
        mNumberOfBars = MaxLong
    Else
        mToDate = Now
        st = GetOffsetSessionTimes(mToDate, _
                                    -1, _
                                    mInstrument.SessionStartTime, _
                                    mInstrument.SessionEndTime)
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
                                        mSpecifier.barTimePeriod, _
                                        -1 * mSpecifier.maxNumberOfBars, _
                                        mInstrument.SessionStartTime, _
                                        mInstrument.SessionEndTime)
        Else
            mFromDate = fromDate
        End If
    Else
        mFromDate = fromDate
        mToDate = OffsetBarStartTime(fromDate, _
                                    mSpecifier.barTimePeriod, _
                                    mSpecifier.maxNumberOfBars, _
                                    mInstrument.SessionStartTime, _
                                    mInstrument.SessionEndTime)
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function highValue() As Double
Const ProcName As String = "highValue"


On Error GoTo Err

highValue = mRs("highprice")

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "logMessage"
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pProjName:=ProjectName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Function lowValue() As Double
Const ProcName As String = "lowValue"


On Error GoTo Err

lowValue = mRs("lowprice")

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function needMoreTicks() As Boolean
Const ProcName As String = "needMoreTicks"


On Error GoTo Err

needMoreTicks = builder.Bars.Count < mSpecifier.maxNumberOfBars And _
            (mSpecifier.fromDate = 0 Or mSpecifier.toDate = 0)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function noTicks() As Boolean
Const ProcName As String = "noTicks"


On Error GoTo Err

noTicks = mRs.EOF

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function openInterest() As Long
Const ProcName As String = "openInterest"


On Error GoTo Err

openInterest = mRs("openInterest")

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function openValue() As Double
Const ProcName As String = "openValue"


On Error GoTo Err

openValue = mRs("openprice")

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub processBar()
Const ProcName As String = "processBar"
                

On Error GoTo Err

builder.NotifyValue toSValue(openValue)

If Not mNotifiedFirstVolume Then
    ' send an initial volume of 1, because the first volume notification
    ' is used by the bar studies to set the accum volume at end of
    ' previous bar
    builder.NotifyVolume toSValue(1)
    mAccumulatedVolume = 1
    mNotifiedFirstVolume = True
End If

If closeValue >= openValue Then
    builder.NotifyValue toSValue(lowValue)
    builder.IncrementTickVolume Int(tickVolume / 3) - 2
    builder.NotifyVolume toSValue(mAccumulatedVolume + Int(volume / 3))
    
    builder.NotifyValue toSValue(lowValue)
    builder.NotifyValue toSValue(highValue)
    builder.IncrementTickVolume Int(tickVolume / 3) - 2
    builder.NotifyVolume toSValue(mAccumulatedVolume + 2 * Int(volume / 3))

    builder.NotifyValue toSValue(highValue)
    builder.NotifyValue toSValue(closeValue)
    builder.IncrementTickVolume tickVolume - 2 * Int(tickVolume / 3) - 2
    builder.NotifyVolume toSValue(mAccumulatedVolume + volume)
Else
    builder.NotifyValue toSValue(highValue)
    builder.IncrementTickVolume Int(tickVolume / 3) - 2
    builder.NotifyVolume toSValue(mAccumulatedVolume + Int(volume / 3))
    
    builder.NotifyValue toSValue(highValue)
    builder.NotifyValue toSValue(lowValue)
    builder.IncrementTickVolume Int(tickVolume / 3) - 2
    builder.NotifyVolume toSValue(mAccumulatedVolume + 2 * Int(volume / 3))

    builder.NotifyValue toSValue(lowValue)
    builder.NotifyValue toSValue(closeValue)
    builder.IncrementTickVolume tickVolume - 2 * Int(tickVolume / 3) - 2
    builder.NotifyVolume toSValue(mAccumulatedVolume + volume)
End If

mAccumulatedVolume = mAccumulatedVolume + volume

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function processBars() As Bars
Const ProcName As String = "processBars"


On Error GoTo Err

If Not mRs.BOF Then
    mRs.MoveLast    ' because the stored procedure returns them in reverse order
End If

Do While Not mRs.BOF
    If filterByTime(mRs("DateTime")) Then processBar
    mRs.MovePrevious
Loop

Set processBars = builder.Bars

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub processTickData()
Const ProcName As String = "processTickData"
Dim val As SValue
Dim tdd As TickDataDecoder
Dim tick As GenericTick
Dim Data() As Byte


On Error GoTo Err

If Not filterByTime(mRs("datetime")) Then Exit Sub
    
Data = mRs("data")
Set tdd = CreateTickDecoder(mRs("datetime"), _
                            mInstrument.TickSize, _
                            mRs("baseprice"), _
                            Data, _
                            mRs("version"))
Do While tdd.GetNextTick(tick)
    val.timestamp = tick.timestamp
    Select Case tick.tickType
    Case TickTypeBid
        If mSpecifier.barType = BarTypeBid Then
            val.value = tick.price
            builder.NotifyValue val
        End If
    Case TickTypeAsk
        If mSpecifier.barType = BarTypeAsk Then
            val.value = tick.price
            builder.NotifyValue val
        End If
    Case TickTypeTrade
        If mSpecifier.barType = BarTypeTrade Then
            val.value = tick.price
            builder.NotifyValue val
        End If
    Case TickTypeVolume
        If mSpecifier.barType = BarTypeTrade Then
            If Not mNotifiedFirstVolume Then
                val.value = 1
                builder.NotifyVolume val
                mNotifiedFirstVolume = True
            End If
            val.value = tick.size + 1
            builder.NotifyVolume val
        End If
    Case TickTypeOpenInterest
        If mSpecifier.barType = BarTypeTrade Then
            val.value = tick.size
            builder.NotifyOpenInterest val
        End If
    End Select
Loop

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setupFetchBarsCommand( _
                        ByVal barType As Long, _
                        ByVal barLength As Long, _
                        ByVal NumberOfBars As Long, _
                        ByVal fromTime As Date, _
                        ByVal toTime As Date)
Const ProcName As String = "setupFetchBarsCommand"


Dim param As ADODB.Parameter
Dim fromLocal As Date
Dim toLocal As Date


On Error GoTo Err

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
                            mInstrument.Id)
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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub setupFetchTicksCommand( _
                ByVal fromTime As Date, _
                ByVal toTime As Date)
Const ProcName As String = "setupFetchTicksCommand"
Dim param As ADODB.Parameter


On Error GoTo Err

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
                            mInstrument.Id)
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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Function tickVolume() As Long
Const ProcName As String = "tickVolume"


On Error GoTo Err

tickVolume = mRs("tickVolume")

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function toSValue( _
                ByVal value As Variant) As SValue
Const ProcName As String = "toSValue"


On Error GoTo Err

toSValue.value = value
toSValue.timestamp = mRs("datetime")

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function volume() As Long
Const ProcName As String = "volume"


On Error GoTo Err

volume = mRs("Volume")

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function



