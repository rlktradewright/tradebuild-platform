VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@===============================================================================
' Description
'@===============================================================================
'
'
'@===============================================================================
' Amendment history
'@===============================================================================
'
'
'
'

'@===============================================================================
' Interfaces
'@===============================================================================

Implements ITask

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================


Private Const ModuleName                        As String = "BarWriter"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private mDB                                     As TradingDB
Private WithEvents mConnector                   As Connector
Attribute mConnector.VB_VarHelpID = -1
Private WithEvents mConnection                  As ADODB.connection
Attribute mConnection.VB_VarHelpID = -1

Private mInstrumentID                           As Long
Private mShortName                              As String

Private mCmd                                    As ADODB.Command

Private mAsynchronous                           As Boolean

Private mTaskContext                            As TaskContext

Private mTimestamp                              As Date
Private mBarType                                As Long
Private mBarLength                              As Long
Private mOpenPrice                              As Double
Private mHighPrice                              As Double
Private mLowPrice                               As Double
Private mClosePrice                             As Double
Private mVolume                                 As Long
Private mTickVolume                             As Long
Private mOpenInterest                           As Long

Private mSuppliedTaskName                       As String

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Terminate()
Const ProcName As String = "Class_Terminate"
'Debug.Print "BarWriter terminated"
'If gLogger.IsLoggable(LogLevelHighDetail) Then logMessage pMsg:=taskName & ": terminated", pProcName:=ProcName, pLogLevel:=LogLevelHighDetail
End Sub

'@===============================================================================
' ITask Interface Members
'@===============================================================================

Private Sub ITask_Cancel()
Const ProcName As String = "ITask_Cancel"
On Error GoTo Err

If Not mCmd Is Nothing Then
    mCmd.Cancel
    mConnection.Close
ElseIf Not mConnection Is Nothing Then
    If gIsStateSet(mConnection.State, adStateOpen) Then
        mConnection.Close
    Else
        mConnection.Cancel
    End If
End If
If gLogger.IsLoggable(LogLevelDetail) Then logMessage pMsg:=taskName & ": cancelled", pProcName:=ProcName, pLogLevel:=LogLevelDetail
mTaskContext.Finish Empty, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

connectToDB
mTaskContext.Suspend -1

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
    If mDB.DatabaseType = DbMySQL5 And _
        pError.SQLState = "S1000" And _
        pError.NativeError = 1213 _
    Then
        ' a deadlock has been detected - retry the command
        logMessage pMsg:="Deadlock detected while writing bar - retrying", pProcName:=ProcName
        Execute
        Exit Sub
    End If
    
    If (mDB.DatabaseType = DbSQLServer Or _
        mDB.DatabaseType = DbSQLServer2000 Or _
        mDB.DatabaseType = DbSQLServer2005 Or _
        mDB.DatabaseType = DbSQLServer7) And _
        pError.Number = -2147217871 _
    Then
        ' this occurs occasionally while running under Hyper-V Server, when the OS
        ' has for some reason taken more than 15 seconds to complete a write to the
        ' SQL Server transaction log. We assume that the write has actually
        ' completed, so ignore this error
        logMessage pMsg:="OS timeout while writing to transaction log detected", pProcName:=ProcName
        
        If mAsynchronous Then
            If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
            mTaskContext.Finish 0, False
        End If
       
        Exit Sub
    End If
    
    If pError.Number = 3712 Then
        ' this error is "Operation has been cancelled by the user"
        ' This should no longer happen, but we'll make the check just in case
        logMessage taskName & ": Error 3712 returned from ADO: operation cancelled by user: retrying", pProcName:=ProcName
        connectToDB
        If Not mAsynchronous Then doWriteBar
        Exit Sub
    End If
    
    Dim s As String
    s = "Bar Writer: execution error " & vbCrLf & _
        gGenerateErrorMessage(pError) & _
        gGenerateConnectionErrorMessages(pConnection)
    logMessage pMsg:=taskName & ": " & s, pProcName:=ProcName
    
    If gIsStateSet(mConnection.State, adStateOpen) Then
        'Debug.Print "BarWriter: close connection(error)"
        mConnection.Close
    End If
    
    If mAsynchronous Then
        mTaskContext.Error pError.Number, s, ProjectName & "." & ModuleName & ":" & ProcName
    Else
        Err.Raise pError.Number, , s
    End If
Else
    If gLogger.IsLoggable(LogLevelMediumDetail) Then logMessage pMsg:=taskName & ": bar written", pProcName:=ProcName, pLogLevel:=LogLevelMediumDetail
    
    ' The following statement has been commented out because it causes a failure on Vista
    ' when running synchronously
    'if gisstateset(mConnection.State,adStateOpen) Then mConnection.Close
    
    If mAsynchronous Then
        If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
        mTaskContext.Finish 0, False
    End If
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
If mAsynchronous Then doWriteBar

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
    mTaskContext.Error errorCode, _
                        "Connection error " & vbCrLf & errorDesc
Else
    Err.Raise errorCode, , "Connection error " & vbCrLf & errorDesc
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal instrumentID As Long, _
                ByVal ShortName As String)
Const ProcName As String = "Initialise"
Set mDB = db
mInstrumentID = instrumentID
mShortName = ShortName
End Sub

Friend Sub WriteBar( _
                ByVal timestamp As Date, _
                ByVal barType As Long, _
                ByVal barLength As Long, _
                ByVal openPrice As Double, _
                ByVal highPrice As Double, _
                ByVal lowPrice As Double, _
                ByVal closePrice As Double, _
                ByVal volume As Long, _
                ByVal tickVolume As Long, _
                ByVal openInterest As Long)
Const ProcName As String = "WriteBar"
On Error GoTo Err

mAsynchronous = False
mTimestamp = timestamp
mBarType = barType
mBarLength = barLength
mOpenPrice = openPrice
mHighPrice = highPrice
mLowPrice = lowPrice
mClosePrice = closePrice
mVolume = volume
mTickVolume = tickVolume
mOpenInterest = openInterest

If gLogger.IsLoggable(LogLevelMediumDetail) Then logMessage pMsg:=taskName & ": about to write bar", pProcName:=ProcName, pLogLevel:=LogLevelMediumDetail

connectToDB

doWriteBar

If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function WriteBarAsync( _
                ByVal timestamp As Date, _
                ByVal barType As Long, _
                ByVal barLength As Long, _
                ByVal openPrice As Double, _
                ByVal highPrice As Double, _
                ByVal lowPrice As Double, _
                ByVal closePrice As Double, _
                ByVal volume As Long, _
                ByVal tickVolume As Long, _
                ByVal openInterest As Long, _
                ByVal cookie As Variant, _
                ByVal pTaskName As String) As TaskController
Const ProcName As String = "WriteBarAsync"
On Error GoTo Err

mAsynchronous = True
mTimestamp = timestamp
mBarType = barType
mBarLength = barLength
mOpenPrice = openPrice
mHighPrice = highPrice
mLowPrice = lowPrice
mClosePrice = closePrice
mVolume = volume
mTickVolume = tickVolume
mOpenInterest = openInterest

mSuppliedTaskName = pTaskName

If gLogger.IsLoggable(LogLevelMediumDetail) Then logMessage pMsg:=taskName & ": about to write bar (async)", pProcName:=ProcName, pLogLevel:=LogLevelMediumDetail

Set WriteBarAsync = StartTask(Me, _
                            PriorityNormal, _
                            taskName, _
                            cookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub connectToDB()
Const ProcName As String = "connectToDB"
On Error GoTo Err

'Debug.Print "BarWriter: Connect to DB " & IIf(mAsynchronous, "(async)", "")
Set mCmd = Nothing
If mConnector Is Nothing Then Set mConnector = mDB.GetConnector
If mAsynchronous Then
    mConnector.ConnectAsync
Else
    mConnector.Connect
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doWriteBar()
Const ProcName As String = "doWriteBar"
On Error GoTo Err

If Not mCmd Is Nothing Then logMessage "mCmd is already set", ProcName
Set mCmd = New ADODB.Command
mCmd.CommandType = adCmdStoredProc

mCmd.CommandText = "WriteBarData"

Dim param As ADODB.Parameter

' @InstrumentID
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mInstrumentID)
mCmd.Parameters.Append param

' @BarType
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mBarType)
mCmd.Parameters.Append param

' @BarLength
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mBarLength)
mCmd.Parameters.Append param

' @DateAndTime
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mTimestamp)
mCmd.Parameters.Append param

' @OpenPrice
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDouble, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mOpenPrice)
mCmd.Parameters.Append param

' @HighPrice
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDouble, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mHighPrice)
mCmd.Parameters.Append param

' @LowPrice
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDouble, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mLowPrice)
mCmd.Parameters.Append param

' @ClosePrice
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDouble, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mClosePrice)
mCmd.Parameters.Append param

' @Volume
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mVolume)
mCmd.Parameters.Append param

' @TickVolume
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mTickVolume)
mCmd.Parameters.Append param

' @OpenInterest
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mOpenInterest)
mCmd.Parameters.Append param

Execute

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Execute()
Const ProcName As String = "Execute"
On Error GoTo Err

Set mCmd.ActiveConnection = mConnection
If mAsynchronous Then
    mCmd.Execute , _
                , _
                ExecuteOptionEnum.adAsyncExecute + ExecuteOptionEnum.adExecuteNoRecords
Else
    mCmd.Execute , _
                , _
                ExecuteOptionEnum.adExecuteNoRecords
End If

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

Private Function taskName() As String
Const ProcName As String = "taskName"

Static lTaskName
If lTaskName = "" Then
    If mSuppliedTaskName <> "" Then
        lTaskName = mSuppliedTaskName
    Else
        lTaskName = "BarWriter/" & mShortName & _
                                "(" & mBarLength & ")" & _
                                FormatTimestamp(mTimestamp, TimestampNoMillisecs) & _
                                "(" & gGetSequenceNumber & ")"
    End If
End If
taskName = lTaskName
End Function






