VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickDataWriter"
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

Implements Task

'@===============================================================================
' Events
'@===============================================================================

Event ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
                
Event DataWritten()

Event Error( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
                
'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName               As String = "TradingDO26"
Private Const ModuleName                As String = "TickDataWriter"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private mDB                             As TradingDB
Private WithEvents mConnector           As Connector
Attribute mConnector.VB_VarHelpID = -1
Private WithEvents mConnection          As ADODB.connection
Attribute mConnection.VB_VarHelpID = -1

Private mInstrumentID                   As Long
Private mShortName                      As String

Private mCmd                            As ADODB.Command

Private mAsynchronous                   As Boolean

Private mTaskContext                    As TaskContext

Private mVersion                        As String
Private mPeriodStartTime                As Date
Private mBasePrice                      As Double
Private mTickSize                       As Double
Private mData()                         As Byte

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Terminate()
Debug.Print "TickDataWriter terminated"
gLogger.Log LogLevelHighDetail, taskName & ": terminated"
End Sub

'@===============================================================================
' Task Interface Members
'@===============================================================================

Private Sub Task_cancel()
If Not mCmd Is Nothing Then
    mCmd.cancel
    mConnection.Close
ElseIf Not mConnection Is Nothing Then
    If gIsStateSet(mConnection.State, adStateOpen) Then
        mConnection.Close
    Else
        mConnection.cancel
    End If
End If
If gLogger.IsLoggable(LogLevelDetail) Then gLogger.Log LogLevelDetail, taskName & ": cancelled"
mTaskContext.finish Empty, True
End Sub

Private Sub Task_run()
connectToDB
mTaskContext.Suspend -1
End Sub

Private Property Let Task_TaskContext(ByVal RHS As TWUtilities30.TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get Task_taskName() As String
Task_taskName = mTaskContext.name
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
Dim errorMsg As String
Dim failpoint As Long
On Error GoTo Err

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    If mDB.databaseType = DbMySQL5 And _
        pError.SQLState = "S1000" And _
        pError.NativeError = 1213 _
    Then
        ' a deadlock has been detected - retry the command
        logMessage pMsg:="Deadlock detected while writing tick data - retrying", pProcName:="mConnection_ExecuteComplete"
        Execute
        Exit Sub
    End If
    
    If (mDB.databaseType = DbSQLServer Or _
        mDB.databaseType = DbSQLServer2000 Or _
        mDB.databaseType = DbSQLServer2005 Or _
        mDB.databaseType = DbSQLServer7) And _
        pError.Number = -2147217871 _
    Then
        ' this occurs occasionally while running under Hyper-V Server, when the OS
        ' has for some reason not been able to complete a write to the SQL Server
        ' transaction log within 15 seconds. It's not clear whether the database
        ' write has completed, so we retry it
        logMessage pMsg:="OS timeout while writing to transaction log detected - retrying", pProcName:="mConnection_ExecuteComplete"
        Execute
        Exit Sub
    End If
    
    If pError.Number = 3712 Then
        ' this error is "Operation has been cancelled by the user"
        ' This occurs occasionally when running async, but I've no idea why
        gLogger.Log LogLevelHighDetail, taskName & ": Error 3712 Operation cancelled by user"
        connectToDB
        If Not mAsynchronous Then DoWriteTickData
        Exit Sub
    End If
    
    errorMsg = "TickDataWriter: execution error " & vbCrLf & _
                gGenerateErrorMessage(pError) & _
                gGenerateConnectionErrorMessages(pConnection)
    logMessage pMsg:=taskName & ": " & errorMsg, pProcName:="mConnection_ExecuteComplete"
    
    If gIsStateSet(mConnection.State, adStateOpen) Then
        Debug.Print "TickDataWriter: close connection(error)"
        mConnection.Close
    End If
    If mAsynchronous Then
        mTaskContext.Error pError.Number, ProjectName & "." & ModuleName & ":" & "mConnection_ExecuteComplete" & vbCrLf & errorMsg
    Else
        RaiseEvent Error(pError.Number, errorMsg)
    End If
Else
    If gLogger.IsLoggable(LogLevelMediumDetail) Then gLogger.Log LogLevelMediumDetail, taskName & ": data written"
    
    ' The following statement has been commented out because it causes a failure on Vista
    ' when running synchronously
    'if gisstateset(mConnection.State,adStateOpen) Then mConnection.Close
    
    If mAsynchronous Then
        If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
        mTaskContext.finish 0, False
    Else
        RaiseEvent DataWritten
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=True, pProcedureName:="mConnection_ExecuteComplete", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

'@===============================================================================
' mConnector Event Handlers
'@===============================================================================

Private Sub mConnector_Connected( _
                ByVal pConnection As ADODB.connection)
Set mConnection = pConnection
If mAsynchronous Then DoWriteTickData
End Sub

Private Sub mConnector_ConnectFailed( _
                ByVal errorCode As Long, ByVal errorDesc As String)
If mAsynchronous Then
    mTaskContext.Error errorCode, errorDesc
Else
    RaiseEvent ConnectFailed(errorCode, errorDesc)
End If
End Sub

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub initialise( _
                ByVal db As TradingDB, _
                ByVal instrumentID As Long, _
                ByVal shortName As String)
Set mDB = db
mInstrumentID = instrumentID
mShortName = shortName
End Sub

Friend Sub WriteData( _
                        ByVal Version As String, _
                        ByVal PeriodStartTime As Date, _
                        ByVal BasePrice As Double, _
                        ByVal tickSize As Double, _
                        ByRef Data() As Byte)
Dim failpoint As Long
On Error GoTo Err

If Not mCmd Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":WriteData", _
                "Object has already been used"
End If

mAsynchronous = False

mVersion = Version
mPeriodStartTime = PeriodStartTime
mBasePrice = BasePrice
mTickSize = tickSize
mData = Data

connectToDB
DoWriteTickData
If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="WriteData", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Function WriteDataAsync( _
                        ByVal Version As String, _
                        ByVal PeriodStartTime As Date, _
                        ByVal BasePrice As Double, _
                        ByVal tickSize As Double, _
                        ByRef Data() As Byte) As TaskController
Dim failpoint As Long
On Error GoTo Err

If Not mCmd Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":WriteData", _
                "Object has already been used"
End If

mAsynchronous = True

mVersion = Version
mPeriodStartTime = PeriodStartTime
mBasePrice = BasePrice
mTickSize = tickSize
mData = Data

' we pass the periodstarttime as the cookie for use by the
' TickWriter object when the task completes
Set WriteDataAsync = StartTask(Me, _
                                PriorityNormal, _
                                taskName, _
                                mPeriodStartTime)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="WriteDataAsync", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Function

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub connectToDB()
Dim failpoint As Long
On Error GoTo Err

Debug.Print "TickDataWriter: Connect to DB " & IIf(mAsynchronous, "(async)", "")
If mConnector Is Nothing Then Set mConnector = mDB.getConnector(, IIf(mAsynchronous, ConnectOptionEnum.adAsyncConnect, 0))
mConnector.connect

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="connectToDB", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub DoWriteTickData()

Dim param As ADODB.Parameter

Dim failpoint As Long
On Error GoTo Err

Set mCmd = New ADODB.Command
mCmd.CommandType = adCmdStoredProc

mCmd.CommandText = "WriteTickData"

' @InstrumentID
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mInstrumentID)
mCmd.Parameters.Append param

' @DataVersion
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adVarChar, _
                            ParameterDirectionEnum.adParamInput, _
                            255, _
                            mVersion)
mCmd.Parameters.Append param

' @DateAndTime
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mPeriodStartTime)
mCmd.Parameters.Append param

' @BasePrice
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDouble, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mBasePrice)
mCmd.Parameters.Append param

' @TickSize
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDouble, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mTickSize)
mCmd.Parameters.Append param

' @Data
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adBinary, _
                            ParameterDirectionEnum.adParamInput, _
                            UBound(mData) + 1, _
                            mData)
mCmd.Parameters.Append param

Execute

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="DoWriteTickData", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub Execute()
Dim failpoint As Long
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Execute", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pProjName:=ProjectName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Function taskName() As String
Static lTaskName
If lTaskName = "" Then lTaskName = "TickDataWriter/" & mShortName & "(" & gGetSequenceNumber & ")"
taskName = lTaskName
End Function

