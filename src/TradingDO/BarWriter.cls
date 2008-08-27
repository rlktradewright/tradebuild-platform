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

Implements Task

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName                       As String = "TradingDO26"
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
Debug.Print "BarWriter terminated"
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
If gLogger.isLoggable(LogLevelDetail) Then gLogger.Log LogLevelDetail, taskName & ": cancelled"
mTaskContext.finish Empty, True
End Sub

Private Sub Task_run()
connectToDB
mTaskContext.suspend -1
End Sub

Private Property Let Task_TaskContext(ByVal value As TWUtilities30.TaskContext)
Set mTaskContext = value
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
Dim s As String
If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    If mDB.databaseType = DbMySQL5 And _
        pError.SQLState = "S1000" And _
        pError.NativeError = 1213 _
    Then
        ' a deadlock has been detected - retry the command
        If gLogger.isLoggable(LogLevelDetail) Then gLogger.Log LogLevelDetail, taskName & ": deadlock detected - retrying"
        Execute
        Exit Sub
    End If
    
    If pError.Number = 3712 Then
        ' this error is "Operation has been cancelled by the user"
        ' This occurs occasionally when running async, but I've no idea why
        gLogger.Log LogLevelHighDetail, taskName & ": Error 3712 Operation cancelled by user"
        connectToDB
        If Not mAsynchronous Then doWriteBar
        Exit Sub
    End If
    
    s = "Bar Writer: execution error " & vbCrLf & _
        gGenerateErrorMessage(pError) & _
        gGenerateConnectionErrorMessages(pConnection)
    gLogger.Log LogLevelNormal, taskName & ": " & s
    
    If gIsStateSet(mConnection.State, adStateOpen) Then
        Debug.Print "BarWriter: close connection(error)"
        mConnection.Close
    End If
    
    If mAsynchronous Then
        mTaskContext.Error pError.Number, ProjectName & "." & ModuleName & ":" & "mConnection_ExecuteComplete" & vbCrLf & s
    Else
        Err.Raise pError.Number, _
                    ProjectName & "." & ModuleName & ":" & "mConnection_ExecuteComplete", _
                    s
    End If
Else
    If gLogger.isLoggable(LogLevelMediumDetail) Then gLogger.Log LogLevelMediumDetail, taskName & ": data written"
    
    ' The following statement has been commented out because it causes a failure on Vista
    ' when running synchronously
    'if gisstateset(mConnection.State,adStateOpen) Then mConnection.Close
    
    If mAsynchronous Then
        If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
        mTaskContext.finish 0, False
    End If
End If
End Sub

'@===============================================================================
' mConnector Event Handlers
'@===============================================================================

Private Sub mConnector_Connected( _
                ByVal pConnection As ADODB.connection)
Set mConnection = pConnection
If mAsynchronous Then doWriteBar
End Sub

Private Sub mConnector_ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
If mAsynchronous Then
    mTaskContext.Error errorCode, _
                        "Connection error " & vbCrLf & errorDesc
Else
    Err.Raise errorCode, _
                ProjectName & "." & ModuleName & ":" & "mConnector_ConnectFailed", _
                "Connection error " & vbCrLf & errorDesc
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

connectToDB

doWriteBar

If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close

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

Set WriteBarAsync = StartTask(Me, _
                            PriorityNormal, _
                            taskName, _
                            cookie)

End Function

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub connectToDB()
Debug.Print "BarWriter: Connect to DB " & IIf(mAsynchronous, "(async)", "")
If mConnector Is Nothing Then Set mConnector = mDB.getConnector(, IIf(mAsynchronous, ConnectOptionEnum.adAsyncConnect, 0))
mConnector.connect
End Sub

Private Sub doWriteBar()

Dim param As ADODB.Parameter

Set mCmd = New ADODB.Command
mCmd.CommandType = adCmdStoredProc

mCmd.CommandText = "WriteBarData"

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
End Sub

Private Sub Execute()
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
End Sub

Private Function taskName() As String
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






