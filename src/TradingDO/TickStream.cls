VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

Event DataAvailable()

Event Error( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
                
'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "TickStream"

'@================================================================================
' Member variables
'@================================================================================

Private mDB                                 As TradingDB
Private WithEvents mConnector               As Connector
Attribute mConnector.VB_VarHelpID = -1
Private WithEvents mConnection              As ADODB.connection
Attribute mConnection.VB_VarHelpID = -1
Private WithEvents mRs                      As ADODB.Recordset
Attribute mRs.VB_VarHelpID = -1

Private mAsynchronous                       As Boolean

Private mInstrumentID                       As Long

Private mFromDate                           As Date
Private mToDate                             As Date

Private mWaitingForData                     As Boolean
Private mFetchComplete                      As Boolean

Private mEos                                As Boolean

Private mDecoder                            As TickDataDecoder

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

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
Dim failpoint As String
On Error GoTo Err

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent Error(pError.Number, _
                    ModuleName & ": fetch error " & vbCrLf & _
                    gGenerateErrorMessage(pError) & _
                    gGenerateConnectionErrorMessages(mConnection))
    If mAsynchronous Then
        If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
    End If
Else
    Set mRs.ActiveConnection = Nothing
    If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
    
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
Dim failpoint As String
On Error GoTo Err

Set mConnection = pConnection
If mAsynchronous Then doFetchTicks

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mConnector_ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
RaiseEvent Error(errorCode, errorDesc)
End Sub

'@===============================================================================
' mRs Event Handlers
'@===============================================================================

Private Sub mRs_FetchComplete( _
                ByVal pError As ADODB.Error, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pRecordset As ADODB.Recordset)
Const ProcName As String = "mRs_FetchComplete"
Dim failpoint As String
On Error GoTo Err

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent Error(pError.Number, _
                    ModuleName & ": fetch error " & vbCrLf & _
                    gGenerateErrorMessage(pError) & _
                    gGenerateConnectionErrorMessages(mConnection))
    Exit Sub
End If

mFetchComplete = True
Debug.Print "Fetch complete"

If Not mAsynchronous Then
    RaiseEvent DataAvailable
ElseIf mWaitingForData Then
    mWaitingForData = False
    RaiseEvent DataAvailable
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mRs_FetchProgress( _
                ByVal Progress As Long, _
                ByVal MaxProgress As Long, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pRecordset As ADODB.Recordset)

Const ProcName As String = "mRs_FetchProgress"
Dim failpoint As String
On Error GoTo Err

Debug.Print "Fetch progress=" & Progress & "; max=" & MaxProgress
If Not mAsynchronous Then Exit Sub

If mWaitingForData Then
    mWaitingForData = False
    RaiseEvent DataAvailable
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Eos() As Boolean
Eos = mEos
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Cancel()
Const ProcName As String = "Cancel"
Dim failpoint As String
On Error GoTo Err

If Not mRs Is Nothing Then
    If gIsStateSet(mRs.State, adStateOpen) Then
        mRs.Close
    Else
        mRs.Cancel
    End If
    Set mRs = Nothing
End If

If Not mConnection Is Nothing Then
    If gIsStateSet(mConnection.State, adStateOpen) Then
        mConnection.Close
    Else
        mConnection.Cancel
    End If
    Set mConnection = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
    
End Sub

Public Function GetNextTick( _
                ByRef tick As GenericTick) As Boolean

Const ProcName As String = "GetNextTick"
Dim failpoint As String
On Error GoTo Err

If mRs Is Nothing Then
    If Not mAsynchronous Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Must call StartStream before calling GetNextTick"
    End If
    
    Debug.Print "GetNextTick: no recordset"
    mWaitingForData = True
    Exit Function
End If

If mDecoder Is Nothing Then
    If Not getNextDecoder Then Exit Function
End If
Do While Not mDecoder.GetNextTick(tick)
    If Not getNextDecoder Then Exit Function
Loop
GetNextTick = True

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal instrumentID As Long, _
                ByVal fromDate As Date, _
                ByVal toDate As Date)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mDB = db
mInstrumentID = instrumentID
mFromDate = gTruncateTimeToMinute(fromDate)
mToDate = gTruncateTimeToNextMinute(toDate)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Public Sub StartStream( _
                ByVal asynchronous As Boolean)
Const ProcName As String = "StartStream"
Dim failpoint As String
On Error GoTo Err

mAsynchronous = asynchronous
Set mConnector = mDB.GetConnector(, IIf(mAsynchronous, ConnectOptionEnum.adAsyncConnect, 0))
mConnector.connect
If Not mAsynchronous Then
    doFetchTicks
    If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createNewTickDecoder() As TickDataDecoder
Dim PeriodStartTime As Date
Dim format As String
Dim BasePrice As Double
Dim TickSize As Double
Dim Data() As Byte

Const ProcName As String = "createNewTickDecoder"
Dim failpoint As String
On Error GoTo Err

PeriodStartTime = gRoundTimeToSecond(mRs("DateTime"))
format = mRs("Version")
BasePrice = mRs("baseprice")
TickSize = mRs("ticksize")
Data = mRs("data")

Set createNewTickDecoder = CreateTickDecoder(PeriodStartTime, TickSize, BasePrice, Data, format)
Debug.Print "createNewTickDecoder: " & FormatTimestamp(PeriodStartTime, TimestampDateAndTime + TimestampNoMillisecs)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub doFetchTicks()
Dim cmd As ADODB.Command
Dim param As ADODB.Parameter

Const ProcName As String = "doFetchTicks"
Dim failpoint As String
On Error GoTo Err

Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc

cmd.CommandText = "FetchTickData"

' @InstrumentID
Set param = cmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mInstrumentID)
cmd.Parameters.Append param

' @From
Set param = cmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            gRoundTimeToSecond(mFromDate))
cmd.Parameters.Append param

' @To
Set param = cmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            gRoundTimeToSecond(mToDate))
cmd.Parameters.Append param

cmd.ActiveConnection = mConnection

Set mRs = New ADODB.Recordset
mRs.cursorlocation = adUseClient
If mAsynchronous Then
    mRs.Properties("Initial Fetch Size") = 0
    mRs.Open cmd, _
            , _
            CursorTypeEnum.adOpenForwardOnly, _
            , _
            CommandTypeEnum.adCmdStoredProc + _
            ExecuteOptionEnum.adAsyncFetchNonBlocking + _
            ExecuteOptionEnum.adAsyncExecute
Else
    mRs.Open cmd, _
            , _
            CursorTypeEnum.adOpenForwardOnly, _
            , _
            CommandTypeEnum.adCmdStoredProc
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function getNextDecoder() As Boolean
Const ProcName As String = "getNextDecoder"
Dim failpoint As String
On Error GoTo Err

Set mDecoder = Nothing

If Not gIsStateSet(mRs.State, adStateOpen) Then
    gLogger.Log LogLevelDetail, ProjectName & "." & ModuleName & ":getNextDecoder: recordset state is " & mRs.State
    Debug.Print "Recordset state is " & mRs.State
    mWaitingForData = True
ElseIf mRs.EOF Then
    If Not mFetchComplete Then
        Debug.Print "getNextDecoder: waiting for data"
        mWaitingForData = True
    Else
        mEos = True
        Debug.Print "getNextDecoder: already at EOS"
    End If
Else
    mWaitingForData = False
    
    Set mDecoder = createNewTickDecoder
    getNextDecoder = True
    
    mRs.MoveNext
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function


