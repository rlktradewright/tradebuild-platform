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
Private WithEvents mConnection              As ADODB.Connection
Attribute mConnection.VB_VarHelpID = -1
Private WithEvents mRS                      As ADODB.Recordset
Attribute mRS.VB_VarHelpID = -1

Private mAsynchronous                       As Boolean

Private mInstrumentID                       As Long

Private mFromDate                           As Date
Private mToDate                             As Date

Private mFetchComplete                      As Boolean

Private mEos                                As Boolean

Private mDecoder                            As ITickDataDecoder

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
                ByVal pConnection As ADODB.Connection)
Const ProcName As String = "mConnection_ExecuteComplete"
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
    If mAsynchronous Then
        Set mRS.ActiveConnection = Nothing
        If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
    Else
        mFetchComplete = True
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
                ByVal pConnection As ADODB.Connection)
Const ProcName As String = "mConnector_Connected"
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
On Error GoTo Err

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent Error(pError.Number, _
                    ModuleName & ": fetch error " & vbCrLf & _
                    gGenerateErrorMessage(pError) & _
                    gGenerateConnectionErrorMessages(mConnection))
    Exit Sub
End If

mFetchComplete = True

RaiseEvent DataAvailable

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
On Error GoTo Err

Debug.Print "Fetch progress=" & Progress & "; max=" & MaxProgress
If Not mAsynchronous Then Exit Sub

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
On Error GoTo Err

If Not mRS Is Nothing Then
    If gIsStateSet(mRS.State, adStateOpen) Then
        mRS.Close
    Else
        mRS.Cancel
    End If
    Set mRS = Nothing
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetNextTick( _
                ByRef tick As GenericTick) As Boolean
Const ProcName As String = "GetNextTick"
On Error GoTo Err

If mRS Is Nothing Then
    Assert mAsynchronous, "Must call StartStream before calling GetNextTick"
        
    Debug.Print "GetNextTick: no recordset"
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
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal instrumentID As Long, _
                ByVal FromDate As Date, _
                ByVal ToDate As Date)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mDB = db
mInstrumentID = instrumentID
mFromDate = gTruncateTimeToMinute(FromDate)
mToDate = gTruncateTimeToNextMinute(ToDate)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StartStream( _
                ByVal asynchronous As Boolean)
Const ProcName As String = "StartStream"
On Error GoTo Err

mAsynchronous = asynchronous
Set mConnector = mDB.GetConnector
If mAsynchronous Then
    mConnector.ConnectAsync
Else
    mConnector.Connect
    doFetchTicks
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createNewTickDecoder() As ITickDataDecoder
Const ProcName As String = "createNewTickDecoder"
On Error GoTo Err

Dim PeriodStartTime As Date
PeriodStartTime = gRoundTimeToSecond(mRS("DateTime"))

Dim format As String
format = mRS("Version")

Dim BasePrice As Double
BasePrice = mRS("baseprice")

Dim TickSize As Double
TickSize = mRS("ticksize")

Dim data() As Byte
data = mRS("data")

Set createNewTickDecoder = CreateTickDecoder(PeriodStartTime, TickSize, BasePrice, data, format)
Debug.Print "createNewTickDecoder: " & FormatTimestamp(PeriodStartTime, TimestampDateAndTime + TimestampNoMillisecs)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub doFetchTicks()
Const ProcName As String = "doFetchTicks"
On Error GoTo Err

Dim cmd As ADODB.Command
Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "FetchTickData"

Dim param As ADODB.Parameter

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

Set mRS = New ADODB.Recordset
mRS.cursorlocation = adUseClient
If mAsynchronous Then
    mRS.Properties("Initial Fetch Size") = 0
    mRS.Open cmd, _
            , _
            CursorTypeEnum.adOpenForwardOnly, _
            , _
            CommandTypeEnum.adCmdStoredProc + _
            ExecuteOptionEnum.adAsyncFetchNonBlocking + _
            ExecuteOptionEnum.adAsyncExecute
Else
    mRS.Open cmd, _
            , _
            CursorTypeEnum.adOpenForwardOnly, _
            , _
            CommandTypeEnum.adCmdStoredProc
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getNextDecoder() As Boolean
Const ProcName As String = "getNextDecoder"
On Error GoTo Err

Set mDecoder = Nothing

If Not gIsStateSet(mRS.State, adStateOpen) Then
    gLogger.Log "recordset state is " & mRS.State, ProcName, ModuleName, LogLevelDetail
    Debug.Print "Recordset state is " & mRS.State
ElseIf mRS.EOF Then
    If Not mFetchComplete Then
        Debug.Print "getNextDecoder: waiting for data"
    Else
        mEos = True
        Debug.Print "getNextDecoder: already at EOS"
    End If
Else
    Set mDecoder = createNewTickDecoder
    getNextDecoder = True
    
    mRS.MoveNext
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


