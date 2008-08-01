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

Private Const ProjectName                   As String = "TradingDO26"
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
If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent Error(pError.Number, _
                    ModuleName & ": fetch error " & vbCrLf & _
                    gGenerateErrorMessage(pError) & _
                    gGenerateConnectionErrorMessages(mConnection))
    If mAsynchronous Then
        If mConnection.State = adStateOpen Then mConnection.Close
    End If
End If
End Sub

'@===============================================================================
' mConnector Event Handlers
'@===============================================================================

Private Sub mConnector_Connected( _
                ByVal pConnection As ADODB.connection)
Set mConnection = pConnection
If mAsynchronous Then doFetchTicks
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
If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent Error(pError.Number, _
                    ModuleName & ": fetch error " & vbCrLf & _
                    gGenerateErrorMessage(pError) & _
                    gGenerateConnectionErrorMessages(mConnection))
    Exit Sub
End If
Set mRs.ActiveConnection = Nothing
If mConnection.State = adStateOpen Then mConnection.Close
mFetchComplete = True

Debug.Print "Fetch complete"

If Not mAsynchronous Then
    RaiseEvent DataAvailable
ElseIf mWaitingForData Then
    mWaitingForData = False
    RaiseEvent DataAvailable
End If
End Sub

Private Sub mRs_FetchProgress( _
                ByVal Progress As Long, _
                ByVal MaxProgress As Long, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pRecordset As ADODB.Recordset)

Debug.Print "Fetch progress=" & Progress & "; max=" & MaxProgress
If Not mAsynchronous Then Exit Sub
If mWaitingForData Then
    mWaitingForData = False
    RaiseEvent DataAvailable
End If
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get eos() As Boolean
eos = mEos
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub cancel()
If Not mRs Is Nothing Then
    If mRs.State = adStateOpen Then
        mRs.Close
    Else
        mRs.cancel
    End If
    Set mRs = Nothing
End If

If Not mConnection Is Nothing Then
    If mConnection.State = adStateOpen Then
        mConnection.Close
    Else
        mConnection.cancel
    End If
    Set mConnection = Nothing
End If
    
End Sub

Public Function getNextTick( _
                ByRef tick As GenericTick) As Boolean

If mRs Is Nothing Then
    If Not mAsynchronous Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & "getNextTick", _
                "Must call startStream befoe calling getNextTick"
    End If
    
    Debug.Print "getNextTick: no recordset"
    mWaitingForData = True
    Exit Function
End If

If mDecoder Is Nothing Then
    If Not getNextDecoder Then Exit Function
End If
Do While Not mDecoder.getNextTick(tick)
    If Not getNextDecoder Then Exit Function
Loop
getNextTick = True

End Function

Friend Sub initialise( _
                ByVal db As TradingDB, _
                ByVal instrumentID As Long, _
                ByVal fromDate As Date, _
                ByVal toDate As Date)
Set mDB = db
mInstrumentID = instrumentID
mFromDate = gTruncateTimeToMinute(fromDate)
mToDate = gTruncateTimeToNextMinute(toDate)

End Sub

Public Sub startStream( _
                ByVal asynchronous As Boolean)
mAsynchronous = asynchronous
Set mConnector = mDB.getConnector(, IIf(mAsynchronous, ConnectOptionEnum.adAsyncConnect, 0))
mConnector.connect
If Not mAsynchronous Then
    doFetchTicks
    If mConnection.State = adStateOpen Then mConnection.Close
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createNewTickDecoder() As TickDataDecoder
Dim PeriodStartTime As Date
Dim format As String
Dim BasePrice As Double
Dim tickSize As Double
Dim Data() As Byte

PeriodStartTime = gRoundTimeToSecond(mRs("DateTime"))
format = mRs("Version")
BasePrice = mRs("baseprice")
tickSize = mRs("ticksize")
Data = mRs("data")

Set createNewTickDecoder = CreateTickDecoder(PeriodStartTime, tickSize, BasePrice, Data, format)
Debug.Print "createNewTickDecoder: " & FormatTimestamp(PeriodStartTime, TimestampDateAndTime + TimestampNoMillisecs)
End Function

Private Sub doFetchTicks()
Dim cmd As ADODB.Command
Dim param As ADODB.Parameter

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
End Sub

Private Function getNextDecoder() As Boolean
Set mDecoder = Nothing

If mRs.State <> adStateOpen Then
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
End Function


