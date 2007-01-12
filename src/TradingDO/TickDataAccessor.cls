VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickDataAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

Event ConnectFailed( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
Event DataFetched()
Event Error( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
Event NotReady()
Event Ready()

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mConnection As ADODB.Connection
Attribute mConnection.VB_VarHelpID = -1
Private WithEvents mRsTick As ADODB.Recordset
Attribute mRsTick.VB_VarHelpID = -1

Private mInstrumentID As Long

Private mIsConnecting As Boolean

Private mFetchInProgress As Boolean

Private mCommandToRetry As ADODB.Command

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
connectToDB
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' mConnection Event Handlers
'================================================================================

Private Sub mConnection_ConnectComplete( _
                            ByVal pError As ADODB.Error, _
                            adStatus As ADODB.EventStatusEnum, _
                            ByVal pConnection As ADODB.Connection)

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent ConnectFailed(pError.Number, _
                            "TickDataAccessor: connection error" & vbCrLf & _
                            pError.Description & vbCrLf & _
                            generateErrorMessage(pConnection))
    Exit Sub
End If

mIsConnecting = False

If Not mCommandToRetry Is Nothing Then
    mCommandToRetry.ActiveConnection = mConnection
    mCommandToRetry.Execute , _
                    , _
                    ExecuteOptionEnum.adAsyncExecute + ExecuteOptionEnum.adExecuteNoRecords
    Set mCommandToRetry = Nothing
End If

RaiseEvent Ready

End Sub

Private Sub mConnection_ExecuteComplete( _
                            ByVal RecordsAffected As Long, _
                            ByVal pError As ADODB.Error, _
                            adStatus As ADODB.EventStatusEnum, _
                            ByVal pCommand As ADODB.Command, _
                            ByVal pRecordset As ADODB.Recordset, _
                            ByVal pConnection As ADODB.Connection)
If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent Error(pError.Number, _
                    "TickDataAccessor: execution error" & vbCrLf & _
                    pError.Description & vbCrLf & _
                    generateErrorMessage(pConnection))
    RaiseEvent NotReady
    Set mCommandToRetry = pCommand
    connectToDB
    Exit Sub
End If
End Sub

Private Sub mConnection_InfoMessage( _
                            ByVal pError As ADODB.Error, _
                            adStatus As ADODB.EventStatusEnum, _
                            ByVal pConnection As ADODB.Connection)
If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    Stop
End If
End Sub

'================================================================================
' mRsTick Event Handlers
'================================================================================

Private Sub mRsTick_FetchComplete( _
                            ByVal pError As ADODB.Error, _
                            adStatus As ADODB.EventStatusEnum, _
                            ByVal pRecordset As ADODB.Recordset)
If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent Error(pError.Number, _
                    "BarDataAccessor: fetch error" & vbCrLf & _
                    pError.Description & vbCrLf & _
                    generateErrorMessage(mConnection))
    Exit Sub
End If
mFetchInProgress = False
RaiseEvent DataFetched
End Sub

Private Sub mRsTick_FetchProgress( _
                            ByVal Progress As Long, _
                            ByVal MaxProgress As Long, _
                            adStatus As ADODB.EventStatusEnum, _
                            ByVal pRecordset As ADODB.Recordset)

End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get connectionString() As String
connectionString = setting("connectionstring")
End Property

Friend Property Let InstrumentID(ByVal value As Long)
mInstrumentID = value
End Property

'================================================================================
' Methods
'================================================================================

Public Sub FetchData( _
                        ByVal Version As String, _
                        ByVal fromTime As Date, _
                        ByVal toTime As Date)

Debug.Assert Not mFetchInProgress
mFetchInProgress = True

DoQueryTickData Version, fromTime, toTime
                        
End Sub

Public Function ReadData( _
                        ByRef PeriodStartTime As Date, _
                        ByRef BasePrice As Double, _
                        ByRef Data() As Byte) As Boolean
' returns false if no more tickdata is available - this can
' occur before the full data has been retrieved from the database,
' so the caller needs to monitor the TickDataFetched event.

If mFetchInProgress Then Exit Function

If mRsTick.EOF Then
    Exit Function
End If

PeriodStartTime = mRsTick("DateTime")
BasePrice = mRsTick("baseprice")
Data = mRsTick("data")
mRsTick.MoveNext

ReadData = True
End Function

Public Sub WriteData( _
                        ByVal Version As String, _
                        ByVal PeriodStartTime As Date, _
                        ByVal BasePrice As Double, _
                        ByRef Data() As Byte)

DoWriteTickData Version, PeriodStartTime, BasePrice, Data

End Sub
'================================================================================
' Helper Functions
'================================================================================

Private Sub connectToDB()
If Not mConnection Is Nothing Then mConnection.Close
mIsConnecting = True
Set mConnection = New ADODB.Connection
mConnection.CursorLocation = CursorLocationEnum.adUseClient
mConnection.Open setting("connectionstring"), , , ConnectOptionEnum.adAsyncConnect
End Sub

Private Sub DoQueryTickData( _
                        ByVal Version As String, _
                        ByVal fromTime As Date, _
                        ByVal toTime As Date)
Dim cmd As ADODB.Command
Dim param As ADODB.Parameter

Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc

cmd.CommandText = "FetchTickData"

Set param = cmd.CreateParameter("@InstrumentID", _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mInstrumentID)
cmd.Parameters.Append param

Set param = cmd.CreateParameter("@DataVersion", _
                                DataTypeEnum.adVarChar, _
                            ParameterDirectionEnum.adParamInput, _
                            255, _
                            Version)
cmd.Parameters.Append param

Set param = cmd.CreateParameter("@FromTime", _
                            DataTypeEnum.adDate, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            fromTime)
cmd.Parameters.Append param

Set param = cmd.CreateParameter("@ToTime", _
                            DataTypeEnum.adDate, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            toTime)
cmd.Parameters.Append param

cmd.ActiveConnection = mConnection

Set mRsTick = New ADODB.Recordset
mRsTick.Open cmd, _
                , _
                CursorTypeEnum.adOpenForwardOnly, _
                , _
                CommandTypeEnum.adCmdStoredProc + ExecuteOptionEnum.adAsyncFetchNonBlocking
                        
End Sub

Private Sub DoWriteTickData( _
                        ByVal Version As String, _
                        ByVal PeriodStartTime As Date, _
                        ByVal BasePrice As Double, _
                        ByRef Data() As Byte)

Dim cmd As ADODB.Command
Dim param As ADODB.Parameter

Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc

cmd.CommandText = "WriteTickData"

Set param = cmd.CreateParameter("@InstrumentID", _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mInstrumentID)
cmd.Parameters.Append param

Set param = cmd.CreateParameter("@DataVersion", _
                            DataTypeEnum.adVarChar, _
                            ParameterDirectionEnum.adParamInput, _
                            255, _
                            Version)
cmd.Parameters.Append param

Set param = cmd.CreateParameter("@DateAndTime", _
                            DataTypeEnum.adDate, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            PeriodStartTime)
cmd.Parameters.Append param

Set param = cmd.CreateParameter("@BasePrice", _
                            DataTypeEnum.adCurrency, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            BasePrice)
cmd.Parameters.Append param

Set param = cmd.CreateParameter("@Data", _
                            DataTypeEnum.adBinary, _
                            ParameterDirectionEnum.adParamInput, _
                            UBound(Data) + 1, _
                            Data)
cmd.Parameters.Append param

cmd.ActiveConnection = mConnection
cmd.Execute , _
            , _
            ExecuteOptionEnum.adAsyncExecute + ExecuteOptionEnum.adExecuteNoRecords

End Sub





