VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickDataReader"
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

'@===============================================================================
' Events
'@===============================================================================

Event ConnectFailed( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
Event DataFetched()
Event Error( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
Event Ready()

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TradingDO26"
Private Const ModuleName As String = "TickDataReader"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private WithEvents mConnection As ADODB.Connection
Attribute mConnection.VB_VarHelpID = -1
Private WithEvents mRsTick As ADODB.recordset
Attribute mRsTick.VB_VarHelpID = -1

Private mInstrumentID As Long

Private mIsConnecting As Boolean

Private mFetchInProgress As Boolean

Private mConnectionString As String

'@===============================================================================
' Class Event Handlers
'@===============================================================================

'@===============================================================================
' XXXX Interface Members
'@===============================================================================

'@===============================================================================
' mConnection Event Handlers
'@===============================================================================

Private Sub mConnection_ConnectComplete( _
                            ByVal pError As ADODB.Error, _
                            adStatus As ADODB.EventStatusEnum, _
                            ByVal pConnection As ADODB.Connection)

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent ConnectFailed(pError.Number, _
                            ModuleName & ": connection error" & vbCrLf & _
                            pError.Description & vbCrLf & _
                            gGenerateErrorMessage(pConnection))
    Exit Sub
End If

mIsConnecting = False

RaiseEvent Ready

End Sub

Private Sub mConnection_InfoMessage( _
                            ByVal pError As ADODB.Error, _
                            adStatus As ADODB.EventStatusEnum, _
                            ByVal pConnection As ADODB.Connection)
If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    Stop
End If
End Sub

'@===============================================================================
' mRsTick Event Handlers
'@===============================================================================

Private Sub mRsTick_FetchComplete( _
                            ByVal pError As ADODB.Error, _
                            adStatus As ADODB.EventStatusEnum, _
                            ByVal pRecordset As ADODB.recordset)
If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent Error(pError.Number, _
                    ModuleName & ": fetch error" & vbCrLf & _
                    pError.Description & vbCrLf & _
                    gGenerateErrorMessage(mConnection))
    Exit Sub
End If
mFetchInProgress = False
RaiseEvent DataFetched
End Sub

Private Sub mRsTick_FetchProgress( _
                            ByVal Progress As Long, _
                            ByVal MaxProgress As Long, _
                            adStatus As ADODB.EventStatusEnum, _
                            ByVal pRecordset As ADODB.recordset)

End Sub

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Public Sub FetchData( _
                        ByVal Version As String, _
                        ByVal fromTime As Date, _
                        ByVal toTime As Date)

Debug.Assert Not mFetchInProgress
mFetchInProgress = True

DoQueryTickData Version, fromTime, toTime
                        
End Sub

Friend Sub initialise( _
                ByVal connectionString As String, _
                ByVal instrumentId As Long)
mConnectionString = connectionString
mInstrumentID = instrumentId
connectToDB
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

If CBool(mRsTick("IsUTCTime")) Then
    PeriodStartTime = mRsTick("DateTime")
Else
    PeriodStartTime = ConvertDateLocalToUTC(mRsTick("DateTime"))
End If
BasePrice = mRsTick("baseprice")
Data = mRsTick("data")
mRsTick.MoveNext

ReadData = True
End Function

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub connectToDB()
If Not mConnection Is Nothing Then mConnection.Close
mIsConnecting = True
Set mConnection = New ADODB.Connection
mConnection.CursorLocation = CursorLocationEnum.adUseClient
mConnection.Open mConnectionString, , , ConnectOptionEnum.adAsyncConnect
End Sub

Private Sub DoQueryTickData( _
                        ByVal Version As String, _
                        ByVal fromTime As Date, _
                        ByVal toTime As Date)
Dim cmd As ADODB.Command
Dim param As ADODB.Parameter

Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc

cmd.CommandText = "FetchTickDataEx"

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

Set param = cmd.CreateParameter("@FromUTC", _
                            DataTypeEnum.adDate, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            fromTime)
cmd.Parameters.Append param

Set param = cmd.CreateParameter("@ToUTC", _
                            DataTypeEnum.adDate, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            toTime)
cmd.Parameters.Append param

Set param = cmd.CreateParameter("@FromLocal", _
                            DataTypeEnum.adDate, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            ConvertDateUTCToLocal(fromTime))
cmd.Parameters.Append param

Set param = cmd.CreateParameter("@ToLocal", _
                            DataTypeEnum.adDate, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            ConvertDateUTCToLocal(toTime))
cmd.Parameters.Append param

cmd.ActiveConnection = mConnection

Set mRsTick = New ADODB.recordset
mRsTick.Open cmd, _
                , _
                CursorTypeEnum.adOpenForwardOnly, _
                , _
                CommandTypeEnum.adCmdStoredProc + _
                ExecuteOptionEnum.adAsyncFetchNonBlocking + _
                ExecuteOptionEnum.adAsyncExecute
                        
End Sub




