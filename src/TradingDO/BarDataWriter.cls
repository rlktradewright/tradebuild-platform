VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarDataWriter"
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

Event DataWritten()

Event Error( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TradingDO26"
Private Const ModuleName As String = "BarDataWriter"

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

Private mInstrumentID As Long

Private mIsConnecting As Boolean

Private mCmd As ADODB.Command

Private mConnectionString As String

Private mParent As instrument

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
    mParent.releaseWriter Me
    Exit Sub
End If

mIsConnecting = False

Set mCmd.ActiveConnection = mConnection
mCmd.Execute , _
            , _
            ExecuteOptionEnum.adAsyncExecute + ExecuteOptionEnum.adExecuteNoRecords

End Sub

Private Sub mConnection_ExecuteComplete( _
                            ByVal RecordsAffected As Long, _
                            ByVal pError As ADODB.Error, _
                            adStatus As ADODB.EventStatusEnum, _
                            ByVal pCommand As ADODB.Command, _
                            ByVal pRecordset As ADODB.recordset, _
                            ByVal pConnection As ADODB.Connection)
If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent Error(pError.Number, _
                    ModuleName & ": execution error" & vbCrLf & _
                    pError.Description & vbCrLf & _
                    gGenerateErrorMessage(pConnection))
    connectToDB
Else
    mConnection.Close
    RaiseEvent DataWritten
    mParent.releaseWriter Me
End If
End Sub

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub initialise( _
                ByVal connectionString As String, _
                ByVal instrumentId As Long, _
                ByVal parent As instrument)
mConnectionString = connectionString
mInstrumentID = instrumentId
Set mParent = parent
End Sub

Public Sub WriteBar( _
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

Dim param As ADODB.Parameter

Set mCmd = New ADODB.Command
mCmd.CommandType = adCmdStoredProc

mCmd.CommandText = "WriteBarDataUTC"

Set param = mCmd.CreateParameter("@InstrumentID", _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mInstrumentID)
mCmd.Parameters.Append param

Set param = mCmd.CreateParameter("@BarType", _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            barType)
mCmd.Parameters.Append param

Set param = mCmd.CreateParameter("@BarLength", _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            barLength)
mCmd.Parameters.Append param

Set param = mCmd.CreateParameter("@DateAndTime", _
                            DataTypeEnum.adDate, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            timestamp)
mCmd.Parameters.Append param

Set param = mCmd.CreateParameter("@OpenPrice", _
                            DataTypeEnum.adCurrency, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            openPrice)
mCmd.Parameters.Append param

Set param = mCmd.CreateParameter("@HighPrice", _
                            DataTypeEnum.adCurrency, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            highPrice)
mCmd.Parameters.Append param

Set param = mCmd.CreateParameter("@LowPrice", _
                            DataTypeEnum.adCurrency, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            lowPrice)
mCmd.Parameters.Append param

Set param = mCmd.CreateParameter("@ClosePrice", _
                            DataTypeEnum.adCurrency, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            closePrice)
mCmd.Parameters.Append param

Set param = mCmd.CreateParameter("@Volume", _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            volume)
mCmd.Parameters.Append param

Set param = mCmd.CreateParameter("@TickVolume", _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            tickVolume)
mCmd.Parameters.Append param

Set param = mCmd.CreateParameter("@OpenInterest", _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            openInterest)
mCmd.Parameters.Append param

connectToDB
                        
End Sub

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




