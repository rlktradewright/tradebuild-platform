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
Private Const ModuleName As String = "TickDataWriter"

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

Private mWriterId As String

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
                            "TickDataAccessor: connection error" & vbCrLf & _
                            pError.Description & vbCrLf & _
                            gGenerateErrorMessage(pConnection))
    Debug.Print "Writer " & mWriterId & " error connecting to db: (" & pError.Number & ") " & pError.Description
    mParent.releaseWriter Me
    Exit Sub
End If

mIsConnecting = False

Debug.Print "Writer " & mWriterId & " connected to db"
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
                    "TickDataAccessor: execution error" & vbCrLf & _
                    pError.Description & vbCrLf & _
                    gGenerateErrorMessage(pConnection))
    Debug.Print "Writer " & mWriterId & " error writing tick data: (" & pError.Number & ") " & pError.Description
    connectToDB
Else
    mConnection.Close
    Debug.Print "Writer " & mWriterId & " data written"
    RaiseEvent DataWritten
    mParent.releaseWriter Me
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

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub initialise( _
                ByVal connectionString As String, _
                ByVal instrumentId As Long, _
                ByVal parent As instrument, _
                ByVal writerId As String)
mConnectionString = connectionString
mInstrumentID = instrumentId
Set mParent = parent
mWriterId = writerId
End Sub

Public Sub WriteData( _
                        ByVal Version As String, _
                        ByVal PeriodStartTime As Date, _
                        ByVal BasePrice As Double, _
                        ByRef Data() As Byte)
If Not mCmd Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":WriteData", _
                "Object has already been used"
End If

DoWriteTickData Version, PeriodStartTime, BasePrice, Data

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

Private Sub DoWriteTickData( _
                        ByVal Version As String, _
                        ByVal PeriodStartTime As Date, _
                        ByVal BasePrice As Double, _
                        ByRef Data() As Byte)

Dim param As ADODB.Parameter

Set mCmd = New ADODB.Command
mCmd.CommandType = adCmdStoredProc

mCmd.CommandText = "WriteTickDataUTC"

Set param = mCmd.CreateParameter("@InstrumentID", _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mInstrumentID)
mCmd.Parameters.Append param

Set param = mCmd.CreateParameter("@DataVersion", _
                            DataTypeEnum.adVarChar, _
                            ParameterDirectionEnum.adParamInput, _
                            255, _
                            Version)
mCmd.Parameters.Append param

Set param = mCmd.CreateParameter("@DateAndTime", _
                            DataTypeEnum.adDate, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            PeriodStartTime)
mCmd.Parameters.Append param

Set param = mCmd.CreateParameter("@BasePrice", _
                            DataTypeEnum.adCurrency, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            BasePrice)
mCmd.Parameters.Append param

Set param = mCmd.CreateParameter("@Data", _
                            DataTypeEnum.adBinary, _
                            ParameterDirectionEnum.adParamInput, _
                            UBound(Data) + 1, _
                            Data)
mCmd.Parameters.Append param

connectToDB
End Sub





