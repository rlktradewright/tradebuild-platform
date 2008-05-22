VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarDataReader"
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
Event BarsFetched()
Event Error( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
Event NotReady()
Event Ready()

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TradingDO26"
Private Const ModuleName As String = "BarDataReader"

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
Private WithEvents mRsBars As ADODB.Recordset
Attribute mRsBars.VB_VarHelpID = -1

Private mInstrumentID As Long

Private mIsConnecting As Boolean

Private mFetchInProgress As Boolean

Private mConnectionString As String

Private mTimeZone As TWUtilities30.TimeZone

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
                            ModuleName & ": connection error " & pError.NativeError & vbCrLf & _
                            pError.Description & vbCrLf & _
                            gGenerateErrorMessage(pConnection))
    Exit Sub
End If

mIsConnecting = False

RaiseEvent Ready

End Sub

Private Sub mConnection_ExecuteComplete( _
                ByVal RecordsAffected As Long, _
                ByVal pError As ADODB.Error, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pCommand As ADODB.Command, _
                ByVal pRecordset As ADODB.Recordset, _
                ByVal pConnection As ADODB.Connection)
mFetchInProgress = False
If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent Error(pError.Number, _
                    ModuleName & ": fetch error " & pError.NativeError & vbCrLf & _
                    pError.Description & vbCrLf & _
                    gGenerateErrorMessage(mConnection))
    mConnection.Close
End If
End Sub

'@===============================================================================
' mRsBars Event Handlers
'@===============================================================================

Private Sub mRsBars_FetchComplete( _
                            ByVal pError As ADODB.Error, _
                            adStatus As ADODB.EventStatusEnum, _
                            ByVal pRecordset As ADODB.Recordset)
mFetchInProgress = False
If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent Error(pError.Number, _
                    ModuleName & ": fetch error " & pError.NativeError & vbCrLf & _
                    pError.Description & vbCrLf & _
                    gGenerateErrorMessage(mConnection))
    mConnection.Close
    Exit Sub
End If
If Not mRsBars.BOF Then
    mRsBars.MoveLast    ' because the stored procedure returns them in reverse order
End If
Set mRsBars.ActiveConnection = Nothing
mConnection.Close
RaiseEvent BarsFetched
End Sub

'@===============================================================================
' Properties
'@===============================================================================

Public Property Get NumberOfBars() As Long
If mFetchInProgress Then
    NumberOfBars = -1
    Exit Property
End If
NumberOfBars = mRsBars.RecordCount
End Property

'@===============================================================================
' Methods
'@===============================================================================

Public Sub FetchBars( _
                        ByVal barType As Long, _
                        ByVal barLength As Long, _
                        ByVal NumberOfBars As Long, _
                        ByVal fromTime As Date, _
                        ByVal toTime As Date)

Dim cmd As ADODB.Command
Dim param As ADODB.Parameter
Dim fromLocal As Date
Dim toLocal As Date

Debug.Assert Not mFetchInProgress
mFetchInProgress = True

fromTime = gRoundTimeToSecond(fromTime)
toTime = gRoundTimeToSecond(toTime)

Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc

cmd.CommandText = "FetchBarData"

' @InstrumentID
Set param = cmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mInstrumentID)
cmd.Parameters.append param

' @BarType
Set param = cmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            barType)
cmd.Parameters.append param

' @BarLength
Set param = cmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            barLength)
cmd.Parameters.append param

' @NumberRequired
Set param = cmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            NumberOfBars)
cmd.Parameters.append param

If fromTime < #1/1/1900# Then fromTime = #1/1/1900# ' don't exceed range of SmallDateTime
If toTime < #1/1/1900# Then toTime = #1/1/1900# ' don't exceed range of SmallDateTime
' @From
Set param = cmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            fromTime)
cmd.Parameters.append param

' @To
Set param = cmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            toTime)
cmd.Parameters.append param

cmd.ActiveConnection = mConnection

Set mRsBars = New ADODB.Recordset
mRsBars.LockType = adLockReadOnly
mRsBars.Open cmd, _
            , _
            CursorTypeEnum.adOpenStatic, _
            , _
            CommandTypeEnum.adCmdStoredProc + _
            ExecuteOptionEnum.adAsyncFetchNonBlocking + _
            ExecuteOptionEnum.adAsyncExecute
                        
End Sub

Friend Sub initialise( _
                ByVal connectionString As String, _
                ByVal instrumentId As Long, _
                ByVal TimeZoneName As String)
mConnectionString = connectionString
mInstrumentID = instrumentId
Set mTimeZone = GetTimeZone(TimeZoneName)
connectToDB
End Sub

Public Function ReadBar( _
                        ByRef timestamp As Date, _
                        ByRef barType As Long, _
                        ByRef barLength As Long, _
                        ByRef openPrice As Double, _
                        ByRef highPrice As Double, _
                        ByRef lowPrice As Double, _
                        ByRef closePrice As Double, _
                        ByRef volume As Long, _
                        ByRef tickVolume As Long, _
                        ByRef openInterest As Long) As Boolean
' returns false if no more bars are available - this can
' occur before the full data has been retrieved from the database,
' so the caller needs to monitor the TickDataFetched event.

If mFetchInProgress Then Exit Function

If mRsBars.BOF Then
    Exit Function
End If

timestamp = gRoundTimeToSecond(mRsBars("DateTime"))
barType = mRsBars("BarType")
barLength = mRsBars("BarLengthMinutes")
openPrice = mRsBars("openprice")
highPrice = mRsBars("highprice")
lowPrice = mRsBars("lowprice")
closePrice = mRsBars("closeprice")
volume = mRsBars("volume")
tickVolume = mRsBars("tickvolume")
openInterest = mRsBars("openinterest")
mRsBars.MovePrevious

ReadBar = True
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


