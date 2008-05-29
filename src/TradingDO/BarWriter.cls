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

'Event ConnectFailed( _
'                ByVal errorCode As Long, _
'                ByVal errorDesc As String)
'
'Event DataWritten()
'
'Event Error( _
'                ByVal errorCode As Long, _
'                ByVal errorDesc As String)

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

Private WithEvents mConnection                  As ADODB.Connection
Attribute mConnection.VB_VarHelpID = -1

Private mInstrumentID                           As Long

Private mCmd                                    As ADODB.Command

Private mConnectionString                       As String

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

'@===============================================================================
' Class Event Handlers
'@===============================================================================

'@===============================================================================
' Task Interface Members
'@===============================================================================

Private Sub Task_cancel()
If Not mCmd Is Nothing Then
    mCmd.cancel
    mConnection.Close
ElseIf Not mConnection Is Nothing Then
    mConnection.cancel
    mConnection.Close
End If
mTaskContext.finish Empty, True
End Sub

Private Sub Task_run()
connectToDB
mTaskContext.sleep -1
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

Private Sub mConnection_ConnectComplete( _
                ByVal pError As ADODB.Error, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pConnection As ADODB.Connection)

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    If mAsynchronous Then
        mTaskContext.Error pError.Number, _
                            "Connection error " & vbCrLf & _
                            gGenerateErrorMessage(pError) & _
                            gGenerateConnectionErrorMessages(pConnection)
        Exit Sub
    Else
        Err.Raise pError.Number, _
                    ProjectName & "." & ModuleName & ":" & "mConnection_ConnectComplete", _
                    "Connection error " & vbCrLf & _
                    gGenerateErrorMessage(pError) & _
                    gGenerateConnectionErrorMessages(pConnection)
    End If
End If

If mAsynchronous Then doWriteBar
End Sub

Private Sub mConnection_ExecuteComplete( _
                ByVal RecordsAffected As Long, _
                ByVal pError As ADODB.Error, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pCommand As ADODB.Command, _
                ByVal pRecordset As ADODB.Recordset, _
                ByVal pConnection As ADODB.Connection)
Dim s As String
If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    s = "Execution error " & vbCrLf & _
        gGenerateErrorMessage(pError) & _
        gGenerateConnectionErrorMessages(pConnection)
    
    mConnection.Close
    If mAsynchronous Then
        mTaskContext.Error pError.Number, ProjectName & "." & ModuleName & ":" & "mConnection_ExecuteComplete" & vbCrLf & s
        Exit Sub
    Else
        Err.Raise pError.Number, _
                    ProjectName & "." & ModuleName & ":" & "mConnection_ExecuteComplete", _
                    s
    End If
Else
    mConnection.Close
    If mAsynchronous Then mTaskContext.finish 0, False
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
                ByVal instrumentId As Long)
mConnectionString = connectionString
mInstrumentID = instrumentId
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

connectToDB

doWriteBar

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
                ByVal cookie As Variant) As TaskController

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

Set WriteBarAsync = StartTask(Me, PriorityNormal, , cookie)

End Function

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub connectToDB()
If Not mConnection Is Nothing Then mConnection.Close
Set mConnection = New ADODB.Connection
mConnection.CursorLocation = CursorLocationEnum.adUseClient
If mAsynchronous Then
    mConnection.Open mConnectionString, , , ConnectOptionEnum.adAsyncConnect
Else
    mConnection.Open mConnectionString
End If
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
mCmd.Parameters.append param

' @BarType
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mBarType)
mCmd.Parameters.append param

' @BarLength
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mBarLength)
mCmd.Parameters.append param

' @DateAndTime
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mTimestamp)
mCmd.Parameters.append param

' @OpenPrice
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDouble, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mOpenPrice)
mCmd.Parameters.append param

' @HighPrice
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDouble, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mHighPrice)
mCmd.Parameters.append param

' @LowPrice
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDouble, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mLowPrice)
mCmd.Parameters.append param

' @ClosePrice
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adDouble, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mClosePrice)
mCmd.Parameters.append param

' @Volume
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mVolume)
mCmd.Parameters.append param

' @TickVolume
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mTickVolume)
mCmd.Parameters.append param

' @OpenInterest
Set param = mCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mOpenInterest)
mCmd.Parameters.append param

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






