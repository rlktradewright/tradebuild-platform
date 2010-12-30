VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickWriter"
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

Implements TaskCompletionListener

'@================================================================================
' Events
'@================================================================================

Event ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)

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

'@================================================================================
' Member variables
'@================================================================================

Private Const ModuleName                    As String = "TickWriter"

Private mTde                                As TickDataEncoder

Private mDataAdded                          As Boolean

Private mTotalBytes                         As Long
Private mOpeningTimestamp                   As Date   ' timestamp at start of minute period

Private mInstrument                         As instrument
Private mTickSize                           As Double

Private mDB                                 As TradingDB

Private WithEvents mWriteTimer              As IntervalTimer
Attribute mWriteTimer.VB_VarHelpID = -1

Private WithEvents mTickWriter              As TickDataWriter
Attribute mTickWriter.VB_VarHelpID = -1

Private mUseAsynchronousWrites              As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Randomize
End Sub

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_taskCompleted( _
                ev As TWUtilities30.TaskCompletionEventData)
If ev.Cancelled Then
ElseIf ev.ErrorNumber <> 0 Then
    RaiseEvent Error(ev.ErrorNumber, ev.ErrorMessage)
Else
    If gLogger.IsLoggable(LogLevelMediumDetail) Then
        gLogger.Log LogLevelMediumDetail, _
                    "TickWriter: successfully wrote data for " & mInstrument.ShortName & ": period " & FormatTimestamp(CDate(ev.cookie), TimestampDateAndTimeISO8601)
    End If
End If
End Sub

'@================================================================================
' mTickWriter Event Handlers
'@================================================================================

Private Sub mTickWriter_ConnectFailed(ByVal errorCode As Long, ByVal errorDesc As String)
Const ProcName As String = "mTickWriter_ConnectFailed"

On Error GoTo Err

RaiseEvent ConnectFailed(errorCode, errorDesc)
Set mTickWriter = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mTickWriter_DataWritten()
Set mTickWriter = Nothing
End Sub

Private Sub mTickWriter_Error(ByVal errorCode As Long, ByVal errorDesc As String)
RaiseEvent Error(errorCode, errorDesc)
Set mTickWriter = Nothing
End Sub

'@================================================================================
' mWriteTimer Event Handlers
'@================================================================================

Private Sub mWriteTimer_TimerExpired()
Const ProcName As String = "mWriteTimer_TimerExpired"

On Error GoTo Err

writePeriodData
startWriteTimer

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get BytesWritten() As Long
BytesWritten = mTotalBytes
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal pInstrument As instrument, _
                ByVal useAsynchronousWrites As Boolean)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mDB = db
Set mInstrument = pInstrument
mTickSize = mInstrument.TickSize
mUseAsynchronousWrites = useAsynchronousWrites
startWriteTimer

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"


On Error GoTo Err

writePeriodData

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub WriteAsk( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)
Const ProcName As String = "WriteAsk"


On Error GoTo Err

If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeAsk timestamp, price, size
mDataAdded = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub WriteBid( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)
Const ProcName As String = "WriteBid"


On Error GoTo Err

If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeBid timestamp, price, size
mDataAdded = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub WriteClose( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Const ProcName As String = "WriteClose"


On Error GoTo Err

If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeClose timestamp, price
mDataAdded = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub WriteHigh( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Const ProcName As String = "WriteHigh"


On Error GoTo Err

If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeHigh timestamp, price
mDataAdded = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub WriteLow( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Const ProcName As String = "WriteLow"


On Error GoTo Err

If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeLow timestamp, price
mDataAdded = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub WriteMarketDepthData( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal marketmaker As String, _
                ByVal operation As Long, _
                ByVal side As Long, _
                ByVal price As Double, _
                ByVal size As Long)
Const ProcName As String = "WriteMarketDepthData"


On Error GoTo Err

If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeMarketDepthData timestamp, position, marketmaker, operation, side, price, size
mDataAdded = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub WriteMarketDepthReset( _
                ByVal timestamp As Date)
Const ProcName As String = "WriteMarketDepthReset"


On Error GoTo Err

checkIfNewPeriod timestamp
mTde.encodeMarketDepthReset timestamp
mDataAdded = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub WriteOpen( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Const ProcName As String = "WriteOpen"


On Error GoTo Err

checkIfNewPeriod timestamp
mTde.encodeOpen timestamp, price
mDataAdded = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub WriteOpenInterest( _
                ByVal timestamp As Date, _
                ByVal size As Long)
Const ProcName As String = "WriteOpenInterest"


On Error GoTo Err

checkIfNewPeriod timestamp
mTde.encodeOpenInterest timestamp, size
mDataAdded = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub WriteTick( _
                ByRef tick As GenericTick)
Const ProcName As String = "WriteTick"


On Error GoTo Err

With tick
    Select Case .tickType
    Case TickTypes.TickTypeAsk
        WriteAsk .timestamp, .price, .size
    Case TickTypes.TickTypeBid
        WriteBid .timestamp, .price, .size
    Case TickTypes.TickTypeClosePrice
        WriteClose .timestamp, .price
    Case TickTypes.TickTypeHighPrice
        WriteHigh .timestamp, .price
    Case TickTypes.TickTypeLowPrice
        WriteLow .timestamp, .price
    Case TickTypes.TickTypeMarketDepth
        WriteMarketDepthData .timestamp, .position, .marketmaker, .operation, .side, .price, .size
    Case TickTypes.TickTypeMarketDepthReset
        WriteMarketDepthReset .timestamp
    Case TickTypes.TickTypeOpenInterest
        WriteOpenInterest .timestamp, .size
    Case TickTypes.TickTypeOpenPrice
        WriteOpen .timestamp, .price
    Case TickTypes.TickTypeTrade
        WriteTrade .timestamp, .price, .size
    Case TickTypes.TickTypeVolume
        WriteVolume .timestamp, .size
    End Select
End With

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub WriteTrade( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)
Const ProcName As String = "WriteTrade"


On Error GoTo Err

If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeTrade timestamp, price, size
mDataAdded = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub WriteVolume( _
                ByVal timestamp As Date, _
                ByVal size As Long)
Const ProcName As String = "WriteVolume"


On Error GoTo Err

checkIfNewPeriod timestamp
mTde.encodeVolume timestamp, size
mDataAdded = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkIfNewPeriod(ByVal timestamp As Date)
Const ProcName As String = "checkIfNewPeriod"
Dim barTime As Date

On Error GoTo Err

barTime = gTruncateTimeToMinute(timestamp)
If barTime = mOpeningTimestamp Then Exit Sub

If Not mTde Is Nothing Then
    writePeriodData
    mTotalBytes = mTotalBytes + mTde.encodedDataLength
End If

Set mTde = CreateTickEncoder(barTime, mTickSize)
If Not mWriteTimer Is Nothing Then mWriteTimer.StopTimer
startWriteTimer

mOpeningTimestamp = barTime

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub startWriteTimer()
Const ProcName As String = "startWriteTimer"


On Error GoTo Err

If mWriteTimer Is Nothing Then Set mWriteTimer = CreateIntervalTimer(4000 + (Int(2000 * Rnd) + 1), ExpiryTimeUnitMilliseconds)
mWriteTimer.StartTimer

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub writePeriodData()
Const ProcName As String = "writePeriodData"
Dim tc As TaskController


On Error GoTo Err

If mTde Is Nothing Then Exit Sub
If Not mDataAdded Then Exit Sub

Set mTickWriter = New TickDataWriter
mTickWriter.Initialise mDB, mInstrument.Id, mInstrument.ShortName

If mUseAsynchronousWrites Then
    Set tc = mTickWriter.WriteDataAsync(mTde.encodingFormatIdentifier, _
                                        mTde.PeriodStartTime, _
                                        IIf(mTde.BasePrice = MaxDouble, 0#, mTde.BasePrice), _
                                        mTde.TickSize, _
                                        mTde.encodedData)
    tc.AddTaskCompletionListener Me
    Set mTickWriter = Nothing
Else
    mTickWriter.WriteData mTde.encodingFormatIdentifier, _
                                mTde.PeriodStartTime, _
                                IIf(mTde.BasePrice = MaxDouble, 0#, mTde.BasePrice), _
                                mTde.TickSize, _
                                mTde.encodedData
End If

mDataAdded = False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub


