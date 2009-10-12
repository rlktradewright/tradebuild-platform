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

Private Const ProjectName                   As String = "TradingDO26"
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
                ev As TWUtilities30.TaskCompletionEvent)
If ev.cancelled Then
ElseIf ev.errorNumber <> 0 Then
    RaiseEvent Error(ev.errorNumber, ev.errorMessage)
Else
    If gLogger.IsLoggable(LogLevelMediumDetail) Then
        gLogger.Log LogLevelMediumDetail, _
                    "TickWriter: successfully wrote data for " & mInstrument.shortName & ": period " & FormatTimestamp(CDate(ev.cookie), TimestampDateAndTimeISO8601)
    End If
End If
End Sub

'@================================================================================
' mTickWriter Event Handlers
'@================================================================================

Private Sub mTickWriter_ConnectFailed(ByVal errorCode As Long, ByVal errorDesc As String)
Const ProcName As String = "mTickWriter_ConnectFailed"
Dim failpoint As Long
On Error GoTo Err

RaiseEvent ConnectFailed(errorCode, errorDesc)
Set mTickWriter = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=True, pProcedureName:=ProcName, pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
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
Dim failpoint As Long
On Error GoTo Err

writePeriodData
startWriteTimer

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=True, pProcedureName:=ProcName, pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get bytesWritten() As Long
bytesWritten = mTotalBytes
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal db As TradingDB, _
                ByVal pInstrument As instrument, _
                ByVal useAsynchronousWrites As Boolean)
Const ProcName As String = "initialise"
Dim failpoint As Long
On Error GoTo Err

Set mDB = db
Set mInstrument = pInstrument
mTickSize = mInstrument.tickSize
mUseAsynchronousWrites = useAsynchronousWrites
startWriteTimer

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub finish()
Dim failpoint As Long
On Error GoTo Err

writePeriodData

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="finish", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub writeAsk( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)
Dim failpoint As Long
On Error GoTo Err

If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeAsk timestamp, price, size
mDataAdded = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="writeAsk", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub writeBid( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)
Dim failpoint As Long
On Error GoTo Err

If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeBid timestamp, price, size
mDataAdded = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="writeBid", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub writeClose( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Dim failpoint As Long
On Error GoTo Err

If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeClose timestamp, price
mDataAdded = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="writeClose", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub writeHigh( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Dim failpoint As Long
On Error GoTo Err

If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeHigh timestamp, price
mDataAdded = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="writeHigh", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub writeLow( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Dim failpoint As Long
On Error GoTo Err

If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeLow timestamp, price
mDataAdded = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="writeLow", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub writeMarketDepthData( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal marketmaker As String, _
                ByVal operation As Long, _
                ByVal side As Long, _
                ByVal price As Double, _
                ByVal size As Long)
Dim failpoint As Long
On Error GoTo Err

If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeMarketDepthData timestamp, position, marketmaker, operation, side, price, size
mDataAdded = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="writeMarketDepthData", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub writeMarketDepthReset( _
                ByVal timestamp As Date)
Dim failpoint As Long
On Error GoTo Err

checkIfNewPeriod timestamp
mTde.encodeMarketDepthReset timestamp
mDataAdded = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="writeMarketDepthReset", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub writeOpen( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Dim failpoint As Long
On Error GoTo Err

checkIfNewPeriod timestamp
mTde.encodeOpen timestamp, price
mDataAdded = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="writeOpen", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub writeOpenInterest( _
                ByVal timestamp As Date, _
                ByVal size As Long)
Dim failpoint As Long
On Error GoTo Err

checkIfNewPeriod timestamp
mTde.encodeOpenInterest timestamp, size
mDataAdded = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="writeOpenInterest", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub writeTick( _
                ByRef tick As GenericTick)
Dim failpoint As Long
On Error GoTo Err

With tick
    Select Case .tickType
    Case TickTypes.TickTypeAsk
        writeAsk .timestamp, .price, .size
    Case TickTypes.TickTypeBid
        writeBid .timestamp, .price, .size
    Case TickTypes.TickTypeClosePrice
        writeClose .timestamp, .price
    Case TickTypes.TickTypeHighPrice
        writeHigh .timestamp, .price
    Case TickTypes.TickTypeLowPrice
        writeLow .timestamp, .price
    Case TickTypes.TickTypeMarketDepth
        writeMarketDepthData .timestamp, .position, .marketmaker, .operation, .side, .price, .size
    Case TickTypes.TickTypeMarketDepthReset
        writeMarketDepthReset .timestamp
    Case TickTypes.TickTypeOpenInterest
        writeOpenInterest .timestamp, .size
    Case TickTypes.TickTypeOpenPrice
        writeOpen .timestamp, .price
    Case TickTypes.TickTypeTrade
        writeTrade .timestamp, .price, .size
    Case TickTypes.TickTypeVolume
        writeVolume .timestamp, .size
    End Select
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="writeTick", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub writeTrade( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)
Dim failpoint As Long
On Error GoTo Err

If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeTrade timestamp, price, size
mDataAdded = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="writeTrade", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub writeVolume( _
                ByVal timestamp As Date, _
                ByVal size As Long)
Dim failpoint As Long
On Error GoTo Err

checkIfNewPeriod timestamp
mTde.encodeVolume timestamp, size
mDataAdded = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="writeVolume", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkIfNewPeriod(ByVal timestamp As Date)
Dim barTime As Date
Dim failpoint As Long
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="checkIfNewPeriod", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub startWriteTimer()
Dim failpoint As Long
On Error GoTo Err

If mWriteTimer Is Nothing Then Set mWriteTimer = CreateIntervalTimer(4000 + (Int(2000 * Rnd) + 1), ExpiryTimeUnitMilliseconds)
mWriteTimer.StartTimer

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="startWriteTimer", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub writePeriodData()
Dim tc As TaskController

Dim failpoint As Long
On Error GoTo Err

If mTde Is Nothing Then Exit Sub
If Not mDataAdded Then Exit Sub

Set mTickWriter = New TickDataWriter
mTickWriter.initialise mDB, mInstrument.id, mInstrument.shortName

If mUseAsynchronousWrites Then
    Set tc = mTickWriter.WriteDataAsync(mTde.encodingFormatIdentifier, _
                                        mTde.PeriodStartTime, _
                                        IIf(mTde.BasePrice = MaxDouble, 0#, mTde.BasePrice), _
                                        mTde.tickSize, _
                                        mTde.encodedData)
    tc.AddTaskCompletionListener Me
    Set mTickWriter = Nothing
Else
    mTickWriter.WriteData mTde.encodingFormatIdentifier, _
                                mTde.PeriodStartTime, _
                                IIf(mTde.BasePrice = MaxDouble, 0#, mTde.BasePrice), _
                                mTde.tickSize, _
                                mTde.encodedData
End If

mDataAdded = False

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="writePeriodData", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub


