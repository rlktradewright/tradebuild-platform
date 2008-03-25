VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@===============================================================================
' Description
'@===============================================================================
'
'   The first byte of each record has the following structure:
'
'    TickType (Bits 0-3)
'    SizeType (Bits 4-5)
'    NoTimestamp (bit 6)
'    NegativeTicks (bit 7)
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

Implements ITickfileWriter

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TBInfoBase26"
Private Const ModuleName As String = "TickfileWriter"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private WithEvents mTradingDB       As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1
Private mConnectionString           As String

Private mTickfileServiceProvider    As tickfileServiceProvider

Private mCommonServiceConsumer      As ICommonServiceConsumer
Private mServiceConsumer            As ITickfileOutputServiceConsumer
Private mServiceProviderName        As String

Private mContract                   As Contract
Private WithEvents mTickWriter      As TickDataWriter
Attribute mTickWriter.VB_VarHelpID = -1

Private mTde                        As TickDataEncoder

Private mTotalBytes                 As Long

Private mTickSize                   As Double
Private mOpeningTimestamp           As Date   ' timestamp at start of minute period

Private mInstrument                 As instrument
Private mUseSynchronousWrites       As Boolean

Private WithEvents mWriteTimer      As IntervalTimer
Attribute mWriteTimer.VB_VarHelpID = -1

Private mDataAdded                  As Boolean

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
Set mWriteTimer = CreateIntervalTimer(5000, ExpiryTimeUnitMilliseconds, 5000, True)
End Sub

Private Sub Class_Terminate()
Debug.Print "TBInfoBase.TickfileWriter terminated"
End Sub

'@===============================================================================
' ITickfileWriter Interface Members
'@===============================================================================

Private Sub ITickfileWriter_CloseOutputFile()
On Error GoTo Err
finish
mTickfileServiceProvider.dataWriterFinished Me
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "ITickfileWriter_CloseOutputFile", _
                Err.Description
End Sub

Private Property Get ITickfileWriter_Contract() As Contract
Set ITickfileWriter_Contract = mContract
End Property

Private Property Let ITickfileWriter_Contract(ByVal RHS As Contract)
Dim lInstrumentFactory As InstrumentFactory

On Error GoTo Err
If Not RHS Is Nothing Then
    Set mContract = RHS
End If

Set lInstrumentFactory = mTradingDB.InstrumentFactory
Set mInstrument = lInstrumentFactory.loadBySpecifier(mContract.specifier)

'?????????????????????????????????????????????????????
' consider inserting this instrument into the database
' if it doesn't exist
' ????????????????????????????????????????????????????
    
Debug.Assert Not mInstrument Is Nothing

mTickSize = mContract.ticksize

Exit Property
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "ITickfileWriter_Contract", _
                Err.Description

End Property

Private Property Get ITickfileWriter_FormatIdentifier() As String

End Property

Private Property Get ITickfileWriter_FormatName() As String

End Property

Private Property Get ITickfileWriter_FormatVersion() As Long

End Property

Private Sub ITickfileWriter_WriteAsk( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)
On Error GoTo Err
If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeAsk timestamp, price, size
mDataAdded = True
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "ITickfileWriter_WriteAsk", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeBid( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)
On Error GoTo Err
If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeBid timestamp, price, size
mDataAdded = True
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeBid", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeClose( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)
On Error GoTo Err
If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeClose timestamp, price
mDataAdded = True
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeClose", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeHigh( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)
On Error GoTo Err
If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeHigh timestamp, price
mDataAdded = True
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeHigh", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeLow( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)
On Error GoTo Err
If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeLow timestamp, price
mDataAdded = True
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeLow", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeMarketDepthData( _
                            ByVal timestamp As Date, _
                            ByVal position As Long, _
                            ByVal marketmaker As String, _
                            ByVal operation As Long, _
                            ByVal side As Long, _
                            ByVal price As Double, _
                            ByVal size As Long)

On Error GoTo Err
If price = 0# Then Exit Sub

checkIfNewPeriod timestamp
mTde.encodeMarketDepthData timestamp, position, marketmaker, operation, side, price, size
mDataAdded = True

Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeMarketDepthData", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeMarketDepthReset( _
                            ByVal timestamp As Date)

On Error GoTo Err
checkIfNewPeriod timestamp
mTde.encodeMarketDepthReset timestamp
mDataAdded = True

Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeMarketDepthReset", _
                Err.Description
End Sub

Private Sub ITickfileWriter_WriteOpenInterest(ByVal timestamp As Date, ByVal size As Long)
On Error GoTo Err
checkIfNewPeriod timestamp
mTde.encodeOpenInterest timestamp, size
mDataAdded = True
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "ITickfileWriter_WriteOpenInterest", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeTrade( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)
If price = 0# Then Exit Sub
checkIfNewPeriod timestamp
mTde.encodeTrade timestamp, price, size
mDataAdded = True
End Sub

Private Sub ITickfileWriter_writeVolume( _
                            ByVal timestamp As Date, _
                            ByVal size As Long)
On Error GoTo Err
checkIfNewPeriod timestamp
mTde.encodeVolume timestamp, size
mDataAdded = True
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeVolume", _
                Err.Description
End Sub

'@===============================================================================
' mTickWriter Event Handlers
'@===============================================================================

Private Sub mTickWriter_ConnectFailed( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
On Error GoTo Err
mServiceConsumer.Ready = False
mServiceConsumer.NotifyEvent StandardSPEventCodes.TFCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc, _
                        mTickfileServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "mTickWriter_ConnectFailed", _
                Err.Description
End Sub

Private Sub mTickWriter_Error( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
On Error GoTo Err
mCommonServiceConsumer.ServiceProviderError errorCode, _
                                    errorDesc, _
                                    mTickfileServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "mTickWriter_Error", _
                Err.Description
End Sub

'@===============================================================================
' mTradingDB Event Handlers
'@===============================================================================

Private Sub mTradingDB_ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
On Error GoTo Err
mServiceConsumer.Ready = False
mServiceConsumer.NotifyEvent StandardSPEventCodes.TFCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc, _
                        mTickfileServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "mTradingDB_ConnectFailed", _
                Err.Description
End Sub

Private Sub mTradingDB_ConnectSucceeded()
On Error GoTo Err

mServiceConsumer.NotifyEvent StandardSPEventCodes.TFConnectedToDataSource, _
                        "Connected to data source", _
                        mTickfileServiceProvider.handle

mServiceConsumer.Ready = True
mServiceConsumer.OutputTickfileCreated mConnectionString
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "mTradingDB_ConnectSucceeded", _
                Err.Description

End Sub
'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub finish()
writePeriodData
mWriteTimer.StopTimer
Set mWriteTimer = Nothing
mCommonServiceConsumer.NotifyListeners mServiceProviderName & " (" & mContract.specifier.localSymbol & ")" & vbCrLf & _
                            "    Total data written: " & mTotalBytes & " bytes", _
                            TradeBuildSP.StandardListenValueTypes.Log, _
                            Me
' we really ought to get rid of mWriter after a suitable delay, or after we
' know that all writes have been completed

End Sub

Friend Sub initialise( _
                ByVal Name As String, _
                ByVal connectionString As String, _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal tickfileOutputServiceConsumer As TradeBuildSP.ITickfileOutputServiceConsumer, _
                ByVal tickfileServiceProvider As TradeBuildSP.ITickfileServiceProvider, _
                ByVal useSynchronousWrites As Boolean)

mServiceProviderName = Name
Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = tickfileOutputServiceConsumer
Set mTickfileServiceProvider = tickfileServiceProvider
                
mUseSynchronousWrites = useSynchronousWrites

mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString)

End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub checkIfNewPeriod(ByVal timestamp As Date)
Dim barTime As Date
barTime = gTruncateTimeToMinute(timestamp)
If barTime = mOpeningTimestamp Then Exit Sub

If Not mTde Is Nothing Then
    writePeriodData
    mTotalBytes = mTotalBytes + mTde.encodedDataLength
End If

Set mTde = CreateTickEncoder(barTime, mTickSize)
mWriteTimer.StopTimer
mWriteTimer.StartTimer

mOpeningTimestamp = barTime

End Sub

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)

mCommonServiceConsumer.FatalServiceProviderError Number, _
                            Source, _
                            Description, _
                            mTickfileServiceProvider.handle

mTickfileServiceProvider.dataWriterFinished Me
End Sub

Private Sub writePeriodData()

If mTde Is Nothing Then Exit Sub
If Not mDataAdded Then Exit Sub

Set mTickWriter = mInstrument.CreateTickDataWriter

mTickWriter.WriteData mTde.encodingFormatIdentifier, _
                            mTde.periodStartTime, _
                            mTde.basePrice, _
                            mTde.ticksize, _
                            mTde.encodedData, _
                            Not mUseSynchronousWrites

mDataAdded = False

End Sub

Private Sub mWriteTimer_TimerExpired()
writePeriodData
End Sub
