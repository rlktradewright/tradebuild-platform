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

Private mTickfileServiceProvider As tickfileServiceProvider

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mServiceConsumer As ITickfileOutputServiceConsumer
Private mServiceProviderName As String

Private mContract As IContract
Private WithEvents mTickWriter As TickDataWriter
Attribute mTickWriter.VB_VarHelpID = -1

Private mData() As Byte
Private mDataIndex As Long

Private mTotalBytes As Long

Private mTickSize As Double
Private mTimestamp As Date          ' current timestamp
Private mBasePrice As Double        ' first price encountered at start of minute period
Private mOpeningTimestamp As Date   ' timestamp at start of minute period

Private mInstrument As instrument

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
ReDim mData(8191) As Byte
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

Private Property Get ITickfileWriter_Contract() As IContract
Set ITickfileWriter_Contract = mContract
End Property

Private Property Let ITickfileWriter_Contract(ByVal RHS As IContract)
Dim lInstrumentFactory As InstrumentFactory

On Error GoTo Err
If Not RHS Is Nothing Then
    Set mContract = RHS.Clone
End If

Set lInstrumentFactory = mTradingDB.InstrumentFactory
With mContract.Specifier
    Set mInstrument = lInstrumentFactory.loadBySpec(.localSymbol, _
                                .Symbol, _
                                .Exchange, _
                                gSecTypeToCategory(.secType), _
                                .CurrencyCode, _
                                Left$(.Expiry, 6), _
                                .Strike, _
                                gOptRightToString(.Right))
End With

'?????????????????????????????????????????????????????
' consider inserting this instrument into the database
' if it doesn't exist
' ????????????????????????????????????????????????????
    
Debug.Assert Not mInstrument Is Nothing

mTickSize = mContract.MinimumTick
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
WriteQuoteRecord TickTypes.Ask, timestamp, price, size
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
WriteQuoteRecord TickTypes.Bid, timestamp, price, size
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
WritePriceRecord TickTypes.closePrice, timestamp, price
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
WritePriceRecord TickTypes.highPrice, timestamp, price
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
WritePriceRecord TickTypes.lowPrice, timestamp, price
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
Dim typeByte As Byte
Dim sizeType As SizeTypes
Dim sideOperationPositionByte As Byte

On Error GoTo Err
If price = 0# Then Exit Sub

CheckIfNewPeriod timestamp

typeByte = TickTypes.marketDepth

sizeType = GetSizeType(size)
typeByte = typeByte Or (sizeType * SizeTypeShifter)

If timestamp = mTimestamp Then typeByte = typeByte Or NoTimestamp

If mBasePrice = 0# Then mBasePrice = price
If price < mBasePrice Then typeByte = typeByte Or NegativeTicks

WriteByte typeByte
WriteTimestamp timestamp
sideOperationPositionByte = CByte(position)
sideOperationPositionByte = sideOperationPositionByte Or CByte(operation * OperationShifter)
sideOperationPositionByte = sideOperationPositionByte Or CByte(side * SideShifter)
WriteByte sideOperationPositionByte

WritePrice price
WriteSize size, sizeType
WriteString marketmaker
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeMarketDepthData", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeMarketDepthReset( _
                            ByVal timestamp As Date)
Dim typeByte As Byte

On Error GoTo Err
typeByte = TickTypes.MarketDepthReset

If timestamp = mTimestamp Then typeByte = typeByte Or NoTimestamp

WriteByte typeByte
WriteTimestamp timestamp
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeMarketDepthReset", _
                Err.Description
End Sub

Private Sub ITickfileWriter_WriteOpenInterest(ByVal timestamp As Date, ByVal size As Long)
On Error GoTo Err
WriteSizeRecord TickTypes.openInterest, timestamp, size
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
WriteQuoteRecord TickTypes.Trade, timestamp, price, size
End Sub

Private Sub ITickfileWriter_writeVolume( _
                            ByVal timestamp As Date, _
                            ByVal size As Long)
On Error GoTo Err
WriteSizeRecord TickTypes.volume, timestamp, size
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
WritePeriodData
mCommonServiceConsumer.NotifyListeners "Total data written: " & mTotalBytes & " bytes", _
                            TradeBuildSP.StandardListenValueTypes.Log, _
                            Me
' we really ought to get rid of mWriter after a suitable delay, or after we
' know that all writes hve been completed

End Sub

Friend Sub initialise( _
                ByVal Name As String, _
                ByVal connectionString As String, _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal tickfileOutputServiceConsumer As TradeBuildSP.ITickfileOutputServiceConsumer, _
                ByVal tickfileServiceProvider As TradeBuildSP.ITickfileServiceProvider)

mServiceProviderName = Name
Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = tickfileOutputServiceConsumer
Set mTickfileServiceProvider = tickfileServiceProvider
                
mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString)

End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub CheckIfNewPeriod(ByVal timestamp As Date)
Dim barTime As Date
barTime = gTruncateTimeToMinute(timestamp)
If barTime = mOpeningTimestamp Then Exit Sub

WritePeriodData

ReDim mData(8191) As Byte
mDataIndex = 0
mOpeningTimestamp = barTime
mTimestamp = mOpeningTimestamp
mBasePrice = 0#
End Sub

Private Function GetSizeType(ByVal size As Long) As SizeTypes
If size < &H100 Then
    GetSizeType = SizeTypes.ShortSize
ElseIf size < &H10000 Then
    GetSizeType = SizeTypes.IntSize
Else
    GetSizeType = SizeTypes.LongSize
End If
End Function

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)
' discard any pending data
mDataIndex = 0
mCommonServiceConsumer.FatalServiceProviderError Number, _
                            Source, _
                            Description, _
                            mTickfileServiceProvider.handle

mTickfileServiceProvider.dataWriterFinished Me
End Sub

Private Sub WriteByte(ByVal theByte As Byte)
If mDataIndex > UBound(mData) Then
    ReDim Preserve mData(UBound(mData) + 8192) As Byte
End If
mData(mDataIndex) = theByte
mDataIndex = mDataIndex + 1
End Sub

Private Sub WriteInt(ByVal theInt As Integer)
WriteByte CByte(theInt And &HFF)
WriteByte CByte(Int(theInt / &H100) And &HFF)
End Sub

Private Sub WriteLong(ByVal theLong As Long)
WriteByte CByte(theLong And &HFF)
WriteByte CByte(Int(theLong / &H100) And &HFF)
WriteByte CByte(Int(theLong / &H10000) And &HFF)
WriteByte CByte(Int(theLong / &H1000000) And &HFF)
End Sub

Private Sub WritePeriodData()

If mDataIndex = 0 Then Exit Sub

ReDim Preserve mData(mDataIndex - 1) As Byte

Set mTickWriter = mInstrument.CreateTickDataWriter

mTickWriter.WriteData TickfileFormatTradeBuildSQL, _
                            mOpeningTimestamp, _
                            mBasePrice, _
                            mData
mTotalBytes = mTotalBytes + mDataIndex
End Sub

Private Sub WritePrice(ByVal price As Double)
Dim numticks As Integer
If price = mBasePrice Then
    WriteByte &H0
Else
    numticks = Abs(price - mBasePrice) / mTickSize
    If numticks <= 127 Then
        WriteByte CByte(Abs(price - mBasePrice) / mTickSize)
    Else
        ' won't fit in 7 bits. Write it out as an int value, with
        ' the high-order byte written first and bit 7 set
        WriteByte CByte(Int(numticks / &H100) And &H7F) Or &H80
        WriteByte CByte(numticks And &HFF)
    End If
End If
End Sub

Private Sub WritePriceRecord( _
                            ByVal tickType As TickTypes, _
                            ByVal timestamp As Date, _
                            ByVal price As Double)
Dim typeByte As Byte

CheckIfNewPeriod timestamp

typeByte = tickType

If timestamp = mTimestamp Then typeByte = typeByte Or NoTimestamp

If mBasePrice = 0# Then mBasePrice = price
If price < mBasePrice Then typeByte = typeByte Or NegativeTicks

WriteByte typeByte
WriteTimestamp timestamp
WritePrice price
End Sub

Private Sub WriteQuoteRecord( _
                            ByVal tickType As TickTypes, _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)
Dim typeByte As Byte
Dim sizeType As SizeTypes

CheckIfNewPeriod timestamp

typeByte = tickType

sizeType = GetSizeType(size)
typeByte = typeByte Or (sizeType * SizeTypeShifter)

If timestamp = mTimestamp Then typeByte = typeByte Or NoTimestamp

If mBasePrice = 0# Then mBasePrice = price
If price < mBasePrice Then typeByte = typeByte Or NegativeTicks

WriteByte typeByte

WriteTimestamp timestamp

WritePrice price

WriteSize size, sizeType
End Sub

Private Sub WriteSize(ByVal size As Long, ByVal sizeType As SizeTypes)
Select Case sizeType
Case SizeTypes.ShortSize
    Debug.Assert size < &H100&
    WriteByte CByte(size)
Case SizeTypes.IntSize
    Debug.Assert size < &H10000
    If size < &H8000& Then
        WriteInt size
    Else
        WriteInt CInt(size - &H10000)
    End If
Case SizeTypes.LongSize
    WriteLong size
End Select
End Sub

Private Sub WriteSizeRecord( _
                            ByVal tickType As TickTypes, _
                            ByVal timestamp As Date, _
                            ByVal size As Long)
Dim typeByte As Byte
Dim sizeType As SizeTypes

CheckIfNewPeriod timestamp

typeByte = tickType

sizeType = GetSizeType(size)
typeByte = typeByte Or (sizeType * SizeTypeShifter)

If timestamp = mTimestamp Then typeByte = typeByte Or NoTimestamp

WriteByte typeByte
WriteTimestamp timestamp
WriteSize size, sizeType
End Sub

Private Sub WriteString(ByVal theString As String)
Dim ar() As Byte
Dim i As Long

WriteByte CByte(Len(theString))

ar = theString
For i = 0 To UBound(ar)
    WriteByte ar(i)
Next
End Sub

Private Sub WriteTimestamp(ByVal timestamp As Date)
Dim diff As Long
If timestamp <> mTimestamp Then
    diff = (timestamp - mOpeningTimestamp) * 86400 * 1000
    Debug.Assert diff < &H10000
    If diff < &H8000& Then
        WriteInt diff
    Else
        WriteInt CInt(diff - &H10000)
    End If
    mTimestamp = timestamp
End If
End Sub

