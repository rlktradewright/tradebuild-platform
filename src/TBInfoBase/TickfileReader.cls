VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Implements ITickfileReader

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TBInfoBase26"
Private Const ModuleName As String = "TickfileReader"

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

Private mCapabilities               As Long

Private mTickfileServiceProvider    As tickfileServiceProvider

Private mCommonServiceConsumer      As ICommonServiceConsumer
Private mServiceConsumer            As ITickfileInputServiceConsumer
Private mDataConsumer               As IStreamingDataConsumer
Private mServiceProviderName        As String

Private mReadyNotified              As Boolean
Private mFetchInProgress            As Boolean
Private mTickRequested              As Boolean

Private mReplayProgressEventFrequency As Long

Private mContract                   As Contract
Private mContractSpecifier          As ContractSpecifier
Private mInstrument                 As instrument
Attribute mInstrument.VB_VarHelpID = -1
Private WithEvents mTickReader      As TickDataReader
Attribute mTickReader.VB_VarHelpID = -1

Private mTdd                        As TickDataDecoder

Private mRecordsRead                As Long

Private mInputTickfileSpecifier     As TradeBuildSP.ITickfileSpecifier

Private mFirstTickFileTime          As Date

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
mReplayProgressEventFrequency = 10
End Sub

Private Sub Class_Terminate()
Debug.Print "TBInfoBase.TickfileReader terminated"
End Sub

'@===============================================================================
' ITickfileReader Interface Members
'@===============================================================================

Private Sub ITickfileReader_CloseInputFile()
On Error GoTo Err
finish
mTickfileServiceProvider.dataReaderFinished Me
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "ITickfileReader_CloseInputFile", _
                Err.Description
End Sub

Private Property Let ITickfileReader_Contract(ByVal RHS As Contract)
Dim st As SessionTimes

Set mContract = RHS

With mInputTickfileSpecifier
    If .EntireSession Then
        st = GetSessionTimes(.FromDate, _
                    mContract.sessionStartTime, _
                    mContract.sessionEndTime)
    .FromDate = st.startTime
    .ToDate = st.endTime
    End If
    
    ' the following is a kludge. We need to modify the service provider handshake so
    ' that so there is something after the 'ready' which tells the tickfile manager
    ' to start the replay, because we haven't got the data at this point (which is after
    ' 'ready' has been sent. Then we can get the firsttickfiletime from the data)
    mFirstTickFileTime = .FromDate
    
    mTickReader.FetchData gTruncateTimeToMinute(.FromDate), _
                            gTruncateTimeToNextMinute(.ToDate)
End With

mFetchInProgress = True

End Property

Private Property Get ITickfileReader_Contract() As Contract
Set ITickfileReader_Contract = mContract
End Property

Private Property Get ITickfileReader_ContractSpecifier() As ContractSpecifier
Set ITickfileReader_ContractSpecifier = mContractSpecifier
End Property

Private Sub ITickfileReader_FireNextTick()

If mFetchInProgress Then
    mTickRequested = True
Else
    processTickRequest
End If
End Sub

Private Property Get ITickfileReader_FirstTickFileTime() As Date
ITickfileReader_FirstTickFileTime = mFirstTickFileTime
End Property

Private Property Get ITickfileReader_InputTickfileSpecifier() As TradeBuildSP.ITickfileSpecifier
Set ITickfileReader_InputTickfileSpecifier = mInputTickfileSpecifier
End Property

Private Property Let ITickfileReader_ReplayProgressEventFrequency(ByVal RHS As Long)
If RHS = 0 Then
    mReplayProgressEventFrequency = 1
Else
    mReplayProgressEventFrequency = RHS
End If
End Property

Private Property Get ITickfileReader_ReplayProgressEventFrequency() As Long
ITickfileReader_ReplayProgressEventFrequency = mReplayProgressEventFrequency
End Property

Private Sub ITickfileReader_ReplayTickfile( _
                ByVal tickfileSpecifier As TradeBuildSP.ITickfileSpecifier)
Dim lInstrumentFactory As InstrumentFactory

On Error GoTo Err
Set mInputTickfileSpecifier = tickfileSpecifier

' we use the supplied contract specifier to retrieve the relevant Instrument
' object from the database so that we can create a TickDataReader object from it.

' Note that a 'canonical contract' object will be notified later via
' ITickfileReader_Contract, and it is that one that we will use for
' session times etc.

Set lInstrumentFactory = mTradingDB.InstrumentFactory

Set mInstrument = lInstrumentFactory.loadBySpecifier(mInputTickfileSpecifier.Contract.specifier)
 
If mInstrument Is Nothing Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFContractSpecifierInvalid, _
                            "The specified contract does not exist", _
                            mTickfileServiceProvider.handle
    Exit Sub
End If

Set mContract = mInputTickfileSpecifier.Contract
Set mTickReader = mInstrument.CreateTickDataReader
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "ITickfileReader_ReplayTickfile", _
                Err.Description

End Sub

Private Function ITickfileReader_Supports( _
                            ByVal capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
ITickfileReader_Supports = capabilities & mCapabilities
End Function

Private Property Get ITickfileReader_tickfileSizeBytes() As Long
ITickfileReader_tickfileSizeBytes = 0
End Property

'@===============================================================================
' mTickReader Event Handlers
'@===============================================================================

Private Sub mTickReader_ConnectFailed( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
On Error GoTo Err
mServiceConsumer.NotifyEvent StandardSPEventCodes.TFCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc, _
                        mTickfileServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "mTickReader_ConnectFailed", _
                Err.Description
End Sub

Private Sub mTickReader_DataFetched()

On Error GoTo Err

mFetchInProgress = False

If Not getNextPeriodData Then
    Exit Sub
End If

If Not mReadyNotified Then
    ' we should get the firsttickfiletime from the data, but it is already too late
    ' for that - see the comment in ITickfileReader_Contract
    'mFirstTickFileTime = mCurrentTimestamp
    mReadyNotified = True
    mServiceConsumer.Ready mInputTickfileSpecifier, 0
End If

If mTickRequested Then
    mTickRequested = False
    processTickRequest
End If
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "mTickReader_DataFetched", _
                Err.Description
End Sub

Private Sub mTickReader_Error( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
On Error GoTo Err
mCommonServiceConsumer.ServiceProviderError errorCode, _
                                    errorDesc, _
                                    mTickfileServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "mTickReader_Error", _
                Err.Description
End Sub

Private Sub mTickReader_Ready()
On Error GoTo Err

mServiceConsumer.NotifyEvent StandardSPEventCodes.TFConnectedToDataSource, _
                        "Connected to data source", _
                        mTickfileServiceProvider.handle

mReadyNotified = True
mServiceConsumer.Ready mInputTickfileSpecifier, 0

Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "mTickReader_Ready", _
                Err.Description

End Sub

'@===============================================================================
' mTradingDB Event Handlers
'@===============================================================================

Private Sub mTradingDB_ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
On Error GoTo Err
mServiceConsumer.NotifyEvent StandardSPEventCodes.TFCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc, _
                        mTickfileServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "mTradingDB_ConnectFailed", _
                Err.Description
End Sub

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub finish()
Set mTickReader = Nothing
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing
End Sub

Friend Sub initialise( _
                ByVal Name As String, _
                ByVal connectionString As String, _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal tickfileInputServiceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                ByVal tickfileServiceProvider As TradeBuildSP.ITickfileServiceProvider, _
                ByVal DataConsumer As TradeBuildSP.IStreamingDataConsumer, _
                ByVal capabilities As Long)

mServiceProviderName = Name
Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = tickfileInputServiceConsumer
Set mTickfileServiceProvider = tickfileServiceProvider
Set mDataConsumer = DataConsumer
                
mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString)

mCapabilities = capabilities

End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Function getNextPeriodData() As Boolean
Dim periodStartTime As Date
Dim format As String
Dim basePrice As Double
Dim ticksize As Double
Dim data() As Byte

If Not mTickReader.ReadData(periodStartTime, format, basePrice, ticksize, data) Then
    If mReadyNotified Then
        mServiceConsumer.TickfileReplayComplete
    Else
        mServiceConsumer.NotifyEvent StandardSPEventCodes.TFTickfileEmpty, _
                                "No data for this specifier", _
                                mTickfileServiceProvider.handle
    End If
    Exit Function
End If

Set mTdd = CreateTickDecoder(periodStartTime, ticksize, basePrice, data, format)
getNextPeriodData = True
End Function

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError Number, _
                                Source, _
                                Description, _
                                mTickfileServiceProvider.handle

finish
mTickfileServiceProvider.dataReaderFinished Me
End Sub

Private Sub processTickRequest()
Dim tick As GenericTick

On Error GoTo Err


Do
    If Not mTdd.getNextTick(tick) Then
        If Not getNextPeriodData Then Exit Sub
    End If
    If tick.timestamp >= mInputTickfileSpecifier.FromDate Then
        Exit Do
    End If
Loop
    
If tick.timestamp >= mInputTickfileSpecifier.ToDate Then
    mServiceConsumer.TickfileReplayComplete
    Exit Sub
End If

mRecordsRead = mRecordsRead + 1

Select Case tick.tickType
Case TickTypeBid
    mDataConsumer.Bid tick.timestamp, tick.price, tick.size
Case TickTypeAsk
    mDataConsumer.Ask tick.timestamp, tick.price, tick.size
Case TickTypeClosePrice
    mDataConsumer.PreviousClose tick.timestamp, tick.price
Case TickTypeHighPrice
    mDataConsumer.High tick.timestamp, tick.price
Case TickTypeLowPrice
    mDataConsumer.Low tick.timestamp, tick.price
Case TickTypeMarketDepth
    If tick.marketmaker = "" Then
        mDataConsumer.UpdateMktDepth tick.timestamp, _
                                    tick.position, _
                                    tick.operation, _
                                    tick.side, _
                                    tick.price, _
                                    tick.size
    Else
        mDataConsumer.UpdateMktDepthL2 tick.timestamp, _
                                    tick.position, _
                                    tick.marketmaker, _
                                    tick.operation, _
                                    tick.side, _
                                    tick.price, _
                                    tick.size
    End If
Case TickTypeMarketDepthReset
    mDataConsumer.ResetMarketDepth tick.timestamp, False
Case TickTypeTrade
    mDataConsumer.Trade tick.timestamp, tick.price, tick.size
Case TickTypeVolume
    mDataConsumer.Volume tick.timestamp, tick.size
Case TickTypeOpenInterest
    mDataConsumer.OpenInterest tick.timestamp, tick.size
End Select

If mRecordsRead Mod mReplayProgressEventFrequency = 0 Then
    mServiceConsumer.ReplayProgress tick.timestamp, _
                            mRecordsRead, _
                            100 * (tick.timestamp - mInputTickfileSpecifier.FromDate) / (mInputTickfileSpecifier.ToDate - mInputTickfileSpecifier.FromDate)
End If

Exit Sub

Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "ITickfileReader_processTickRequest", _
                Err.Description

End Sub


