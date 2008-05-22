VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QTTickfileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements ITickfileReader
Implements LockConsumer

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const InitialTickQueueCapacity As Long = 1000

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTickfileServiceProvider As QTTickfileServiceProvider

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mServiceConsumer As ITickfileInputServiceConsumer
Private mDataConsumer As IStreamingDataConsumer

Private mName As String

Private mElapsedTimer As ElapsedTimer

Private WithEvents mQTAPI As QTAPI
Attribute mQTAPI.VB_VarHelpID = -1
Private mPort As Long
Private mServer As String
Private mPassword As String

Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean
Private mProviderKey As String

Private mLockId As Long

Private mTickQueue As Queue
Private mNumPlayedTicks As Long

Private mAllTicksReceived As Boolean
Private mTickRequestPending As Boolean

Private mReplayProgressEventFrequency As Long

Private mInputTickfileSpecifier As TradeBuildSP.ITickfileSpecifier
Private mTickRequestId As Long

Private mLogLevel As LogLevels

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

Set mTickQueue = New Queue
mTickQueue.initialCapacity = InitialTickQueueCapacity

Set mElapsedTimer = New ElapsedTimer

End Sub

Private Sub Class_Terminate()
Debug.Print "QTSP.QTTickFileReader terminated"
End Sub

'================================================================================
' ITickfileReader Interface Members
'================================================================================

Private Sub ITickfileReader_CloseInputFile()
On Error GoTo err
finish
mTickfileServiceProvider.dataReaderFinished Me
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTTickfileReader" & "::" & "ITickfileReader_CloseInputFile", _
                err.Description
End Sub

Private Property Let ITickfileReader_Contract(ByVal RHS As Contract)
End Property

Private Property Get ITickfileReader_Contract() As Contract
Set ITickfileReader_Contract = mInputTickfileSpecifier.Contract
End Property

Private Property Get ITickfileReader_ContractSpecifier() As ContractSpecifier
Set ITickfileReader_ContractSpecifier = mInputTickfileSpecifier.Contract.specifier
End Property

Private Sub ITickfileReader_FireNextTick()
On Error GoTo err
FireNextTick
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTTickfileReader" & "::" & "ITickfileReader_FireNextTick", _
                err.Description
End Sub

Private Property Get ITickfileReader_FirstTickFileTime() As Date
ITickfileReader_FirstTickFileTime = FirstTickFileTime
End Property

Private Property Get ITickfileReader_InputTickfileSpecifier() As TradeBuildSP.ITickfileSpecifier
Set ITickfileReader_InputTickfileSpecifier = InputTickfileSpecifier
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
On Error GoTo err
Set mInputTickfileSpecifier = tickfileSpecifier

If mQTAPI Is Nothing Then
    Set mQTAPI = gGetQTAPIInstance(mServer, mPort, mPassword, mProviderKey, mConnectionRetryIntervalSecs, mkeepConnection)
    mQTAPI.CommonServiceConsumer = mCommonServiceConsumer
    mQTAPI.logLevel = mLogLevel
End If

If mQTAPI.connectionState = ConnConnected Then
    mServiceConsumer.Ready mInputTickfileSpecifier, 0
    mTickRequestId = mQTAPI.requestHistoricalTickData(mInputTickfileSpecifier, Me, mLockId)
End If
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTTickfileReader" & "::" & "ITickfileReader_ReplayTickfile", _
                err.Description

End Sub

Private Function ITickfileReader_Supports( _
                            ByVal capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
ITickfileReader_Supports = gSupports(capabilities, FormatIdentifier)
End Function

Private Property Get ITickfileReader_tickfileSizeBytes() As Long
ITickfileReader_tickfileSizeBytes = 0
End Property

'================================================================================
' LockConsumer Interface Members
'================================================================================

Private Sub LockConsumer_lockAcquired(ByVal lockId As Long)
On Error GoTo err
If mLockId <> lockId Then err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "QTSP.QTTickfileReader::lockAcquired", _
                                    "Lock id acquired is " & lockId & ": expected " & mLockId
    
    mTickRequestId = mQTAPI.requestHistoricalTickData(mInputTickfileSpecifier, Me, mLockId)
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTTickfileReader" & "::" & "LockConsumer_lockAcquired", _
                err.Description
End Sub

Private Sub LockConsumer_waitingForLock(ByVal lockId As Long)
mLockId = lockId
End Sub

'================================================================================
' mQTAPI Event Handlers
'================================================================================

Private Sub mQTAPI_Connected()
On Error GoTo err
mServiceConsumer.NotifyEvent StandardSPEventCodes.TFConnectedToDataSource, _
                        "Connected to data source", _
                        mTickfileServiceProvider.handle

mServiceConsumer.Ready mInputTickfileSpecifier, 0
mTickRequestId = mQTAPI.requestHistoricalTickData(mInputTickfileSpecifier, Me, mLockId)
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTTickfileReader" & "::" & "mQTAPI_Connected", _
                err.Description
End Sub

Private Sub mQTAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
On Error GoTo err
If retrying Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFRetryConnectDataSource, _
                            Description, _
                            mTickfileServiceProvider.handle
Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFCantConnectDataSource, _
                            Description, _
                            mTickfileServiceProvider.handle
End If
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTTickfileReader" & "::" & "mQTAPI_ConnectFailed", _
                err.Description
End Sub

Private Sub mQTAPI_ConnectionClosed(ByVal reconnecting As Boolean)
On Error GoTo err
If reconnecting Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFReConnectingDataSource, _
                        "Lost connection to TWS (reconnecting)", _
                        mTickfileServiceProvider.handle
Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFLostConnectionToDataSource, _
                        "Lost connection to TWS", _
                        mTickfileServiceProvider.handle
End If
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTTickfileReader" & "::" & "mQTAPI_ConnectionClosed", _
                err.Description
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let CommonServiceConsumer(ByVal RHS As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = RHS
End Property


Friend Property Let ConnectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
End Property

Friend Property Get ConnectionRetryIntervalSecs() As Long
ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
End Property

Friend Property Get FirstTickFileTime() As Date
' not supported
End Property

Friend Property Get InputTickfileSpecifier() As TradeBuildSP.ITickfileSpecifier
Set InputTickfileSpecifier = mInputTickfileSpecifier
End Property

Public Property Let keepConnection(ByVal value As Boolean)
mkeepConnection = value
End Property

Friend Property Let logLevel(ByVal value As LogLevels)
mLogLevel = value
End Property

Friend Property Let Name(ByVal value As String)
mName = value
End Property

Public Property Let password(ByVal value As String)
mPassword = value
End Property

Friend Property Let port(ByVal value As Long)
mPort = value
End Property

Friend Property Let providerKey(ByVal value As String)
mProviderKey = value
End Property

Friend Property Let server(ByVal value As String)
mServer = IIf(value = "", "127.0.0.1", value)
End Property

Friend Property Let StreamingDataConsumer(ByVal RHS As IStreamingDataConsumer)
Set mDataConsumer = RHS
End Property

Friend Property Let TickfileInputServiceConsumer(ByVal RHS As ITickfileInputServiceConsumer)
Set mServiceConsumer = RHS
End Property

Friend Property Let tickfileServiceProvider( _
                ByVal value As QTTickfileServiceProvider)
Set mTickfileServiceProvider = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub finish()
If Not mQTAPI Is Nothing Then
    mQTAPI.cancelHistoricalTickData mTickRequestId, mLockId
    gReleaseQTAPIInstance mQTAPI
    Set mQTAPI = Nothing
End If
End Sub

Friend Sub historicalTick( _
                ByRef pTick As Tick)
mTickQueue.addToEnd pTick
'If mTickQueue.size Mod 10 = 0 Then
'    Debug.Print "QTSP.QTTickfileReader: queue length: " & mTickQueue.size
'End If
If mTickRequestPending Then FireNextTick
End Sub

Friend Sub historicalTickDataComplete()
mAllTicksReceived = True
If mTickRequestPending Then FireNextTick
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub FireNextTick()
Dim lTick As Tick

If mTickQueue.size = 0 Then
    If mAllTicksReceived Then
        mServiceConsumer.TickfileReplayComplete
        Exit Sub
    Else
        mTickRequestPending = True
        Exit Sub
    End If
End If

mTickRequestPending = False

lTick = mTickQueue.removeFromFront

With lTick
    mNumPlayedTicks = mNumPlayedTicks + 1
    
    If mNumPlayedTicks Mod mReplayProgressEventFrequency = 0 Then
        mServiceConsumer.ReplayProgress .timestamp, _
                                mNumPlayedTicks, _
                                0#
    End If

    Select Case .TickType
    Case TickTypes.bid
        mDataConsumer.bid .timestamp, .Price, .size
    Case TickTypes.ask
        mDataConsumer.ask .timestamp, .Price, .size
    Case TickTypes.last
        mDataConsumer.Trade .timestamp, .Price, .size
    Case TickTypes.volume
        mDataConsumer.volume .timestamp, .size
    Case TickTypes.prevClose
        mDataConsumer.PreviousClose .timestamp, .Price
    Case TickTypes.high
        mDataConsumer.high .timestamp, .Price
    Case TickTypes.low
        mDataConsumer.low .timestamp, .Price
    Case TickTypes.openInterest
        mDataConsumer.openInterest .timestamp, .size
    End Select
End With

End Sub

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

Private Sub logMessage( _
                ByVal message As String, _
                ByVal logLevel As LogLevels)
If Not gLogger.isLoggable(logLevel) Then Exit Sub
gLogger.Log logLevel, _
            message, _
            Now
End Sub


