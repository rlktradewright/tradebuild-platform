VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QTHistDataReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Implements TradeBuildSP.IBarDataReader
Implements LockConsumer

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const InitialRequestQueueCapacity As Long = 100

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mCommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer
Private mServiceConsumer As TradeBuildSP.IBarDataInputServiceConsumer
Private mDataConsumer As TradeBuildSP.IBarDataConsumer
Private mServiceProviderName As String

Private WithEvents mQTAPI As QTAPI
Attribute mQTAPI.VB_VarHelpID = -1
Private mServer As String
Private mPort As Long
Private mPassword As String
Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean
Private mProviderKey As String

Private mRequestInProgress As Boolean
Private mLockId As Long

Private mRequestId As Long

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

Private mBarDataSpecifier  As TradeBuildSP.BarDataSpecifier
Private mContractSpecifier As IContractSpecifier

Private mWaitingForBarData As Boolean
Private mBarRequested As Boolean

Private mLogLevel As LogLevels

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTimer = New IntervalTimer
mTimer.RepeatNotifications = False
mTimer.TimerIntervalMillisecs = 1

mPort = 7496
mConnectionRetryIntervalSecs = 10
mLogLevel = LogLevelLow

mRequestId = -1
End Sub

Private Sub Class_Terminate()
If Not mQTAPI Is Nothing Then gReleaseQTAPIInstance mQTAPI
Set mQTAPI = Nothing
Debug.Print "IBTWSSP.HistDataReader terminated"
End Sub

'================================================================================
' IBarDataReader Interface Members
'================================================================================

Private Property Get IBarDataReader_BarDataSpecifier( _
                            ) As TradeBuildSP.BarDataSpecifier
IBarDataReader_BarDataSpecifier = mBarDataSpecifier
End Property

Private Sub IBarDataReader_CancelFetch()
mQTAPI.cancelHistoricalBarData mRequestId, mLockId
End Sub

Private Sub IBarDataReader_FetchBars( _
                            barSpecifier As TradeBuildSP.BarDataSpecifier)

If mRequestInProgress Then err.Raise ErrorCodes.IllegalStateException, _
                                    "IBTWSSP.HistDataReader::FetchBars", _
                                    "Request already in progress"

mRequestInProgress = True
mWaitingForBarData = True
mBarDataSpecifier = barSpecifier

If mQTAPI Is Nothing Then
    Set mQTAPI = gGetQTAPIInstance(mServer, mPort, mPassword, mProviderKey, mConnectionRetryIntervalSecs, mkeepConnection)
    mQTAPI.CommonServiceConsumer = mCommonServiceConsumer
    mQTAPI.logLevel = mLogLevel
End If

If mQTAPI.connectionState = ConnConnected Then
    mServiceConsumer.Ready
    mRequestId = mQTAPI.requestHistoricalBarData(mBarDataSpecifier, Me, mLockId)
End If

End Sub

Private Sub IBarDataReader_FireNextBar()
If mWaitingForBarData Then
    mBarRequested = True
Else
    ProcessBar
End If
End Sub

Private Property Get IBarDataReader_NumberOfBars() As Long

End Property

Private Sub IBarDataReader_ReleaseDataStore()
If Not mQTAPI Is Nothing Then gReleaseQTAPIInstance mQTAPI
Set mQTAPI = Nothing
End Sub

Private Function IBarDataReader_Supports( _
                            ByVal capabilities As Long) As Boolean
IBarDataReader_Supports = False
End Function

'================================================================================
' LockConsumer Interface Members
'================================================================================

Private Sub LockConsumer_lockAcquired(ByVal lockId As Long)
If mLockId <> lockId Then err.Raise ErrorCodes.IllegalStateException, _
                                    "QTSP.QTHistDataReader::lockAcquired", _
                                    "Lock id acquired is " & lockId & ": expected " & mLockId
    
mRequestId = mQTAPI.requestHistoricalBarData(mBarDataSpecifier, Me, mLockId)
End Sub

Private Sub LockConsumer_waitingForLock(ByVal lockId As Long)
mLockId = lockId
End Sub

'================================================================================
' mqtapi Event Handlers
'================================================================================

Private Sub mQTAPI_Connected()
If mRequestInProgress Then
    mServiceConsumer.Ready
    mRequestId = mQTAPI.requestHistoricalBarData(mBarDataSpecifier, Me, mLockId)
End If
End Sub

Private Sub mQTAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
If Not retrying Then
    mServiceConsumer.error StandardSPErrorCodes.HDCantConnectDataSource, _
                            Description
End If
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let barDataInputServiceConsumer(ByVal value As IBarDataInputServiceConsumer)
Set mServiceConsumer = value
End Property

Friend Property Let CommonServiceConsumer(ByVal value As ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property

Friend Property Let ConnectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
End Property

Friend Property Let dataConsumer(ByVal value As IBarDataConsumer)
Set mDataConsumer = value
End Property

Public Property Let keepConnection(ByVal value As Boolean)
mkeepConnection = value
End Property

Friend Property Let logLevel(ByVal value As LogLevels)
mLogLevel = value
End Property

Public Property Let password(ByVal value As String)
mPassword = value
End Property

Friend Property Let port(ByVal value As Long)
mPort = value
End Property

Public Property Let providerKey(ByVal value As String)
mProviderKey = value
End Property

Friend Property Let server(ByVal value As String)
mServer = IIf(value = "", "127.0.0.1", value)
End Property

Friend Property Let ServiceProviderName(ByVal value As String)
mServiceProviderName = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub historicalDataAvailable()
mWaitingForBarData = False
If mBarRequested Then
    mBarRequested = False
    ProcessBar
End If
End Sub

Friend Sub historicalDataRequestFailed(ByVal reason As String)
mServiceConsumer.error StandardSPErrorCodes.HDRequestFailed, reason
End Sub

Friend Sub historicalDataRequestInvalid(ByVal reason As String)
mServiceConsumer.error StandardSPErrorCodes.HDRequestInvalid, reason
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub ProcessBar()
Dim lBar As TradeBuildSP.Bar

With lBar
    If Not mQTAPI.readHistoricalBar( _
                            mRequestId, _
                            .timestamp, _
                            .barType, _
                            .PeriodMinutes, _
                            .openPrice, _
                            .highPrice, _
                            .lowPrice, _
                            .closePrice, _
                            .volume, _
                            .tickVolume, _
                            .openInterest) _
    Then
        mServiceConsumer.BarDataComplete
        mRequestInProgress = False
        mRequestId = -1
        Exit Sub
    End If
End With
mDataConsumer.Bar lBar
End Sub




