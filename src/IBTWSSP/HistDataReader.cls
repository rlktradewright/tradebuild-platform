VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataReader"
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

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

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

Private WithEvents mTWSAPI As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1
Private mClientID As Long
Private mServer As String
Private mPort As Long
Private mProviderKey As String
Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean

Private mRequestInProgress As Boolean

Private mRequestId As Long

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

Private mBarDataSpecifier  As TradeBuildSP.BarDataSpecifier

Private mWaitingForBarData As Boolean
Private mBarRequested As Boolean

Private mLogLevel As LogLevels
Private mTWSLogLevel As TWSLogLevels

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTimer = New IntervalTimer
mTimer.RepeatNotifications = False
mTimer.TimerIntervalMillisecs = 1

mClientID = -1
mPort = 7496
mConnectionRetryIntervalSecs = 10
mLogLevel = LogLevelLow
mTWSLogLevel = TWSLogLevelError

mRequestId = -1
End Sub

Private Sub Class_Terminate()
If Not mTWSAPI Is Nothing Then gReleaseTWSAPIInstance mTWSAPI
Set mTWSAPI = Nothing
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
If mRequestId <> -1 Then mTWSAPI.CancelHistoricalData mRequestId
End Sub

Private Sub IBarDataReader_FetchBars( _
                            barSpecifier As TradeBuildSP.BarDataSpecifier)

If mRequestInProgress Then err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "IBTWSSP.HistDataReader::FetchBars", _
                                    "Request already in progress"

mRequestInProgress = True
mWaitingForBarData = True
mBarDataSpecifier = barSpecifier

If mTWSAPI Is Nothing Then
    If mClientID = -1 Then mClientID = Int(&H7FFFFFFF * Rnd) + 1
    Set mTWSAPI = gGetTWSAPIInstance(mServer, mPort, mClientID, mProviderKey, mConnectionRetryIntervalSecs, mkeepConnection)
    mTWSAPI.logLevel = mLogLevel
    mTWSAPI.TWSLogLevel = mTWSLogLevel
End If

If mTWSAPI.connectionState = ConnConnected Then
    mServiceConsumer.Ready
    mRequestId = mTWSAPI.RequestHistoricalData(mBarDataSpecifier, Me)
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
If Not mTWSAPI Is Nothing Then gReleaseTWSAPIInstance mTWSAPI
Set mTWSAPI = Nothing
End Sub

Private Function IBarDataReader_Supports( _
                            ByVal capabilities As Long) As Boolean
IBarDataReader_Supports = gHistDataSupports(capabilities)
End Function

'================================================================================
' mTWSAPI Event Handlers
'================================================================================

Private Sub mTWSAPI_Connected()
If mRequestInProgress Then
    mServiceConsumer.Ready
    mRequestId = mTWSAPI.RequestHistoricalData(mBarDataSpecifier, Me)
End If
End Sub

Private Sub mTWSAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
If Not retrying Then
    mServiceConsumer.Error StandardSPErrorCodes.HDCantConnectDataSource, _
                            Description
End If
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let barDataInputServiceConsumer(ByVal value As IBarDataInputServiceConsumer)
Set mServiceConsumer = value
End Property

Friend Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Friend Property Let commonServiceConsumer(ByVal value As ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property

Friend Property Let connectionRetryIntervalSecs(ByVal value As Long)
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

Friend Property Let port(ByVal value As Long)
mPort = value
End Property

Friend Property Let providerKey(ByVal value As String)
mProviderKey = value
End Property

Friend Property Let server(ByVal value As String)
mServer = IIf(value = "", "127.0.0.1", value)
End Property

Friend Property Let ServiceProviderName(ByVal value As String)
mServiceProviderName = value
End Property

Friend Property Let TWSLogLevel(ByVal value As TWSLogLevels)
mTWSLogLevel = value
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
mServiceConsumer.Error StandardSPErrorCodes.HDRequestFailed, reason
End Sub

Friend Sub historicalDataRequestInvalid(ByVal reason As String)
mServiceConsumer.Error StandardSPErrorCodes.HDRequestInvalid, reason
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub ProcessBar()
Dim lBar As TradeBuildSP.Bar

With lBar
    If Not mTWSAPI.ReadHistoricalBar( _
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


