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

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mHistDataServiceProvider As QTHistDataServiceProvider

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

Private mBarDataSpecifier  As TradeBuildSP.IBarDataSpecifier

Private mWaitingForBarData As Boolean
Private mBarRequested As Boolean

Private mLogLevel As LogLevels

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

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
                            ) As TradeBuildSP.IBarDataSpecifier
Set IBarDataReader_BarDataSpecifier = mBarDataSpecifier
End Property

Private Sub IBarDataReader_CancelFetch()
On Error GoTo err
mQTAPI.cancelHistoricalBarData mRequestId, mLockId
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTHistDataReader" & "::" & "IBarDataReader_CancelFetch", _
                err.Description
End Sub

Private Sub IBarDataReader_FetchBars( _
                ByVal barSpecifier As TradeBuildSP.IBarDataSpecifier)

On Error GoTo err
If mRequestInProgress Then err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "IBTWSSP.HistDataReader::FetchBars", _
                                    "Request already in progress"

Select Case barSpecifier.BarLengthUnits
Case TimePeriodMinute, _
    TimePeriodHour, _
    TimePeriodDay, _
    TimePeriodWeek, _
    TimePeriodMonth, _
    TimePeriodYear
Case Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.HDRequestInvalid, _
                            "Bar time unit not supported", _
                            mHistDataServiceProvider.handle
    Exit Sub
End Select

mRequestInProgress = True
mWaitingForBarData = True
Set mBarDataSpecifier = barSpecifier

If mQTAPI Is Nothing Then
    Set mQTAPI = gGetQTAPIInstance(mServer, mPort, mPassword, mProviderKey, mConnectionRetryIntervalSecs, mkeepConnection)
    mQTAPI.CommonServiceConsumer = mCommonServiceConsumer
    mQTAPI.logLevel = mLogLevel
End If

If mQTAPI.connectionState = ConnConnected Then
    mServiceConsumer.Ready
    mRequestId = mQTAPI.requestHistoricalBarData( _
                                mBarDataSpecifier, _
                                Me, _
                                mLockId)
End If
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTHistDataReader" & "::" & "IBarDataReader_FetchBars", _
                err.Description

End Sub

Private Sub IBarDataReader_FireNextBar()
On Error GoTo err
If mWaitingForBarData Then
    mBarRequested = True
Else
    ProcessBar
End If
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTHistDataReader" & "::" & "IBarDataReader_FireNextBar", _
                err.Description
End Sub

Private Property Get IBarDataReader_NumberOfBars() As Long

End Property

Private Sub IBarDataReader_ReleaseDataStore()
On Error GoTo err
finish
mHistDataServiceProvider.dataReaderFinished Me
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTHistDataReader" & "::" & "IBarDataReader_ReleaseDataStore", _
                err.Description
End Sub

Private Function IBarDataReader_Supports( _
                            ByVal capabilities As Long) As Boolean
IBarDataReader_Supports = False
End Function

'================================================================================
' LockConsumer Interface Members
'================================================================================

Private Sub LockConsumer_lockAcquired(ByVal lockId As Long)
On Error GoTo err
If mLockId <> lockId Then err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "QTSP.QTHistDataReader::lockAcquired", _
                                    "Lock id acquired is " & lockId & ": expected " & mLockId
    
mRequestId = mQTAPI.requestHistoricalBarData(mBarDataSpecifier, Me, mLockId)
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTHistDataReader" & "::" & "LockConsumer_lockAcquired", _
                err.Description
End Sub

Private Sub LockConsumer_waitingForLock(ByVal lockId As Long)
mLockId = lockId
End Sub

'================================================================================
' mqtapi Event Handlers
'================================================================================

Private Sub mQTAPI_Connected()
On Error GoTo err

mServiceConsumer.NotifyEvent StandardSPEventCodes.HDConnectedToDataSource, _
                        "Connected to data source", _
                        mHistDataServiceProvider.handle

If mRequestInProgress Then
    mServiceConsumer.Ready
    mRequestId = mQTAPI.requestHistoricalBarData(mBarDataSpecifier, Me, mLockId)
End If
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTHistDataReader" & "::" & "mQTAPI_Connected", _
                err.Description
End Sub

Private Sub mQTAPI_ConnectFailed( _
                ByVal Description As String, _
                ByVal retrying As Boolean)
On Error GoTo err
If retrying Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.HDRetryConnectDataSource, _
                            Description, _
                            mHistDataServiceProvider.handle
Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.HDCantConnectDataSource, _
                            Description, _
                            mHistDataServiceProvider.handle
End If
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTHistDataReader" & "::" & "mQTAPI_ConnectFailed", _
                err.Description
End Sub

Private Sub mQTAPI_ConnectionClosed( _
                ByVal reconnecting As Boolean)
On Error GoTo err
If reconnecting Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.HDReConnectingDataSource, _
                        "Lost connection to QT (reconnecting)", _
                        mHistDataServiceProvider.handle
Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.HDLostConnectionToDataSource, _
                        "Lost connection to QT", _
                        mHistDataServiceProvider.handle
End If
Exit Sub
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTHistDataReader" & "::" & "mQTAPI_ConnectionClosed", _
                err.Description
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

Friend Property Let histDataServiceProvider( _
                ByVal value As QTHistDataServiceProvider)
Set mHistDataServiceProvider = value
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

Friend Sub finish()
If Not mQTAPI Is Nothing Then gReleaseQTAPIInstance mQTAPI
Set mQTAPI = Nothing
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing
End Sub

Friend Sub historicalDataAvailable()
mWaitingForBarData = False
If mBarRequested Then
    mBarRequested = False
    ProcessBar
End If
End Sub

Friend Sub historicalDataRequestFailed(ByVal reason As String)
mServiceConsumer.NotifyEvent StandardSPEventCodes.HDRequestFailed, _
                        reason, _
                        mHistDataServiceProvider.handle
End Sub

Friend Sub historicalDataRequestInvalid(ByVal reason As String)
mServiceConsumer.NotifyEvent StandardSPEventCodes.HDRequestInvalid, _
                        reason, _
                        mHistDataServiceProvider.handle
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError Number, _
                                Source, _
                                Description, _
                                mHistDataServiceProvider.handle

finish
mHistDataServiceProvider.dataReaderFinished Me
End Sub

Private Sub ProcessBar()
Dim lBar As TradeBuildSP.Bar
Dim barLength As Long
Dim barUnits As TimePeriodUnits

With lBar
    If Not mQTAPI.readHistoricalBar( _
                            mRequestId, _
                            .timestamp, _
                            barLength, _
                            barUnits, _
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
mDataConsumer.Bar lBar, barLength, barUnits
End Sub




