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

Private mHistDataServiceProvider As HistDataServiceProvider

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

Private mBarDataSpecifier  As TradeBuildSP.BarDataSpecifier

Private mWaitingForBarData As Boolean
Private mBarRequested As Boolean

Private mLogLevel As LogLevels
Private mTWSLogLevel As TWSLogLevels

Private mSessionBuilder As SessionBuilder
Private mSession As Session

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mClientID = -1
mPort = 7496
mConnectionRetryIntervalSecs = 10
mLogLevel = LogLevelLow
mTWSLogLevel = TWSLogLevelError

mRequestId = -1
End Sub

Private Sub Class_Terminate()
finish
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
On Error GoTo err
If mRequestId <> -1 Then mTWSAPI.CancelHistoricalData mRequestId
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "HistDataReader" & "::" & "IBarDataReader_CancelFetch", _
                err.Description
End Sub

Private Sub IBarDataReader_FetchBars( _
                            barSpecifier As TradeBuildSP.BarDataSpecifier)

On Error GoTo err
If mRequestInProgress Then err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "IBTWSSP.HistDataReader::FetchBars", _
                                    "Request already in progress"

Select Case barSpecifier.BarLengthUnits
Case TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear, _
        TimePeriodVolume
Case Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.HDRequestInvalid, _
                            "Bar time unit not supported", _
                            mHistDataServiceProvider.handle
    Exit Sub
End Select

mRequestInProgress = True
mWaitingForBarData = True
mBarDataSpecifier = barSpecifier

Set mSessionBuilder = New SessionBuilder
mSessionBuilder.SessionStartTime = mBarDataSpecifier.contract.SessionStartTime
mSessionBuilder.SessionEndTime = mBarDataSpecifier.contract.SessionEndTime
Set mSession = mSessionBuilder.Session

If mTWSAPI Is Nothing Then
    If mClientID = -1 Then mClientID = Int(&H7FFFFFFF * Rnd) + 1
    Set mTWSAPI = gGetTWSAPIInstance(mServer, mPort, mClientID, mProviderKey, mConnectionRetryIntervalSecs, mkeepConnection)
    mTWSAPI.logLevel = mLogLevel
    mTWSAPI.TWSLogLevel = mTWSLogLevel
    If mTWSAPI.connectionState = ConnConnected Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.HDConnectedToDataSource, _
                            "Connected to " & mTWSAPI.connectionString, _
                            mHistDataServiceProvider.handle
    End If
End If

If mTWSAPI.connectionState = ConnConnected Then
    mServiceConsumer.Ready
    mRequestId = mTWSAPI.RequestHistoricalData(mBarDataSpecifier, Me)
    
    If mRequestId = -1 Then
        ' the request couldn't be serviced
        mServiceConsumer.BarDataComplete
        mRequestInProgress = False
    End If

End If
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "HistDataReader" & "::" & "IBarDataReader_FetchBars", _
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
handleFatalError err.number, _
                "IBTWSSP" & "." & "HistDataReader" & "::" & "IBarDataReader_FireNextBar", _
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
handleFatalError err.number, _
                "IBTWSSP" & "." & "HistDataReader" & "::" & "IBarDataReader_ReleaseDataStore", _
                err.Description
End Sub

Private Function IBarDataReader_Supports( _
                            ByVal capabilities As Long) As Boolean
IBarDataReader_Supports = gHistDataSupports(capabilities)
End Function

'================================================================================
' mTWSAPI Event Handlers
'================================================================================

Private Sub mTWSAPI_Connected()
On Error GoTo err
mServiceConsumer.NotifyEvent StandardSPEventCodes.HDConnectedToDataSource, _
                        "Connected to " & mTWSAPI.connectionString, _
                        mHistDataServiceProvider.handle
If mRequestInProgress Then
    mServiceConsumer.Ready
    mRequestId = mTWSAPI.RequestHistoricalData(mBarDataSpecifier, Me)
End If
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "HistDataReader" & "::" & "mTWSAPI_Connected", _
                err.Description
End Sub

Private Sub mTWSAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
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
handleFatalError err.number, _
                "IBTWSSP" & "." & "HistDataReader" & "::" & "mTWSAPI_ConnectFailed", _
                err.Description
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed( _
                ByVal reconnecting As Boolean)
On Error GoTo err
If reconnecting Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.HDReConnectingDataSource, _
                        "Lost connection to TWS (reconnecting)", _
                        mHistDataServiceProvider.handle
Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.HDLostConnectionToDataSource, _
                        "Lost connection to TWS", _
                        mHistDataServiceProvider.handle
End If
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "HistDataReader" & "::" & "mTWSAPI_ConnectionToTWSClosed", _
                err.Description
End Sub

Private Sub mTWSAPI_Disconnected( _
                ByVal reason As String)
On Error GoTo err
mServiceConsumer.NotifyEvent StandardSPEventCodes.HDDisconnectedFromDataSource, _
                        reason, _
                        mHistDataServiceProvider.handle
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "HistDataReader" & "::" & "mTWSAPI_Disconnected", _
                err.Description
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

Friend Property Let HistDataServiceProvider( _
                ByVal value As HistDataServiceProvider)
Set mHistDataServiceProvider = value
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

Friend Sub finish()
If Not mTWSAPI Is Nothing Then gReleaseTWSAPIInstance mTWSAPI
Set mTWSAPI = Nothing
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

Friend Sub historicalDataRequestFailed( _
                ByVal reason As String)
mServiceConsumer.NotifyEvent StandardSPEventCodes.HDRequestFailed, _
                        reason, _
                        mHistDataServiceProvider.handle
End Sub

Friend Sub historicalDataRequestInvalid( _
                ByVal reason As String)
mServiceConsumer.NotifyEvent StandardSPEventCodes.HDRequestInvalid, _
                        reason, _
                        mHistDataServiceProvider.handle
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub handleFatalError( _
                ByVal number As Long, _
                ByVal source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError number, _
                            source, _
                            Description, _
                            mHistDataServiceProvider.handle

finish
mHistDataServiceProvider.dataReaderFinished Me
End Sub

Private Sub ProcessBar()
Dim lBar As TradeBuildSP.Bar
Dim barLength As Long
Dim barUnit As TimePeriodUnits

Do While mTWSAPI.ReadHistoricalBar(mRequestId, lBar, barLength, barUnit)
    mSessionBuilder.setSessionCurrentTime lBar.timestamp
    If mBarDataSpecifier.includeBarsOutsideSession Or _
        mSession.isTimeInSession(lBar.timestamp) _
    Then
        mDataConsumer.Bar lBar, barLength, barUnit
        Exit Sub
    End If
Loop

mServiceConsumer.BarDataComplete
mRequestInProgress = False
mRequestId = -1

End Sub


