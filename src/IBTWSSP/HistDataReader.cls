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

Implements DeferredAction
Implements IBarDataReader

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "HistDataReader"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mHistDataServiceProvider As histDataServiceProvider
Private mCommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer
Private mServiceConsumer As TradeBuildSP.IBarDataInputServiceConsumer
Private mDataConsumer As TradeBuildSP.IBarDataConsumer

Private WithEvents mTWSAPI As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1

Private mRequestInProgress As Boolean

Private mRequestId As Long

Private mBarDataSpecifier  As BarDataSpecifier

Private mWaitingForBarData As Boolean
Private mBarRequested As Boolean

Private mSessionBuilder As SessionBuilder
Private mSession As Session

Private mHandle As Long

Private mNumBarsFetched As Long
Private mNumBarsProcessed As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mRequestId = -1
End Sub

Private Sub Class_Terminate()
Debug.Print "IBTWSSP.HistDataReader terminated"
End Sub

'================================================================================
' DeferredAction Interface Members
'================================================================================

Private Sub DeferredAction_run(ByVal data As Variant)
mServiceConsumer.Ready
If mRequestId = -1 Then
    ' the request couldn't be serviced
    mServiceConsumer.BarDataComplete
    mRequestInProgress = False
End If
End Sub

'================================================================================
' IBarDataReader Interface Members
'================================================================================

Private Property Get IBarDataReader_BarDataSpecifier( _
                            ) As BarDataSpecifier
Set IBarDataReader_BarDataSpecifier = mBarDataSpecifier
End Property

Private Sub IBarDataReader_CancelFetch()
On Error GoTo err
If mRequestId <> -1 And Not mTWSAPI Is Nothing Then mTWSAPI.CancelHistoricalData mRequestId
Exit Sub
err:
logMessage "IBTWSSP:HistDataReader:IBarDataReader_CancelFetch (" & err.number & "): " & err.Description, LogLevelSevere
handleFatalError err.number, _
                "IBTWSSP" & "." & "HistDataReader" & "::" & "IBarDataReader_CancelFetch", _
                err.Description
End Sub

Private Sub IBarDataReader_FetchBars( _
                            ByVal barSpecifier As BarDataSpecifier)

On Error GoTo err
If mRequestInProgress Then err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "IBTWSSP.HistDataReader::FetchBars", _
                                    "Request already in progress"

Select Case barSpecifier.barTimePeriod.Units
Case TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear, _
        TimePeriodVolume, _
        TimePeriodTickMovement

Case Else
    mServiceConsumer.notifyEvent StandardSPEventCodes.HDRequestInvalid, _
                            "Bar time unit not supported", _
                            mHandle
    Exit Sub
End Select

If barSpecifier.Contract.specifier.sectype = SecTypeCombo Then
    mServiceConsumer.notifyEvent StandardSPEventCodes.HDRequestInvalid, _
                            "Combo contracts not supported", _
                            mHandle
    Exit Sub
End If

mRequestInProgress = True
mWaitingForBarData = True
Set mBarDataSpecifier = barSpecifier

Set mSessionBuilder = New SessionBuilder
mSessionBuilder.SessionStartTime = mBarDataSpecifier.Contract.SessionStartTime
mSessionBuilder.SessionEndTime = mBarDataSpecifier.Contract.SessionEndTime
Set mSession = mSessionBuilder.Session

If mTWSAPI.connectionState = ConnConnected Then
    mRequestId = mTWSAPI.RequestHistoricalData(mBarDataSpecifier, Me)
    DeferAction Me
End If
Exit Sub
err:
logMessage "IBTWSSP:HistDataReader:IBarDataReader_FetchBars (" & err.number & "): " & err.Description, LogLevelSevere
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
logMessage "IBTWSSP:HistDataReader:IBarDataReader_FireNextBar (" & err.number & "): " & err.Description, LogLevelSevere
handleFatalError err.number, _
                "IBTWSSP" & "." & "HistDataReader" & "::" & "IBarDataReader_FireNextBar", _
                err.Description
End Sub

Private Property Get IBarDataReader_NumberOfBars() As Long

End Property

Private Sub IBarDataReader_ReleaseDataStore()
On Error GoTo err

mHistDataServiceProvider.dataReaderFinished Me
finish

Exit Sub
err:
logMessage "IBTWSSP:HistDataReader:IBarDataReader_ReleaseDataStore (" & err.number & "): " & err.Description, LogLevelSevere
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
If mRequestInProgress Then
    mServiceConsumer.Ready
    mRequestId = mTWSAPI.RequestHistoricalData(mBarDataSpecifier, Me)
End If
Exit Sub
err:
handleFatalError err.number, _
                "mTWSAPI_Connected", _
                err.Description
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub finish()
Set mHistDataServiceProvider = Nothing
Set mCommonServiceConsumer = Nothing
Set mServiceConsumer = Nothing
Set mTWSAPI = Nothing
End Sub

Friend Sub historicalDataAvailable( _
                ByVal numberOfBars As Long)
mWaitingForBarData = False
mNumBarsFetched = numberOfBars
mNumBarsProcessed = 0
If mBarRequested Then
    mBarRequested = False
     mServiceConsumer.Progress 0, 0
    ProcessBar
End If
End Sub

Friend Sub Initialise( _
                ByVal histDataServiceProvider As IBarDataServiceProvider, _
                ByVal commonServiceConsumer As ICommonServiceConsumer, _
                ByVal histDataInputServiceConsumer As IBarDataInputServiceConsumer, _
                ByVal dataConsumer As IBarDataConsumer, _
                ByVal pTWSAPI As TWSAPI, _
                ByVal handle As Long)
Set mHistDataServiceProvider = histDataServiceProvider
Set mCommonServiceConsumer = commonServiceConsumer
Set mServiceConsumer = histDataInputServiceConsumer
Set mDataConsumer = dataConsumer
Set mTWSAPI = pTWSAPI
mHandle = handle
End Sub

Friend Sub historicalDataRequestFailed( _
                ByVal reason As String)
mServiceConsumer.notifyEvent StandardSPEventCodes.HDRequestFailed, _
                        reason, _
                        mHandle
End Sub

Friend Sub historicalDataRequestInvalid( _
                ByVal reason As String)
mServiceConsumer.notifyEvent StandardSPEventCodes.HDRequestInvalid, _
                        reason, _
                        mHandle
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub handleFatalError( _
                ByVal number As Long, _
                ByVal source As String, _
                ByVal Description As String)

logMessage "Error " & err.number & ": " & _
            ProjectName & "." & ModuleName & ":" & source & vbCrLf & _
            err.Description, _
            LogLevelSevere

gReleaseAllTWSAPIInstances

mCommonServiceConsumer.NotifyFatalError number, _
                            source, _
                            Description, _
                            mHandle

End Sub

Private Sub logMessage( _
                ByVal message As String, _
                Optional ByVal logLevel As LogLevels = LogLevelNormal)
If Not gLogger.isLoggable(logLevel) Then Exit Sub
gLogger.Log logLevel, _
            ProjectName & "." & ModuleName & ": " & message
End Sub

Private Sub ProcessBar()
Dim lBar As Bar
Dim barTimePeriod As TimePeriod

If mTWSAPI.ReadHistoricalBar(mRequestId, lBar, barTimePeriod) Then
        mDataConsumer.Bar lBar, barTimePeriod
        mNumBarsProcessed = mNumBarsProcessed + 1
        If mNumBarsProcessed Mod 10 = 0 Then mServiceConsumer.Progress mNumBarsProcessed, 100 * mNumBarsProcessed / mNumBarsFetched
        Exit Sub
End If

mServiceConsumer.BarDataComplete
mRequestInProgress = False
mRequestId = -1

End Sub


