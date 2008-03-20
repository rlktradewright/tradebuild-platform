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

Implements TradeBuildSP.IBarDataReader

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TBInfoBase26"
Private Const ModuleName As String = "HistDataReader"

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

Private mHistDataServiceProvider As histDataServiceProvider

Private mCommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer
Private mServiceConsumer As TradeBuildSP.IBarDataInputServiceConsumer
Private mDataConsumer As TradeBuildSP.IBarDataConsumer
Private mServiceProviderName As String

Private mRequestInProgress As Boolean

Private mBarDataSpecifier  As TradeBuildSP.IBarDataSpecifier

Private mContract As Contract

Private WithEvents mBarReader As BarDataReader
Attribute mBarReader.VB_VarHelpID = -1

Private mWaitingForBarData As Boolean
Private mBarRequested As Boolean

Private mSessionBuilder As SessionBuilder
Private mSession As Session

Private mMode As AccessModes

'@===============================================================================
' Class Event Handlers
'@===============================================================================

'@===============================================================================
' IBarDataReader Interface Members
'@===============================================================================

Private Property Get IBarDataReader_BarDataSpecifier( _
                            ) As TradeBuildSP.IBarDataSpecifier
Set IBarDataReader_BarDataSpecifier = mBarDataSpecifier
End Property

Private Sub IBarDataReader_CancelFetch()
On Error GoTo Err
mRequestInProgress = False
finish
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataReader" & "::" & "IBarDataReader_CancelFetch", _
                Err.Description
End Sub

Private Sub IBarDataReader_FetchBars( _
                            ByVal barSpecifier As TradeBuildSP.IBarDataSpecifier)
Dim lInstrumentFactory As InstrumentFactory
Dim lInstrument As instrument

On Error GoTo Err
If mRequestInProgress Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "TBInfoBase.HistDataReader::FetchBars", _
                                    "Request already in progress"

Select Case barSpecifier.BarLengthUnits
Case TimePeriodMinute, _
    TimePeriodHour, _
    TimePeriodDay, _
    TimePeriodWeek, _
    TimePeriodMonth, _
    TimePeriodYear, _
    TimePeriodVolume, _
    TimePeriodTickMovement
Case Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.HDRequestInvalid, _
                            "Bar time unit not supported", _
                            mHistDataServiceProvider.handle
    Exit Sub
End Select

mRequestInProgress = True
Set mBarDataSpecifier = barSpecifier
Set mContract = mBarDataSpecifier.Contract

Set mSessionBuilder = New SessionBuilder
mSessionBuilder.sessionStartTime = mContract.sessionStartTime
mSessionBuilder.sessionEndTime = mContract.sessionEndTime
mSessionBuilder.TimeZone = mContract.TimeZone
Set mSession = mSessionBuilder.Session

Set lInstrumentFactory = mTradingDB.InstrumentFactory

Set lInstrument = lInstrumentFactory.loadBySpecifier(mBarDataSpecifier.Contract.specifier)
Debug.Assert Not lInstrument Is Nothing

Set mBarReader = lInstrument.CreateBarDataReader
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataReader" & "::" & "IBarDataReader_FetchBars", _
                Err.Description

End Sub

Private Sub IBarDataReader_FireNextBar()
On Error GoTo Err
If mWaitingForBarData Then
    mBarRequested = True
Else
    ProcessBar
End If
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataReader" & "::" & "IBarDataReader_FireNextBar", _
                Err.Description
End Sub

Private Property Get IBarDataReader_NumberOfBars() As Long

End Property

Private Sub IBarDataReader_ReleaseDataStore()
On Error GoTo Err
finish
mHistDataServiceProvider.dataReaderFinished Me
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataReader" & "::" & "IBarDataReader_ReleaseDataStore", _
                Err.Description
End Sub

Private Function IBarDataReader_Supports( _
                            ByVal capabilities As Long) As Boolean
IBarDataReader_Supports = gHistDataSupports(capabilities, mMode)
End Function

'@===============================================================================
' mBarReader Event Handlers
'@===============================================================================

Private Sub mBarReader_BarsFetched()

If Not mRequestInProgress Then Exit Sub
mWaitingForBarData = False
If mBarRequested Then
    mBarRequested = False
    ProcessBar
End If

End Sub

Private Sub mBarReader_ConnectFailed( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
On Error GoTo Err
mServiceConsumer.NotifyEvent StandardSPEventCodes.HDCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc, _
                        mHistDataServiceProvider.handle

Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataReader" & "::" & "mBarReader_ConnectFailed", _
                Err.Description
End Sub

Private Sub mBarReader_Error( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
mCommonServiceConsumer.ServiceProviderError errorCode, _
                                    errorDesc, _
                                    mHistDataServiceProvider.handle
End Sub

Private Sub mBarReader_Ready()
Dim BarLengthMinutes As Long
Dim numberOfBars As Long
Dim fromDate As Date
Dim toDate As Date
Dim st As SessionTimes

mWaitingForBarData = True

mServiceConsumer.NotifyEvent StandardSPEventCodes.HDConnectedToDataSource, _
                        "Connected to data source", _
                        mHistDataServiceProvider.handle

With mBarDataSpecifier

    numberOfBars = .numberOfBars
    
    Select Case .BarLengthUnits
    Case TimePeriodUnits.TimePeriodDay
        BarLengthMinutes = 60
        numberOfBars = .numberOfBars * 24 * .barLength
    Case TimePeriodUnits.TimePeriodHour
        BarLengthMinutes = 60
        numberOfBars = .numberOfBars * .barLength
    Case TimePeriodUnits.TimePeriodMinute
        If (.barLength Mod 60) = 0 Then
            BarLengthMinutes = 60
            numberOfBars = .numberOfBars * (.barLength / 60)
        ElseIf (.barLength Mod 15) = 0 Then
            BarLengthMinutes = 15
            numberOfBars = .numberOfBars * (.barLength / 15)
        ElseIf (.barLength Mod 5) = 0 Then
            BarLengthMinutes = 5
            numberOfBars = .numberOfBars * (.barLength / 5)
        Else
            BarLengthMinutes = 1
            numberOfBars = .numberOfBars * .barLength
        End If
    Case TimePeriodUnits.TimePeriodMonth
        BarLengthMinutes = 60
        numberOfBars = .numberOfBars * 24 * 22 * .barLength
    Case TimePeriodUnits.TimePeriodSecond
        BarLengthMinutes = 1
        numberOfBars = .numberOfBars / 60 * .barLength
    Case TimePeriodUnits.TimePeriodWeek
        BarLengthMinutes = 60
        numberOfBars = .numberOfBars * 24 * 5 * .barLength
    Case TimePeriodUnits.TimePeriodYear
        BarLengthMinutes = 60
        numberOfBars = .numberOfBars * 24 * 260 * .barLength
    Case TimePeriodUnits.TimePeriodVolume, _
            TimePeriodUnits.TimePeriodTickMovement, _
            TimePeriodUnits.TimePeriodTickVolume
        BarLengthMinutes = 1
    End Select
    
    
    If .BarLengthUnits = TimePeriodVolume Or _
        .BarLengthUnits = TimePeriodTickMovement Or _
        .BarLengthUnits = TimePeriodTickVolume _
    Then
        If .fromDate <> 0 And .toDate <> 0 Then
            st = GetSessionTimes(.fromDate, _
                                        mContract.sessionStartTime, _
                                        mContract.sessionEndTime)
            fromDate = st.startTime
            toDate = .toDate
            numberOfBars = MaxLong
        ElseIf .toDate = 0 Then
            st = GetSessionTimes(.fromDate, _
                                        mContract.sessionStartTime, _
                                        mContract.sessionEndTime)
            fromDate = st.startTime
            st = GetOffsetSessionTimes(fromDate, _
                                        .numberOfBars, _
                                        mContract.sessionStartTime, _
                                        mContract.sessionEndTime)
            toDate = st.startTime
            numberOfBars = MaxLong
        ElseIf .fromDate = 0 Then
            toDate = .toDate
            st = GetOffsetSessionTimes(toDate, _
                                        -.numberOfBars, _
                                        mContract.sessionStartTime, _
                                        mContract.sessionEndTime)
            fromDate = st.startTime
            numberOfBars = MaxLong
        Else
            toDate = Now
            st = GetOffsetSessionTimes(toDate, _
                                        -.numberOfBars, _
                                        mContract.sessionStartTime, _
                                        mContract.sessionEndTime)
            fromDate = st.startTime
            numberOfBars = MaxLong
        End If
    Else
        If .toDate <> 0 Then
            toDate = .toDate
            
            If .fromDate = 0 Then
                ' calculate the earliest possible date for the supplied number of bars (assuming
                ' there are no gaps in the data)
                fromDate = OffsetBarStartTime(.toDate, _
                                            .barLength, _
                                            .BarLengthUnits, _
                                            -1 * .numberOfBars, _
                                            mContract.sessionStartTime, _
                                            mContract.sessionEndTime)
            Else
                fromDate = .fromDate
            End If
        Else
            fromDate = .fromDate
            toDate = OffsetBarStartTime(.fromDate, _
                                        .barLength, _
                                        .BarLengthUnits, _
                                        .numberOfBars, _
                                        mContract.sessionStartTime, _
                                        mContract.sessionEndTime)
        End If
    End If
    
    If .includeBarsOutsideSession Then
        mBarReader.FetchBars .barType, BarLengthMinutes, numberOfBars, fromDate, toDate
    Else
        mBarReader.FetchBars .barType, BarLengthMinutes, MaxLong, fromDate, toDate
    End If
End With

mServiceConsumer.Ready
End Sub

'@===============================================================================
' mTradingDB Event Handlers
'@===============================================================================

Private Sub mTradingDB_ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
On Error GoTo Err
mServiceConsumer.NotifyEvent StandardSPEventCodes.HDCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc, _
                        mHistDataServiceProvider.handle

Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataReader" & "::" & "mTradingDB_ConnectFailed", _
                Err.Description
End Sub

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub finish()
Set mBarReader = Nothing
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing
End Sub

Friend Sub initialise( _
                ByVal Name As String, _
                ByVal connectionString As String, _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal barDataInputServiceConsumer As TradeBuildSP.IBarDataInputServiceConsumer, _
                ByVal histDataServiceProvider As TradeBuildSP.IBarDataServiceProvider, _
                ByVal DataConsumer As TradeBuildSP.IBarDataConsumer, _
                ByVal mode As AccessModes)
                
mServiceProviderName = Name
Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = barDataInputServiceConsumer
Set mHistDataServiceProvider = histDataServiceProvider
Set mDataConsumer = DataConsumer
mMode = mode
                
mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString)

End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

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
Dim barType As BarTypes
Dim barLength As Long

With lBar
    Do While ReadBar(.timestamp, _
                    barType, _
                    barLength, _
                    .openPrice, _
                    .highPrice, _
                    .lowPrice, _
                    .closePrice, _
                    .volume, _
                    .tickVolume, _
                    .openInterest)
        mSessionBuilder.setSessionCurrentTime .timestamp
        If mBarDataSpecifier.includeBarsOutsideSession Or _
            mSession.isTimeInSession(.timestamp) _
        Then
            mDataConsumer.Bar lBar, barLength, TimePeriodMinute
            Exit Do
        End If
    Loop
End With

End Sub

Private Function ReadBar( _
                ByRef timestamp As Date, _
                ByRef barType As Long, _
                ByRef barLength As Long, _
                ByRef openPrice As Double, _
                ByRef highPrice As Double, _
                ByRef lowPrice As Double, _
                ByRef closePrice As Double, _
                ByRef volume As Long, _
                ByRef tickVolume As Long, _
                ByRef openInterest As Long) As Boolean
If Not mBarReader.ReadBar(timestamp, _
                        barType, _
                        barLength, _
                        openPrice, _
                        highPrice, _
                        lowPrice, _
                        closePrice, _
                        volume, _
                        tickVolume, _
                        openInterest) _
Then
    mServiceConsumer.BarDataComplete
    mRequestInProgress = False
    ReadBar = False
Else
    ReadBar = True
End If
End Function



