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

Private Const ConstMomentumBarsName As String = "Constant momentum bars"
Private Const ConstVolBarsName As String = "Constant volume bars"

Private Const ConstMomentumBarsParamTicksPerBar As String = "Ticks move per bar"
Private Const ConstVolBarsParamVolPerBar As String = "Volume per bar"

Private Const InputValueNamePrice As String = "Price"
Private Const InputValueNameVolume As String = "Volume"

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
Private mInstrument As instrument

Private mHistDataServiceProvider    As histDataServiceProvider

Private mCommonServiceConsumer      As TradeBuildSP.ICommonServiceConsumer
Private mServiceConsumer            As TradeBuildSP.IBarDataInputServiceConsumer
Private mDataConsumer               As TradeBuildSP.IBarDataConsumer
Private mServiceProviderName        As String

Private mRequestInProgress          As Boolean
            
Private mBarDataSpecifier           As TradeBuildSP.IBarDataSpecifier

Private mContract                   As Contract

Private WithEvents mBarReader       As BarDataReader
Attribute mBarReader.VB_VarHelpID = -1
Private WithEvents mTickReader      As TickDataReader
Attribute mTickReader.VB_VarHelpID = -1
Private mRequestTicks               As Boolean

Private mWaitingForData             As Boolean
Private mBarRequested               As Boolean

Private mSessionBuilder             As SessionBuilder
Private mSession                    As Session

Private mMode                       As AccessModes

Private mBarLengthMinutes           As Long
Private mNumberOfBars               As Long
Private mFromDate                   As Date
Private mToDate                     As Date

Private mStudyManager               As StudyManager
Private mPriceHandle                As Long
Private mVolumeHandle               As Long

Private mBarsStudy                  As Study
Private mBars                       As Bars

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

Set mInstrument = lInstrumentFactory.loadBySpecifier(mBarDataSpecifier.Contract.specifier)
Debug.Assert Not mInstrument Is Nothing

generateRequestDetails

If mRequestTicks Then
    Set mTickReader = mInstrument.CreateTickDataReader
Else
    Set mBarReader = mInstrument.CreateBarDataReader
End If

Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataReader" & "::" & "IBarDataReader_FetchBars", _
                Err.Description

End Sub

Private Sub IBarDataReader_FireNextBar()
On Error GoTo Err
If mWaitingForData Then
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
mWaitingForData = False
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

mWaitingForData = True

mServiceConsumer.NotifyEvent StandardSPEventCodes.HDConnectedToDataSource, _
                        "Connected to data source", _
                        mHistDataServiceProvider.handle

If mBarDataSpecifier.includeBarsOutsideSession Then
    mBarReader.FetchBars mBarDataSpecifier.barType, _
                        mBarLengthMinutes, _
                        mNumberOfBars, _
                        mFromDate, _
                        mToDate
Else
    mBarReader.FetchBars mBarDataSpecifier.barType, _
                        mBarLengthMinutes, _
                        MaxLong, _
                        mFromDate, _
                        mToDate
End If

mServiceConsumer.Ready
End Sub

'@===============================================================================
' mTickReader Event Handlers
'@===============================================================================

Private Sub mTickReader_ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
On Error GoTo Err
mServiceConsumer.NotifyEvent StandardSPEventCodes.HDCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc, _
                        mHistDataServiceProvider.handle

Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataReader" & "::" & "mTickReader_ConnectFailed", _
                Err.Description
End Sub

Private Sub mTickReader_DataFetched()
Dim data() As Byte
Dim tdd As TickDataDecoder
Dim startTime As Date
Dim formatId As String
Dim basePrice As Double
Dim ticksize As Double
Dim tick As GenericTick
Dim et As Single

Do While mTickReader.ReadData(startTime, formatId, basePrice, ticksize, data)
    
    Set tdd = CreateTickDecoder(startTime, ticksize, basePrice, data, formatId)
    
    Do While tdd.getNextTick(tick)
        processTick tick
    Loop
Loop

If Not mRequestInProgress Then Exit Sub
mWaitingForData = False
If mBarRequested Then
    mBarRequested = False
    ProcessBar
End If

End Sub

Private Sub mTickReader_Error( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
mCommonServiceConsumer.ServiceProviderError errorCode, _
                                    errorDesc, _
                                    mHistDataServiceProvider.handle
End Sub

Private Sub mTickReader_Ready()
mWaitingForData = True

mServiceConsumer.NotifyEvent StandardSPEventCodes.HDConnectedToDataSource, _
                        "Connected to data source", _
                        mHistDataServiceProvider.handle

mTickReader.FetchData mFromDate, _
                        mToDate

mServiceConsumer.Ready

setupBars
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

Private Sub generateRequestDetails()
Dim st As SessionTimes

With mBarDataSpecifier

    mNumberOfBars = .NumberOfBars
    
    Select Case .BarLengthUnits
    Case TimePeriodUnits.TimePeriodDay
        mBarLengthMinutes = 60
        mNumberOfBars = .NumberOfBars * 24 * .barLength
    Case TimePeriodUnits.TimePeriodHour
        mBarLengthMinutes = 60
        mNumberOfBars = .NumberOfBars * .barLength
    Case TimePeriodUnits.TimePeriodMinute
        If (.barLength Mod 60) = 0 Then
            mBarLengthMinutes = 60
            mNumberOfBars = .NumberOfBars * (.barLength / 60)
        ElseIf (.barLength Mod 15) = 0 Then
            mBarLengthMinutes = 15
            mNumberOfBars = .NumberOfBars * (.barLength / 15)
        ElseIf (.barLength Mod 5) = 0 Then
            mBarLengthMinutes = 5
            mNumberOfBars = .NumberOfBars * (.barLength / 5)
        Else
            mBarLengthMinutes = 1
            mNumberOfBars = .NumberOfBars * .barLength
        End If
    Case TimePeriodUnits.TimePeriodMonth
        mBarLengthMinutes = 60
        mNumberOfBars = .NumberOfBars * 24 * 22 * .barLength
    Case TimePeriodUnits.TimePeriodSecond
        mBarLengthMinutes = 1
        mNumberOfBars = .NumberOfBars / 60 * .barLength
    Case TimePeriodUnits.TimePeriodWeek
        mBarLengthMinutes = 60
        mNumberOfBars = .NumberOfBars * 24 * 5 * .barLength
    Case TimePeriodUnits.TimePeriodYear
        mBarLengthMinutes = 60
        mNumberOfBars = .NumberOfBars * 24 * 260 * .barLength
    Case TimePeriodUnits.TimePeriodVolume, _
            TimePeriodUnits.TimePeriodTickMovement, _
            TimePeriodUnits.TimePeriodTickVolume
        mBarLengthMinutes = 1
    End Select
    
    
    If .BarLengthUnits = TimePeriodVolume Or _
        .BarLengthUnits = TimePeriodTickMovement Or _
        .BarLengthUnits = TimePeriodTickVolume _
    Then
        mRequestTicks = True
        
        If .FromDate <> 0 And .ToDate <> 0 Then
            st = GetSessionTimes(.FromDate, _
                                        mContract.sessionStartTime, _
                                        mContract.sessionEndTime)
            mFromDate = st.startTime
            mToDate = .ToDate
            mNumberOfBars = MaxLong
        ElseIf .ToDate = 0 Then
            st = GetSessionTimes(.FromDate, _
                                        mContract.sessionStartTime, _
                                        mContract.sessionEndTime)
            mFromDate = st.startTime
            st = GetOffsetSessionTimes(mFromDate, _
                                        .NumberOfBars, _
                                        mContract.sessionStartTime, _
                                        mContract.sessionEndTime)
            mToDate = st.startTime
            mNumberOfBars = MaxLong
        ElseIf .FromDate = 0 Then
            mToDate = .ToDate
            st = GetOffsetSessionTimes(mToDate, _
                                        -.NumberOfBars, _
                                        mContract.sessionStartTime, _
                                        mContract.sessionEndTime)
            mFromDate = st.startTime
            mNumberOfBars = MaxLong
        Else
            mToDate = Now
            st = GetOffsetSessionTimes(mToDate, _
                                        -.NumberOfBars, _
                                        mContract.sessionStartTime, _
                                        mContract.sessionEndTime)
            mFromDate = st.startTime
            mNumberOfBars = MaxLong
        End If
    Else
        If .ToDate <> 0 Then
            mToDate = .ToDate
            
            If .FromDate = 0 Then
                ' calculate the earliest possible date for the supplied number of bars (assuming
                ' there are no gaps in the data)
                mFromDate = OffsetBarStartTime(.ToDate, _
                                            .barLength, _
                                            .BarLengthUnits, _
                                            -1 * .NumberOfBars, _
                                            mContract.sessionStartTime, _
                                            mContract.sessionEndTime)
            Else
                mFromDate = .FromDate
            End If
        Else
            mFromDate = .FromDate
            mToDate = OffsetBarStartTime(.FromDate, _
                                        .barLength, _
                                        .BarLengthUnits, _
                                        .NumberOfBars, _
                                        mContract.sessionStartTime, _
                                        mContract.sessionEndTime)
        End If
    End If
End With

End Sub

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

Private Sub processTick( _
                ByRef tick As GenericTick)
Select Case tick.tickType
Case TickTypeBid
    If mBarDataSpecifier.barType = BarTypeBid Then mStudyManager.notifyInput mPriceHandle, tick.price, tick.timestamp
Case TickTypeAsk
    If mBarDataSpecifier.barType = BarTypeAsk Then mStudyManager.notifyInput mPriceHandle, tick.price, tick.timestamp
Case TickTypeTrade
    If mBarDataSpecifier.barType = BarTypeTrade Then mStudyManager.notifyInput mPriceHandle, tick.price, tick.timestamp
Case TickTypeVolume
    If mBarDataSpecifier.barType = BarTypeTrade Then mStudyManager.notifyInput mVolumeHandle, tick.size, tick.timestamp
End Select
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
Static i As Long
Dim lBar As StudyUtils26.Bar

If mRequestTicks Then
    i = i + 1
    If i > mBars.Count Then
        mServiceConsumer.BarDataComplete
        mRequestInProgress = False
        ReadBar = False
    Else
        Set lBar = mBars.Bar(i)
        timestamp = lBar.DateTime
        barType = mBarDataSpecifier.barType
        barLength = mBarDataSpecifier.barLength
        openPrice = lBar.openValue
        highPrice = lBar.highValue
        lowPrice = lBar.lowValue
        closePrice = lBar.closeValue
        volume = mBarsStudy.getStudyValue("volume", i).value
        ReadBar = True
    End If
Else
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
End If
End Function

Private Sub setupBars()
Dim lInputStudy As InputStudy
Dim lBarStudy As BarStudy
Dim inputValueNames(1) As String

mWaitingForData = True

mServiceConsumer.NotifyEvent StandardSPEventCodes.HDConnectedToDataSource, _
                        "Connected to data source", _
                        mHistDataServiceProvider.handle

mTickReader.FetchData mFromDate, _
                        mToDate

mServiceConsumer.Ready

Set mStudyManager = New StudyManager
mStudyManager.setSessionTimes mContract.sessionStartTime, mContract.sessionEndTime, mContract.TimeZone, False
Set lInputStudy = mStudyManager.addSource("Tick data")
mPriceHandle = mStudyManager.addInput(lInputStudy, _
                                    InputValueNamePrice, _
                                    "", _
                                    InputTypeReal, _
                                    True, _
                                    mContract.ticksize)
mVolumeHandle = mStudyManager.addInput(lInputStudy, _
                                    InputValueNameVolume, _
                                    "", _
                                    InputTypeInteger, _
                                    False, _
                                    1)
                                    
inputValueNames(0) = InputValueNamePrice
inputValueNames(1) = InputValueNameVolume
Select Case mBarDataSpecifier.BarLengthUnits
Case TimePeriodUnits.TimePeriodTickMovement
    Set lBarStudy = mStudyManager.addStudy(ConstMomentumBarsName, _
                                            lInputStudy, _
                                            inputValueNames, _
                                            CreateParametersFromString(ConstMomentumBarsParamTicksPerBar & "=" & mBarDataSpecifier.barLength))
Case TimePeriodUnits.TimePeriodTickVolume
Case TimePeriodUnits.TimePeriodVolume
    Set lBarStudy = mStudyManager.addStudy(ConstVolBarsName, _
                                            lInputStudy, _
                                            inputValueNames, _
                                            CreateParametersFromString(ConstVolBarsParamVolPerBar & "=" & mBarDataSpecifier.barLength))
End Select

mStudyManager.startStudy lBarStudy, 0
Set mBarsStudy = lBarStudy
Set mBars = lBarStudy.Bars
End Sub

