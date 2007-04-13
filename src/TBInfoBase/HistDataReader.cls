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

Private mHistDataServiceProvider As histDataServiceProvider

Private mCommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer
Private mServiceConsumer As TradeBuildSP.IBarDataInputServiceConsumer
Private mDataConsumer As TradeBuildSP.IBarDataConsumer
Private mServiceProviderName As String

Private mRequestInProgress As Boolean

Private mBarDataSpecifier  As TradeBuildSP.BarDataSpecifier

Private mContract As IContract

Private WithEvents mBarReader As TradingDO2.BarDataAccessor
Attribute mBarReader.VB_VarHelpID = -1

Private mWaitingForBarData As Boolean
Private mBarRequested As Boolean

Private mSessionBuilder As SessionBuilder
Private mSession As Session

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IBarDataReader Interface Members
'================================================================================

Private Property Get IBarDataReader_BarDataSpecifier( _
                            ) As TradeBuildSP.BarDataSpecifier
IBarDataReader_BarDataSpecifier = mBarDataSpecifier
End Property

Private Sub IBarDataReader_CancelFetch()
On Error GoTo Err
mRequestInProgress = False
finish
Exit Sub
Err:
handleFatalError Err.Number, _
                "TBInfoBase" & "." & "HistDataReader" & "::" & "IBarDataReader_CancelFetch", _
                Err.Description
End Sub

Private Sub IBarDataReader_FetchBars( _
                            ByRef barSpecifier As TradeBuildSP.BarDataSpecifier)
Dim lInstrumentFactory As TradingDO2.cInstrumentFactory
Dim lInstrument As TradingDO2.cInstrument

On Error GoTo Err
If mRequestInProgress Then Err.Raise ErrorCodes.IllegalStateException, _
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
mBarDataSpecifier = barSpecifier
Set mContract = mBarDataSpecifier.Contract

Set mSessionBuilder = New SessionBuilder
mSessionBuilder.sessionStartTime = mContract.sessionStartTime
mSessionBuilder.sessionEndTime = mContract.sessionEndTime
Set mSession = mSessionBuilder.Session

Set lInstrumentFactory = New TradingDO2.cInstrumentFactory

With mBarDataSpecifier.Contract.Specifier
    Set lInstrument = lInstrumentFactory.loadBySpec(.symbol, _
                                secTypeToString(.SecType), _
                                Left$(.Expiry, 6), _
                                .exchange, _
                                .currencyCode, _
                                .Strike, _
                                optRightToString(.Right), _
                                .localSymbol)
End With
Debug.Assert Not lInstrument Is Nothing

Set mBarReader = lInstrument.CreateBarDataAccessor
Exit Sub
Err:
handleFatalError Err.Number, _
                "TBInfoBase" & "." & "HistDataReader" & "::" & "IBarDataReader_FetchBars", _
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
                "TBInfoBase" & "." & "HistDataReader" & "::" & "IBarDataReader_FireNextBar", _
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
                "TBInfoBase" & "." & "HistDataReader" & "::" & "IBarDataReader_ReleaseDataStore", _
                Err.Description
End Sub

Private Function IBarDataReader_Supports( _
                            ByVal capabilities As Long) As Boolean
IBarDataReader_Supports = gHistDataSupports(capabilities)
End Function

'================================================================================
' mBarReader Event Handlers
'================================================================================

Private Sub mBarReader_BarsFetched()
Dim lBar As TradeBuildSP.Bar

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
mServiceConsumer.NotifyEvent StandardSPEventCodes.HDCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc, _
                        mHistDataServiceProvider.handle

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

mWaitingForBarData = True

mServiceConsumer.NotifyEvent StandardSPEventCodes.HDConnectedToDataSource, _
                        "Connected to data source", _
                        mHistDataServiceProvider.handle

With mBarDataSpecifier

    If .From < #1/1/1900# Then .From = #1/1/1900# ' don't exceed range of SmallDateTime
    
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
        numberOfBars = 2880
    End Select
    
    
    setBarSpecifierTimes mBarDataSpecifier, _
                        mContract
    
    mBarReader.FetchBars .barType, BarLengthMinutes, numberOfBars, .From, .To
End With

mServiceConsumer.Ready
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let BarDataInputServiceConsumer(ByVal value As IBarDataInputServiceConsumer)
Set mServiceConsumer = value
End Property

Friend Property Let DataConsumer(ByVal value As IBarDataConsumer)
Set mDataConsumer = value
End Property

Friend Property Let CommonServiceConsumer(ByVal value As ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property

Friend Property Let histDataServiceProvider( _
                ByVal value As histDataServiceProvider)
Set mHistDataServiceProvider = value
End Property

Friend Property Let ServiceProviderName(ByVal value As String)
mServiceProviderName = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub finish()
Set mBarReader = Nothing
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing
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

Private Sub setBarSpecifierTimes( _
                ByRef barSpecifier As BarDataSpecifier, _
                ByVal pContract As IContract)

With barSpecifier
    If .BarLengthUnits = TimePeriodVolume Or .BarLengthUnits = TimePeriodTickMovement Or .BarLengthUnits = TimePeriodTickVolume Then Exit Sub
    If .To <> 0 Then
        .From = OffsetBarStartTime(.To, _
                                    .barLength, _
                                    .BarLengthUnits, _
                                    -.numberOfBars, _
                                    pContract.sessionStartTime, _
                                    pContract.sessionEndTime)
    Else
        .To = OffsetBarStartTime(.From, _
                                    .barLength, _
                                    .BarLengthUnits, _
                                    .numberOfBars, _
                                    pContract.sessionStartTime, _
                                    pContract.sessionEndTime)
    End If
End With
End Sub


