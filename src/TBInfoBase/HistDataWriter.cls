VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataWriter"
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

Implements IBarDataWriter

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TBInfoBase26"
Private Const ModuleName As String = "HistDataWriter"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

Private Type TimeframeEntry
    timeframe               As TimePeriod
    saveIntervalSeconds     As Long
    prevTimestamp           As Date
    barsWritten             As Long
    writes                  As Long
End Type

'@===============================================================================
' Member variables
'@===============================================================================

Private WithEvents mTradingDB       As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1
Private mConnectionString           As String

Private mHistDataServiceProvider    As HistDataServiceProvider

Private mCommonServiceConsumer      As ICommonServiceConsumer
Private mServiceConsumer            As IBarDataOutputServiceConsumer
Private mServiceProviderName        As String

Private mContractSpec               As ContractSpecifier
Private WithEvents mBarWriter       As BarDataWriter
Attribute mBarWriter.VB_VarHelpID = -1

Private mPath                       As String

Private mInstrument                 As instrument

Private mTimeframeTable(3)          As TimeframeEntry

Private mUseSynchronousWrites As Boolean

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
mTimeframeTable(0).timeframe.length = 1
mTimeframeTable(0).timeframe.units = TimePeriodMinute
mTimeframeTable(0).saveIntervalSeconds = 5

mTimeframeTable(1).timeframe.length = 5
mTimeframeTable(1).timeframe.units = TimePeriodMinute
mTimeframeTable(1).saveIntervalSeconds = 5

mTimeframeTable(2).timeframe.length = 15
mTimeframeTable(2).timeframe.units = TimePeriodMinute
mTimeframeTable(2).saveIntervalSeconds = 5

mTimeframeTable(3).timeframe.length = 60
mTimeframeTable(3).timeframe.units = TimePeriodMinute
mTimeframeTable(3).saveIntervalSeconds = 5

End Sub

'@===============================================================================
' IBarDataWriter Interface Members
'@===============================================================================

Private Property Get IBarDataWriter_ContractSpecifier() As ContractSpecifier
Set IBarDataWriter_ContractSpecifier = mContractSpec
End Property

Private Sub IBarDataWriter_ReleaseDataStore()
On Error GoTo Err
finish
mHistDataServiceProvider.dataWriterFinished Me
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "IBarDataWriter_ReleaseDataStore", _
                Err.Description
End Sub

Private Sub IBarDataWriter_WriteBar( _
                ByVal data As Bar, _
                ByVal barType As BarTypes, _
                ByVal barLength As Long, _
                ByVal barUnit As TimePeriodUnits)
Dim index As Long

On Error GoTo Err

Set mBarWriter = mInstrument.CreateBarDataWriter

With data
    mBarWriter.WriteBar .timestamp, _
                        barType, _
                        barLength, _
                        .openValue, _
                        .highValue, _
                        .lowValue, _
                        .closeValue, _
                        .Volume, _
                        .tickVolume, _
                        0, _
                        Not mUseSynchronousWrites
    
    For index = 0 To UBound(mTimeframeTable)
        If mTimeframeTable(index).timeframe.length = barLength Then
            With mTimeframeTable(index)
                If .prevTimestamp <> data.timestamp Then
                    .barsWritten = .barsWritten + 1
                    .prevTimestamp = data.timestamp
                End If
                .writes = .writes + 1
            End With
            Exit For
        End If
    Next
    
End With

Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "IBarDataWriter_WriteBar", _
                Err.Description
End Sub

'@===============================================================================
' mBarWriter Event Handlers
'@===============================================================================

Private Sub mBarWriter_ConnectFailed(ByVal errorCode As Long, ByVal errorDesc As String)
On Error GoTo Err
mServiceConsumer.NotifyEvent StandardSPEventCodes.HDCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc, _
                        mHistDataServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "mBarWriter_ConnectFailed", _
                Err.Description
End Sub

Private Sub mBarWriter_Error(ByVal errorCode As Long, ByVal errorDesc As String)
On Error GoTo Err
mCommonServiceConsumer.ServiceProviderError errorCode, _
                                    errorDesc, _
                                    mHistDataServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "mBarWriter_Error", _
                Err.Description
End Sub

'@===============================================================================
' mTradingDB Event Handlers
'@===============================================================================

Private Sub mTradingDB_ConnectFailed(ByVal errorCode As Long, ByVal errorDesc As String)
On Error GoTo Err
mServiceConsumer.NotReady
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "mTradingDB_ConnectFailed", _
                Err.Description
End Sub

Private Sub mTradingDB_ConnectSucceeded()
Dim i As Long

On Error GoTo Err

mServiceConsumer.NotifyEvent StandardSPEventCodes.HDConnectedToDataSource, _
                        "Connected to data source", _
                        mHistDataServiceProvider.handle

For i = 0 To UBound(mTimeframeTable)
    mServiceConsumer.RequiredTimePeriod mTimeframeTable(i).timeframe, _
                                        mTimeframeTable(i).saveIntervalSeconds
Next

mServiceConsumer.Ready
mServiceConsumer.OutputBarfileCreated mConnectionString

getInstrument

Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "mTradingDB_ConnectSucceeded", _
                Err.Description
End Sub
'@===============================================================================
' Properties
'@===============================================================================

Friend Property Let HistDataServiceProvider( _
                ByVal value As HistDataServiceProvider)
Set mHistDataServiceProvider = value
End Property

Friend Property Let Path(ByVal value As String)
mPath = value
End Property

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub finish()
Dim s As String
Dim i As Long

s = mServiceProviderName & " (" & mContractSpec.localSymbol & ")"
For i = 0 To UBound(mTimeframeTable)
    s = s & vbCrLf & "    Total of " & mTimeframeTable(i).barsWritten & " bars (" & _
            TimePeriodToString(mTimeframeTable(i).timeframe) & _
            ") written in " & _
            mTimeframeTable(i).writes & " database writes"
Next
gLogger.Log LogLevelNormal, _
            s, _
            Me
Set mBarWriter = Nothing
Set mServiceConsumer = Nothing
End Sub

Friend Sub initialise( _
                ByVal Name As String, _
                ByVal connectionString As String, _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal barDataOutputServiceConsumer As TradeBuildSP.IBarDataOutputServiceConsumer, _
                ByVal HistDataServiceProvider As TradeBuildSP.IBarDataServiceProvider, _
                ByVal contractSpec As ContractSpecifier, _
                ByVal useSynchronousWrites As Boolean)
mServiceProviderName = Name

Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = barDataOutputServiceConsumer

mUseSynchronousWrites = useSynchronousWrites

mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString)

Set mHistDataServiceProvider = HistDataServiceProvider
Set mContractSpec = contractSpec

End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Friend Sub getInstrument()
Dim lInstrumentFactory As InstrumentFactory

Set lInstrumentFactory = mTradingDB.InstrumentFactory

Set mInstrument = lInstrumentFactory.loadBySpecifier(mContractSpec)
Debug.Assert Not mInstrument Is Nothing

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
mHistDataServiceProvider.dataWriterFinished Me
End Sub



