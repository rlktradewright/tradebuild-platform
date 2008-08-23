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
Implements TaskCompletionListener

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

Private mTradingDB                  As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1
Private mConnectionString           As String

Private mHistDataServiceProvider    As HistDataServiceProvider

Private mCommonServiceConsumer      As ICommonServiceConsumer
Private mServiceConsumer            As IBarDataOutputServiceConsumer
Private mServiceProviderName        As String

Private mContract                   As Contract

Private mPath                       As String

Private mInstrument                 As Instrument

Private mTimeframeTable(3)          As TimeframeEntry

Private mUseSynchronousWrites       As Boolean

Private WithEvents mInstrumentTC    As TaskController
Attribute mInstrumentTC.VB_VarHelpID = -1

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
Set mTimeframeTable(0).timeframe = GetTimePeriod(1, TimePeriodMinute)
mTimeframeTable(0).saveIntervalSeconds = 5

Set mTimeframeTable(1).timeframe = GetTimePeriod(5, TimePeriodMinute)
mTimeframeTable(1).saveIntervalSeconds = 5

Set mTimeframeTable(2).timeframe = GetTimePeriod(15, TimePeriodMinute)
mTimeframeTable(2).saveIntervalSeconds = 5

Set mTimeframeTable(3).timeframe = GetTimePeriod(60, TimePeriodMinute)
mTimeframeTable(3).saveIntervalSeconds = 5

End Sub

'@===============================================================================
' IBarDataWriter Interface Members
'@===============================================================================

Private Property Get IBarDataWriter_ContractSpecifier() As ContractSpecifier
Set IBarDataWriter_ContractSpecifier = mContract.specifier
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
                ByVal barTimePeriod As TimePeriod)
Dim index As Long
Dim tc As TaskController

On Error GoTo Err

With data
    If mUseSynchronousWrites Then
        mInstrument.WriteBar .TimeStamp, _
                            barType, _
                            barTimePeriod.length, _
                            .openValue, _
                            .highValue, _
                            .lowValue, _
                            .closeValue, _
                            .volume, _
                            .tickVolume, _
                            .openInterest
    Else
        Set tc = mInstrument.WriteBarAsync(.TimeStamp, _
                                            barType, _
                                            barTimePeriod.length, _
                                            .openValue, _
                                            .highValue, _
                                            .lowValue, _
                                            .closeValue, _
                                            .volume, _
                                            .tickVolume, _
                                            .openInterest)
        tc.addTaskCompletionListener Me
    End If
    For index = 0 To UBound(mTimeframeTable)
        If mTimeframeTable(index).timeframe.length = barTimePeriod.length Then
            With mTimeframeTable(index)
                If .prevTimestamp <> data.TimeStamp Then
                    .barsWritten = .barsWritten + 1
                    .prevTimestamp = data.TimeStamp
                End If
                .writes = .writes + 1
            End With
            Exit For
        End If
    Next
    
End With

Exit Sub
Err:
gLogger.Log LogLevelSevere, ProjectName & ":" & ModuleName & ":IBarDataWriter_WriteBar: error (" & Err.Number & "): " & Err.Description
mCommonServiceConsumer.ServiceProviderError Err.Number, _
                                            Err.Description, _
                                            mHistDataServiceProvider.handle
End Sub

'@===============================================================================
' TaskCompletionListener Interface Members
'@===============================================================================

Private Sub TaskCompletionListener_taskCompleted( _
                ev As TWUtilities30.TaskCompletionEvent)
If ev.errorNumber <> 0 Then
    gLogger.Log LogLevelSevere, ProjectName & ":" & ModuleName & ":TaskCompletionListener_taskCompleted: error (" & Err.Number & "): " & Err.Description
    mCommonServiceConsumer.ServiceProviderError ev.errorNumber, _
                                        ev.errorMessage, _
                                        mHistDataServiceProvider.handle
End If
End Sub

'@===============================================================================
' mInstrumentTC Event Handlers
'@===============================================================================

Private Sub mInstrumentTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
If ev.cancelled Then
ElseIf ev.errorNumber <> 0 Then
    mServiceConsumer.NotifyEvent HDRequestInvalid, _
                                "Error " & ev.errorNumber & ": " & ev.errorMessage, _
                                mHistDataServiceProvider.handle
ElseIf IsEmpty(ev.result) Then
    mServiceConsumer.NotifyEvent HDRequestInvalid, _
                                "Error: specified instrument does not exist: " & mContract.specifier.toString, _
                                mHistDataServiceProvider.handle
Else
    Set mInstrument = ev.result
    
    requestTimeframes
    
    mServiceConsumer.Ready
    mServiceConsumer.OutputBarfileCreated mConnectionString
End If
Set mInstrumentTC = Nothing
End Sub

''@===============================================================================
'' mTradingDB Event Handlers
''@===============================================================================
'
'Private Sub mTradingDB_ConnectFailed(ByVal errorCode As Long, ByVal errorDesc As String)
'On Error GoTo Err
'mServiceConsumer.NotifyEvent StandardSPEventCodes.HDCantConnectDataSource, _
'                        "error " & errorCode & ": " & errorDesc, _
'                        mHistDataServiceProvider.handle
'mServiceConsumer.NotReady
'Exit Sub
'Err:
'handleFatalError Err.Number, _
'                ProjectName & "." & ModuleName & "::" & "mTradingDB_ConnectFailed", _
'                Err.Description
'End Sub
'
'Private Sub mTradingDB_ConnectSucceeded()
'Dim i As Long
'
'On Error GoTo Err
'
'mServiceConsumer.NotifyEvent StandardSPEventCodes.HDConnectedToDataSource, _
'                        "Connected to data source", _
'                        mHistDataServiceProvider.handle
'
'gLogger.Log LogLevelDetail, "HistDataWriter: requesting " & (UBound(mTimeframeTable) + 1) & " timeframes"
'
'For i = 0 To UBound(mTimeframeTable)
'    mServiceConsumer.RequiredTimePeriod mTimeframeTable(i).timeframe, _
'                                        mTimeframeTable(i).saveIntervalSeconds
'Next
'
'Set mInstrumentTC = mTradingDB.InstrumentFactory.loadBySpecifierAsync(mContractSpec)
'
'Exit Sub
'Err:
'handleFatalError Err.Number, _
'                ProjectName & "." & ModuleName & "::" & "mTradingDB_ConnectSucceeded", _
'                Err.Description
'End Sub

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

Dim failpoint As Long
On Error GoTo Err

If Not mInstrumentTC Is Nothing Then
    mInstrumentTC.cancelTask
    Set mInstrumentTC = Nothing
End If

s = mServiceProviderName & " (" & mContract.specifier.localSymbol & ")"
For i = 0 To UBound(mTimeframeTable)
    s = s & vbCrLf & "    Total of " & mTimeframeTable(i).barsWritten & " bars (" & _
            mTimeframeTable(i).timeframe.toString & _
            ") written in " & _
            mTimeframeTable(i).writes & " database writes"
Next
gLogger.Log LogLevelNormal, _
            s, _
            Me
Set mServiceConsumer = Nothing

Exit Sub

Err:
gLogger.Log LogLevelSevere, _
        "Error " & Err.Number & " at " & _
        ProjectName & "." & ModuleName & ":" & "finish" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, "") & vbCrLf & _
        Err.Description

End Sub

Friend Sub initialise( _
                ByVal Name As String, _
                ByVal dbType As DatabaseTypes, _
                ByVal connectionString As String, _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal barDataOutputServiceConsumer As TradeBuildSP.IBarDataOutputServiceConsumer, _
                ByVal HistDataServiceProvider As TradeBuildSP.IBarDataServiceProvider, _
                ByVal pContract As Contract, _
                ByVal useSynchronousWrites As Boolean)
mServiceProviderName = Name

Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = barDataOutputServiceConsumer

mUseSynchronousWrites = useSynchronousWrites

gLogger.Log LogLevelDetail, "HistDataWriter: connecting to database"

mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString, dbType)

Set mHistDataServiceProvider = HistDataServiceProvider
Set mContract = pContract

If gLogger.isLoggable(LogLevelDetail) Then gLogger.Log LogLevelDetail, "HistDataWriter: retrieving instrument object for " & mContract.specifier.toString

Set mInstrumentTC = mTradingDB.InstrumentFactory.loadBySpecifierAsync(mContract.specifier)

End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)
gLogger.Log LogLevelSevere, Source & ": error (" & Number & "): " & Description

mCommonServiceConsumer.FatalServiceProviderError Number, _
                            Source, _
                            Description, _
                            mHistDataServiceProvider.handle

finish
mHistDataServiceProvider.dataWriterFinished Me
End Sub

Private Sub requestTimeframes()
Dim i As Long

gLogger.Log LogLevelDetail, "HistDataWriter: requesting " & (UBound(mTimeframeTable) + 1) & " timeframes"

For i = 0 To UBound(mTimeframeTable)
    mServiceConsumer.RequiredTimePeriod mTimeframeTable(i).timeframe, _
                                        mTimeframeTable(i).saveIntervalSeconds
Next

End Sub

