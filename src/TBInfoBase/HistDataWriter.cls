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

'@===============================================================================
' Member variables
'@===============================================================================

Private WithEvents mTradingDB       As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1
Private mConnectionString           As String

Private mHistDataServiceProvider    As histDataServiceProvider

Private mCommonServiceConsumer      As ICommonServiceConsumer
Private mServiceConsumer            As IBarDataOutputServiceConsumer
Private mServiceProviderName        As String

Private mContractSpec               As IContractSpecifier
Private WithEvents mBarWriter       As BarDataWriter
Attribute mBarWriter.VB_VarHelpID = -1

Private mPath                       As String

Private mBarsWritten                As Long
Private mWrites                     As Long

Private mInstrument                 As instrument

'@===============================================================================
' Class Event Handlers
'@===============================================================================

'@===============================================================================
' IBarDataWriter Interface Members
'@===============================================================================

Private Property Get IBarDataWriter_ContractSpecifier() As TradeBuildSP.IContractSpecifier

End Property

Private Sub IBarDataWriter_ReleaseDataStore()
On Error GoTo Err
finish
mHistDataServiceProvider.dataWriterFinished Me
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataWriter" & "::" & "IBarDataWriter_ReleaseDataStore", _
                Err.Description
End Sub

Private Sub IBarDataWriter_WriteBar( _
                ByRef data As TradeBuildSP.Bar, _
                ByVal barType As BarTypes, _
                ByVal barLength As Long, _
                ByVal barUnit As TimePeriodUnits)
Static prevTimestamp As Date
On Error GoTo Err

Set mBarWriter = mInstrument.CreateBarDataWriter

With data
    mBarWriter.WriteBar .timestamp, _
                        barType, _
                        barLength, _
                        .openPrice, _
                        .highPrice, _
                        .lowPrice, _
                        .closePrice, _
                        .volume, _
                        .tickVolume, _
                        0
    If prevTimestamp <> .timestamp Then
        mBarsWritten = mBarsWritten + 1
        prevTimestamp = .timestamp
    End If
End With
mWrites = mWrites + 1
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataWriter" & "::" & "IBarDataWriter_WriteBar", _
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
                ProjectName & "." & "HistDataWriter" & "::" & "mBarWriter_ConnectFailed", _
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
                ProjectName & "." & "HistDataWriter" & "::" & "mBarWriter_Error", _
                Err.Description
End Sub

'@===============================================================================
' mTradingDB Event Handlers
'@===============================================================================

Private Sub mTradingDB_ConnectFailed(ByVal errorCode As Long, ByVal errorDesc As String)
On Error GoTo Err
mServiceConsumer.Ready = False
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataWriter" & "::" & "mBarWriter_NotReady", _
                Err.Description
End Sub

Private Sub mTradingDB_ConnectSucceeded()
On Error GoTo Err

mServiceConsumer.NotifyEvent StandardSPEventCodes.HDConnectedToDataSource, _
                        "Connected to data source", _
                        mHistDataServiceProvider.handle

mServiceConsumer.Ready = True
mServiceConsumer.OutputBarfileCreated mConnectionString

getInstrument

Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataWriter" & "::" & "mBarWriter_Ready", _
                Err.Description
End Sub
'@===============================================================================
' Properties
'@===============================================================================

Friend Property Let histDataServiceProvider( _
                ByVal value As histDataServiceProvider)
Set mHistDataServiceProvider = value
End Property

Friend Property Let Path(ByVal value As String)
mPath = value
End Property

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub finish()
mCommonServiceConsumer.NotifyListeners mServiceProviderName & vbCrLf & _
                            "    Total bars written (" & mContractSpec.localSymbol & "): " & mBarsWritten & _
                            " in " & mWrites & " database writes", _
                            TradeBuildSP.StandardListenValueTypes.Log, _
                            Me
Set mBarWriter = Nothing
Set mServiceConsumer = Nothing
End Sub

Friend Sub initialise( _
                ByVal Name As String, _
                ByVal connectionString As String, _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal barDataOutputServiceConsumer As TradeBuildSP.IBarDataOutputServiceConsumer, _
                ByVal histDataServiceProvider As TradeBuildSP.IBarDataServiceProvider, _
                ByVal contractSpec As TradeBuildSP.IContractSpecifier)
mServiceProviderName = Name

Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = barDataOutputServiceConsumer

mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString)

Set mHistDataServiceProvider = histDataServiceProvider
Set mContractSpec = contractSpec.Clone

End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Friend Sub getInstrument()
Dim lInstrumentFactory As InstrumentFactory

Set lInstrumentFactory = mTradingDB.InstrumentFactory

With mContractSpec
    Set mInstrument = lInstrumentFactory.loadBySpec( _
                                .localSymbol, _
                                .Symbol, _
                                .Exchange, _
                                gSecTypeToCategory(.secType), _
                                .CurrencyCode, _
                                Left$(.Expiry, 6), _
                                .Strike, _
                                gOptRightToString(.Right))
End With
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



