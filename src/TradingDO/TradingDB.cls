VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradingDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
' @see
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName As String = "TradingDO26"
Private Const ModuleName As String = "TradingDB"

'@================================================================================
' Member variables
'@================================================================================

Private mDatabaseType As DatabaseTypes
Private mConnectionString As String

Private WithEvents mConnector As Connector
Attribute mConnector.VB_VarHelpID = -1

Private mExchangeFactory As ExchangeFactory
Private mInstrumentFactory As InstrumentFactory
Private mInstrumentClassFactory As InstrumentClassFactory
Private mTimeZoneFactory As TimeZoneFactory

Private mNumberOfConnections As Long

Private mDbID As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
gLogger.Log LogLevelDetail, "TradingDB: number of connection requests serviced: " & mNumberOfConnections
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@===============================================================================
' mConnect Event Handlers
'@===============================================================================

Private Sub mConnector_Connected( _
                ByVal pConnection As ADODB.connection)
gLogger.Log LogLevelDetail, "TradingDB " & mDbID & ": initial connect succeeded"
'RaiseEvent ConnectSucceeded
pConnection.Close
End Sub

Private Sub mConnector_ConnectFailed(ByVal errorCode As Long, ByVal errorDesc As String)
gLogger.Log LogLevelDetail, "TradingDB " & mDbID & ": initial connect failed"
'RaiseEvent ConnectFailed(errorCode, errorDesc)
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get connectionString() As String
connectionString = mConnectionString
End Property

Public Property Get databaseType() As DatabaseTypes
databaseType = mDatabaseType
End Property

Public Property Get ExchangeFactory() As ExchangeFactory
Set ExchangeFactory = mExchangeFactory
End Property

Public Property Get InstrumentFactory() As InstrumentFactory
Set InstrumentFactory = mInstrumentFactory
End Property

Public Property Get InstrumentClassFactory() As InstrumentClassFactory
Set InstrumentClassFactory = mInstrumentClassFactory
End Property

Public Property Get TimeZoneFactory() As TimeZoneFactory
Set TimeZoneFactory = mTimeZoneFactory
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function fetchBars( _
                ByVal specifier As ContractSpecifier, _
                ByVal barTimePeriod As TimePeriod, _
                ByVal maxNumberOfBars As Long, _
                Optional ByVal fromDate As Date, _
                Optional ByVal toDate As Date = MaxDateValue, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal barType As BarTypes = BarTypeTrade) As Bars
Dim lInstrument As instrument

Dim failpoint As Long
On Error GoTo Err

Set lInstrument = InstrumentFactory.loadBySpecifier(specifier)
If lInstrument Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "fetchBars", _
            "The specified contract is not known"
End If

Set fetchBars = lInstrument.fetchBars(barTimePeriod, _
                maxNumberOfBars, _
                fromDate, _
                toDate, _
                customSessionStartTime, _
                customSessionEndTime, _
                includeBarsOutsideSession, _
                barType)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "fetchBars" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description
End Function

Public Function fetchbarsAsync( _
                ByVal specifier As ContractSpecifier, _
                ByVal barTimePeriod As TimePeriod, _
                ByVal maxNumberOfBars As Long, _
                Optional ByVal fromDate As Date, _
                Optional ByVal toDate As Date = MaxDateValue, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal barType As BarTypes = BarTypeTrade, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim fbt As New FetchBarsTask
Dim failpoint As Long
On Error GoTo Err

fbt.initialise Me, _
                specifier, _
                barTimePeriod, _
                maxNumberOfBars, _
                fromDate, _
                toDate, _
                customSessionStartTime, _
                customSessionEndTime, _
                includeBarsOutsideSession, _
                barType
                
If taskName = "" Then taskName = "FetchBarsAsync/" & specifier.localSymbol & "(" & barTimePeriod.ToString & ")" & FormatTimestamp(fromDate, TimestampNoMillisecs) & "-" & FormatTimestamp(toDate, TimestampNoMillisecs) & "(" & gGetSequenceNumber & ")"
Set fetchbarsAsync = StartTask(fbt, _
                            PriorityNormal, _
                            taskName, _
                            cookie)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "fetchbarsAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function fetchTicks( _
                ByVal specifier As ContractSpecifier, _
                ByVal fromDate As Date, _
                ByVal toDate As Date) As TickStream
Dim lInstrument As instrument

Dim failpoint As Long
On Error GoTo Err

Set lInstrument = InstrumentFactory.loadBySpecifier(specifier)
If lInstrument Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "fetchTicks", _
            "The specified contract is not known"
End If

Set fetchTicks = lInstrument.createTickStream(fromDate, toDate)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "fetchTicks" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function fetchTicksAsync( _
                ByVal specifier As ContractSpecifier, _
                ByVal fromDate As Date, _
                ByVal toDate As Date, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim ftt As New FetchTicksTask
Dim failpoint As Long
On Error GoTo Err

ftt.initialise Me, _
                specifier, _
                fromDate, _
                toDate
If taskName = "" Then taskName = "FetchTicksAsync/" & specifier.localSymbol & "/" & FormatTimestamp(fromDate, TimestampNoMillisecs) & "-" & FormatTimestamp(toDate, TimestampNoMillisecs) & "(" & gGetSequenceNumber & ")"
Set fetchTicksAsync = StartTask(ftt, _
                        PriorityNormal, _
                        taskName, _
                        cookie)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "fetchTicksAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Friend Function getConnector( _
                Optional ByVal cursorlocation As CursorLocationEnum = CursorLocationEnum.adUseClient, _
                Optional ByVal Options As ConnectOptionEnum) As Connector
Set getConnector = New Connector
getConnector.initialise Me, cursorlocation, Options
mNumberOfConnections = mNumberOfConnections + 1
End Function

Friend Sub initialise( _
                ByVal connectionString As String, _
                ByVal dbType As DatabaseTypes)
mConnectionString = connectionString
mDatabaseType = dbType

Set mExchangeFactory = New ExchangeFactory
mExchangeFactory.initialise Me

Set mInstrumentFactory = New InstrumentFactory
mInstrumentFactory.initialise Me

Set mInstrumentClassFactory = New InstrumentClassFactory
mInstrumentClassFactory.initialise Me

Set mTimeZoneFactory = New TimeZoneFactory
mTimeZoneFactory.initialise Me

mDbID = GenerateGUIDString
gLogger.Log LogLevelDetail, "TradingDB " & mDbID & ": initial connect"
Set mConnector = getConnector(CursorLocationEnum.adUseClient, ConnectOptionEnum.adAsyncConnect)
mConnector.connect
End Sub

Public Function loadContracts( _
                ByVal contractSpec As ContractSpecifier) As Contracts
Dim instruments As DataObjects
Dim instrument As BusinessDataObject
Dim builder As ContractsBuilder

Set instruments = Me.InstrumentFactory.queryObjectsBySpec(contractSpec)

Set builder = CreateContractsBuilder(contractSpec)

For Each instrument In instruments
    builder.AddContract gContractFromInstrument(instrument)
Next

Set loadContracts = builder.Contracts
End Function

Public Function loadContractsAsync( _
                ByVal contractSpec As ContractSpecifier, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim crt As New ContractsRetrievalTask
crt.initialise Me, contractSpec
If taskName = "" Then taskName = "LoadContractsAsync/" & contractSpec.ToString & "(" & gGetSequenceNumber & ")"
Set loadContractsAsync = StartTask(crt, _
                            PriorityNormal, _
                            taskName, _
                            cookie)
End Function

'@================================================================================
' Helper Functions
'@================================================================================


