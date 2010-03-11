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
'gLogger.Log LogLevelDetail, "TradingDB: number of connection requests serviced: " & mNumberOfConnections
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
pConnection.Close
End Sub

Private Sub mConnector_ConnectFailed(ByVal errorCode As Long, ByVal errorDesc As String)
gLogger.Log LogLevelDetail, "TradingDB " & mDbID & ": initial connect failed"
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get ConnectionString() As String
ConnectionString = mConnectionString
End Property

Public Property Get DatabaseType() As DatabaseTypes
DatabaseType = mDatabaseType
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

Public Function FetchBars( _
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

Const ProcName As String = "FetchBars"
Dim failpoint As String
On Error GoTo Err

Set lInstrument = InstrumentFactory.LoadBySpecifier(specifier)
If lInstrument Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "The specified contract is not known"
End If

Set FetchBars = lInstrument.FetchBars(barTimePeriod, _
                maxNumberOfBars, _
                fromDate, _
                toDate, _
                customSessionStartTime, _
                customSessionEndTime, _
                includeBarsOutsideSession, _
                barType)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function FetchbarsAsync( _
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

Const ProcName As String = "FetchbarsAsync"
Dim failpoint As String
On Error GoTo Err

fbt.Initialise Me, _
                specifier, _
                barTimePeriod, _
                maxNumberOfBars, _
                fromDate, _
                toDate, _
                customSessionStartTime, _
                customSessionEndTime, _
                includeBarsOutsideSession, _
                barType
                
If taskName = "" Then taskName = "FetchBarsAsync/" & specifier.LocalSymbol & "(" & barTimePeriod.ToString & ")" & FormatTimestamp(fromDate, TimestampNoMillisecs) & "-" & FormatTimestamp(toDate, TimestampNoMillisecs) & "(" & gGetSequenceNumber & ")"
Set FetchbarsAsync = StartTask(fbt, _
                            PriorityNormal, _
                            taskName, _
                            cookie)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function FetchTicks( _
                ByVal specifier As ContractSpecifier, _
                ByVal fromDate As Date, _
                ByVal toDate As Date) As TickStream
Dim lInstrument As instrument

Const ProcName As String = "FetchTicks"
Dim failpoint As String
On Error GoTo Err

Set lInstrument = InstrumentFactory.LoadBySpecifier(specifier)
If lInstrument Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "The specified contract is not known"
End If

Set FetchTicks = lInstrument.CreateTickStream(fromDate, toDate)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function FetchTicksAsync( _
                ByVal specifier As ContractSpecifier, _
                ByVal fromDate As Date, _
                ByVal toDate As Date, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim ftt As New FetchTicksTask

Const ProcName As String = "FetchTicksAsync"
Dim failpoint As String
On Error GoTo Err

ftt.Initialise Me, _
                specifier, _
                fromDate, _
                toDate
If taskName = "" Then taskName = "FetchTicksAsync/" & specifier.LocalSymbol & "/" & FormatTimestamp(fromDate, TimestampNoMillisecs) & "-" & FormatTimestamp(toDate, TimestampNoMillisecs) & "(" & gGetSequenceNumber & ")"
Set FetchTicksAsync = StartTask(ftt, _
                        PriorityNormal, _
                        taskName, _
                        cookie)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Function GetConnector( _
                Optional ByVal cursorlocation As CursorLocationEnum = CursorLocationEnum.adUseClient, _
                Optional ByVal Options As ConnectOptionEnum) As Connector
Set GetConnector = New Connector
GetConnector.Initialise Me, cursorlocation, Options
mNumberOfConnections = mNumberOfConnections + 1
End Function

Friend Sub Initialise( _
                ByVal ConnectionString As String, _
                ByVal dbType As DatabaseTypes)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

mConnectionString = ConnectionString
mDatabaseType = dbType

Set mExchangeFactory = New ExchangeFactory
mExchangeFactory.Initialise Me

Set mInstrumentFactory = New InstrumentFactory
mInstrumentFactory.Initialise Me

Set mInstrumentClassFactory = New InstrumentClassFactory
mInstrumentClassFactory.Initialise Me

Set mTimeZoneFactory = New TimeZoneFactory
mTimeZoneFactory.Initialise Me

mDbID = GenerateGUIDString
gLogger.Log LogLevelDetail, "TradingDB " & mDbID & ": initial connect"
Set mConnector = GetConnector(CursorLocationEnum.adUseClient, ConnectOptionEnum.adAsyncConnect)
mConnector.connect

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function LoadContracts( _
                ByVal contractSpec As ContractSpecifier) As Contracts
Dim instruments As DataObjects
Dim instrument As BusinessDataObject
Dim builder As ContractsBuilder

Const ProcName As String = "LoadContracts"
Dim failpoint As String
On Error GoTo Err

Set instruments = Me.InstrumentFactory.QueryObjectsBySpec(contractSpec)

Set builder = CreateContractsBuilder(contractSpec)

For Each instrument In instruments
    builder.AddContract gContractFromInstrument(instrument)
Next

Set LoadContracts = builder.Contracts

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function LoadContractsAsync( _
                ByVal contractSpec As ContractSpecifier, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim crt As New ContractsRetrievalTask
Const ProcName As String = "LoadContractsAsync"
Dim failpoint As String
On Error GoTo Err

crt.Initialise Me, contractSpec
If taskName = "" Then taskName = "LoadContractsAsync/" & contractSpec.ToString & "(" & gGetSequenceNumber & ")"
Set LoadContractsAsync = StartTask(crt, _
                            PriorityNormal, _
                            taskName, _
                            cookie)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


