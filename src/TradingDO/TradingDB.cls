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

Implements IStringable

'@================================================================================
' Events
'@================================================================================

Event Connected()

Event ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)

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

Private mConnectionParams As ConnectionParams

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
Const ProcName As String = "Class_Terminate"
On Error GoTo Err

gLogger.Log "TradingDB terminated: number of connection requests serviced: " & mNumberOfConnections, ProcName, ModuleName, LogLevelDetail

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IStringable Interface Members
'@================================================================================

Private Function IStringable_ToString() As String
IStringable_ToString = ToString
End Function

'@===============================================================================
' mConnect Event Handlers
'@===============================================================================

Private Sub mConnector_Connected( _
                ByVal pConnection As ADODB.connection)
Const ProcName As String = "mConnector_Connected"
On Error GoTo Err

gLogger.Log "TradingDB " & mDbID & ": initial Connect succeeded", ProcName, ModuleName
pConnection.Close
RaiseEvent Connected

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mConnector_ConnectFailed(ByVal errorCode As Long, ByVal errorDesc As String)
Const ProcName As String = "mConnector_ConnectFailed"
On Error GoTo Err

gLogger.Log "TradingDB " & mDbID & ": initial Connect failed", ProcName, ModuleName, LogLevelSevere
RaiseEvent ConnectFailed(errorCode, errorDesc)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ConnectionString() As String
ConnectionString = mConnectionParams.ConnectionString
End Property

Friend Property Get DatabaseType() As DatabaseTypes
DatabaseType = mConnectionParams.DatabaseType
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
                ByVal Specifier As IContractSpecifier, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal maxNumberOfBars As Long, _
                Optional ByVal fromDate As Date, _
                Optional ByVal toDate As Date = MaxDateValue, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal barType As BarTypes = BarTypeTrade, _
                Optional ByVal pListener As IBarFetchListener) As Bars
Const ProcName As String = "FetchBars"
On Error GoTo Err

Dim lInstrument As instrument
Set lInstrument = InstrumentFactory.LoadBySpecifier(Specifier)
AssertArgument Not lInstrument Is Nothing, "The specified contract is not known"

Set FetchBars = lInstrument.FetchBars(BarTimePeriod, _
                                        maxNumberOfBars, _
                                        fromDate, _
                                        toDate, _
                                        customSessionStartTime, _
                                        customSessionEndTime, _
                                        includeBarsOutsideSession, _
                                        barType, _
                                        pListener)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchBarsAsync( _
                ByVal Specifier As IContractSpecifier, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal maxNumberOfBars As Long, _
                Optional ByVal fromDate As Date, _
                Optional ByVal toDate As Date = MaxDateValue, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal barType As BarTypes = BarTypeTrade, _
                Optional ByVal cookie As Variant, _
                Optional ByVal pListener As IBarFetchListener, _
                Optional ByVal taskName As String) As IFuture
Const ProcName As String = "FetchbarsAsync"
On Error GoTo Err

Dim fbt As New FetchBarsTask
fbt.Initialise Me, _
                Specifier, _
                BarTimePeriod, _
                maxNumberOfBars, _
                fromDate, _
                toDate, _
                customSessionStartTime, _
                customSessionEndTime, _
                includeBarsOutsideSession, _
                barType, _
                pListener, _
                cookie
                
If taskName = "" Then taskName = "FetchBarsAsync/" & Specifier.LocalSymbol & "(" & BarTimePeriod.ToString & ")" & FormatTimestamp(fromDate, TimestampNoMillisecs) & "-" & FormatTimestamp(toDate, TimestampNoMillisecs) & "(" & gGetSequenceNumber & ")"
StartTask fbt, PriorityNormal, taskName, cookie

Set FetchBarsAsync = fbt.BarsFuture

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchContracts( _
                ByVal contractSpec As IContractSpecifier) As IContracts
Const ProcName As String = "FetchContracts"
On Error GoTo Err

Dim instruments As DataObjects
Set instruments = Me.InstrumentFactory.QueryObjectsBySpec(contractSpec)

Dim builder As ContractsBuilder
Set builder = New ContractsBuilder

Dim instrument As BusinessDataObject
For Each instrument In instruments
    builder.Add gContractFromInstrument(instrument)
Next

Set FetchContracts = builder.Contracts

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchContractsAsync( _
                ByVal pContractSpec As IContractSpecifier, _
                Optional ByVal pListener As IContractFetchListener, _
                Optional ByVal pCookie As Variant, _
                Optional ByVal pTaskName As String) As IFuture
Const ProcName As String = "FetchContractsAsync"
On Error GoTo Err

Dim crt As New ContractsRetrievalTask
crt.Initialise Me, pContractSpec, pListener, pCookie
If pTaskName = "" Then pTaskName = "LoadContractsAsync/" & pContractSpec.ToString & "(" & gGetSequenceNumber & ")"
StartTask crt, PriorityNormal, pTaskName, pCookie
                            
Set FetchContractsAsync = crt.ContractsFuture

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchTicks( _
                ByVal Specifier As IContractSpecifier, _
                ByVal fromDate As Date, _
                ByVal toDate As Date) As TickStream
Const ProcName As String = "FetchTicks"
On Error GoTo Err

Dim lInstrument As instrument
Set lInstrument = InstrumentFactory.LoadBySpecifier(Specifier)
AssertArgument Not lInstrument Is Nothing, "The specified contract is not known"

Set FetchTicks = lInstrument.CreateTickStream(fromDate, toDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchTicksAsync( _
                ByVal Specifier As IContractSpecifier, _
                ByVal fromDate As Date, _
                ByVal toDate As Date, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "FetchTicksAsync"
On Error GoTo Err

Dim ftt As New FetchTicksTask
ftt.Initialise Me, Specifier, fromDate, toDate
If taskName = "" Then taskName = "FetchTicksAsync/" & Specifier.LocalSymbol & "/" & FormatTimestamp(fromDate, TimestampNoMillisecs) & "-" & FormatTimestamp(toDate, TimestampNoMillisecs) & "(" & gGetSequenceNumber & ")"
Set FetchTicksAsync = StartTask(ftt, PriorityNormal, taskName, cookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetConnector( _
                Optional ByVal cursorlocation As CursorLocationEnum = CursorLocationEnum.adUseClient) As Connector
Const ProcName As String = "GetConnector"
On Error GoTo Err

Set GetConnector = New Connector
GetConnector.Initialise Me, cursorlocation
mNumberOfConnections = mNumberOfConnections + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pConnectionParams As ConnectionParams)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mConnectionParams = pConnectionParams

Set mExchangeFactory = New ExchangeFactory
mExchangeFactory.Initialise Me

Set mInstrumentFactory = New InstrumentFactory
mInstrumentFactory.Initialise Me

Set mInstrumentClassFactory = New InstrumentClassFactory
mInstrumentClassFactory.Initialise Me

Set mTimeZoneFactory = New TimeZoneFactory
mTimeZoneFactory.Initialise Me

mDbID = GenerateGUIDString
gLogger.Log "TradingDB " & mDbID & ": initial Connect", ProcName, ModuleName, LogLevelNormal, mConnectionParams.ToString
Set mConnector = GetConnector(CursorLocationEnum.adUseClient)
mConnector.ConnectAsync False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ToString() As String
ToString = mConnectionParams.ToString
End Function

'@================================================================================
' Helper Functions
'@================================================================================


