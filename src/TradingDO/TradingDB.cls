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

Event ConnectSucceeded()
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

Private Const ProjectName As String = "TradingDO26"
Private Const ModuleName As String = "TradingDB"

'@================================================================================
' Member variables
'@================================================================================

Private mConnectionString As String

Private WithEvents mConnection As ADODB.Connection
Attribute mConnection.VB_VarHelpID = -1

Private mExchangeFactory As ExchangeFactory
Private mInstrumentFactory As InstrumentFactory
Private mInstrumentClassFactory As InstrumentClassFactory
Private mTimeZoneFactory As TimeZoneFactory

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@===============================================================================
' mConnection Event Handlers
'@===============================================================================

Private Sub mConnection_ConnectComplete( _
                            ByVal pError As ADODB.Error, _
                            adStatus As ADODB.EventStatusEnum, _
                            ByVal pConnection As ADODB.Connection)

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent ConnectFailed(pError.Number, _
                            ModuleName & ": connection error " & pError.NativeError & vbCrLf & _
                            pError.Description & vbCrLf & _
                            gGenerateErrorMessage(pConnection))
Else
    RaiseEvent ConnectSucceeded
    mConnection.Close
End If


End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get connectionString() As String
connectionString = mConnectionString
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
                ByVal barLength As Long, _
                ByVal barLengthUnits As TimePeriodUnits, _
                ByVal maxNumberOfBars As Long, _
                Optional ByVal fromDate As Date, _
                Optional ByVal toDate As Date = MaxDateValue, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal barType As BarTypes = BarTypeTrade) As bars
Dim lInstrument As instrument

Dim failpoint As Long
On Error GoTo Err

Set lInstrument = InstrumentFactory.loadBySpecifier(specifier)
If lInstrument Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "fetchbars", _
            "The specified contract is not known"
End If

Set FetchBars = lInstrument.FetchBars(barLength, _
                barLengthUnits, _
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
        ProjectName & "." & ModuleName & ":" & "fetchbars" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description
End Function

Public Function fetchbarsAsync( _
                ByVal specifier As ContractSpecifier, _
                ByVal barLength As Long, _
                ByVal barLengthUnits As TimePeriodUnits, _
                ByVal maxNumberOfBars As Long, _
                Optional ByVal fromDate As Date, _
                Optional ByVal toDate As Date = MaxDateValue, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal barType As BarTypes = BarTypeTrade) As TaskController
Dim fbt As New FetchBarsTask
fbt.initialise Me, _
                specifier, _
                barLength, _
                barLengthUnits, _
                maxNumberOfBars, _
                fromDate, _
                toDate, _
                customSessionStartTime, _
                customSessionEndTime, _
                includeBarsOutsideSession, _
                barType
Set fetchbarsAsync = StartTask(fbt, PriorityNormal)
End Function

Friend Sub initialise( _
                ByVal connectionString As String)
mConnectionString = connectionString
Set gLogger = GetLogger(InfoTypeTradingDO)

Set mExchangeFactory = New ExchangeFactory
mExchangeFactory.initialise Me

Set mInstrumentFactory = New InstrumentFactory
mInstrumentFactory.initialise Me

Set mInstrumentClassFactory = New InstrumentClassFactory
mInstrumentClassFactory.initialise Me

Set mTimeZoneFactory = New TimeZoneFactory
mTimeZoneFactory.initialise Me

Set mConnection = New ADODB.Connection
mConnection.CursorLocation = CursorLocationEnum.adUseClient
mConnection.Open mConnectionString, , , ConnectOptionEnum.adAsyncConnect
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


