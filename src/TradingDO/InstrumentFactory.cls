VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InstrumentFactory"
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
Private Const ModuleName As String = "InstrumentFactory"

'@================================================================================
' Member variables
'@================================================================================

Private mTradingDB As tradingDB
Private mConnectionString As String

Private mConn As ADODB.Connection
Private mRs As ADODB.recordset

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let connectionString( _
                ByVal value As String)
mConnectionString = value
End Property

Friend Property Let tradingDB( _
                ByVal value As tradingDB)
Set mTradingDB = value
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function loadByID( _
                ByVal id As Long) As instrument
Dim where As String
where = "ID=" & id
Set loadByID = loadDetails(where)
End Function

Public Function loadBySpec( _
                ByVal pShortName As String, _
                ByVal pSymbol As String, _
                ByVal pExchange As String, _
                ByVal pCategory As String, _
                ByVal pCurrency As String, _
                ByVal pExpiry As String, _
                ByVal pStrike As Double, _
                ByVal pRight As String) As instrument
Dim where As String
If pShortName <> "" Then where = "shortname='" & Replace(pShortName, "'", "''") & "'"
If pSymbol <> "" Then where = IIf(where = "", "", where & " AND ") & "Symbol='" & Replace(pSymbol, "'", "''") & "'"
If pExchange <> "" Then where = IIf(where = "", "", where & " AND ") & "Exchange='" & pExchange & "'"
If pCategory <> "" Then where = IIf(where = "", "", where & " AND ") & "Category='" & pCategory & "'"
If pCurrency <> "" Then where = IIf(where = "", "", where & " AND ") & "Currency='" & pCurrency & "'"
If pExpiry <> "" Then where = IIf(where = "", "", where & " AND ") & "ExpiryMonth='" & pExpiry & "'"
If pStrike <> 0 Then where = IIf(where = "", "", where & " AND ") & "StrikePrice=" & pStrike
If pRight <> "" Then where = IIf(where = "", "", where & " AND ") & "OptRight='" & pRight & "'"
Set loadBySpec = loadDetails(where)
End Function

Public Function loadSummaries( _
                Optional ByVal pShortName As String, _
                Optional ByVal pSymbol As String, _
                Optional ByVal pExchange As String, _
                Optional ByVal pCategory As String, _
                Optional ByVal pCurrency As String, _
                Optional ByVal pExpiry As String, _
                Optional ByVal pStrike As Double, _
                Optional ByVal pRight As String) As InstrumentSummaries
Dim where As String
If pShortName <> "" Then where = "shortname='" & Replace(pShortName, "'", "''") & "'"
If pSymbol <> "" Then where = IIf(where = "", "", where & " AND ") & "Symbol='" & Replace(pSymbol, "'", "''") & "'"
If pExchange <> "" Then where = IIf(where = "", "", where & " AND ") & "Exchange='" & pExchange & "'"
If pCategory <> "" Then where = IIf(where = "", "", where & " AND ") & "Category='" & pCategory & "'"
If pCurrency <> "" Then where = IIf(where = "", "", where & " AND ") & "Currency='" & pCurrency & "'"
If pExpiry <> "" Then where = IIf(where = "", "", where & " AND ") & "ExpiryMonth='" & pExpiry & "'"
If pStrike <> 0 Then where = IIf(where = "", "", where & " AND ") & "StrikePrice=" & pStrike
If pRight <> "" Then where = IIf(where = "", "", where & " AND ") & "OptRight='" & pRight & "'"
Set loadSummaries = searchWhere(where)
End Function

Public Function makeNew() As instrument
Dim SQL As String
Dim instrument As instrument

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

Set mConn = New ADODB.Connection
mConn.Open mConnectionString

failpoint = 200

Set mRs = New recordset
mRs.CursorLocation = adUseClient
SQL = "SELECT * " & _
        "FROM vInstrumentDetails " & _
        "WHERE 1=0"


mRs.Open SQL, mConn, adOpenStatic, adLockBatchOptimistic

Set instrument = New instrument

instrument.recordset = mRs
instrument.connectionString = mConnectionString
Set makeNew = instrument

mConn.Close
Set mRs.ActiveConnection = Nothing
Set mRs = Nothing

Exit Function

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "InstrumentFactory" & "::" & "makeNew" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function loadDetails(ByVal where As String) As instrument
Dim instrument As instrument
Dim SQL As String

Dim failpoint As Long
On Error GoTo Err

failpoint = 100 '===================================================

Set mConn = New ADODB.Connection
mConn.Open mConnectionString

failpoint = 200 '===================================================

Set mRs = New recordset
mRs.CursorLocation = adUseClient
SQL = "SELECT * " & _
        "FROM vInstrumentDetails " & _
        "WHERE " & where
mRs.Open SQL, mConn, adOpenStatic, adLockBatchOptimistic

failpoint = 300 '===================================================
If Not mRs.EOF Then
    Set instrument = New instrument
    instrument.connectionString = mConnectionString
    instrument.recordset = mRs
    Set loadDetails = instrument
Else
    Set loadDetails = Nothing
End If

mRs.ActiveConnection = Nothing
Set mRs = Nothing
mConn.Close
Set instrument = Nothing

Exit Function

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "InstrumentFactory" & "::" & "loadDetails" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Private Function searchWhere(ByVal where As String) As InstrumentSummaries
Dim summ As InstrumentSummaries
Dim SQL As String
Dim InstrumentSummary As InstrumentSummary


Dim failpoint As Long
On Error GoTo Err

failpoint = 100 '===================================================

Set mConn = New ADODB.Connection
mConn.Open mConnectionString

failpoint = 200 '===================================================

Set mRs = New recordset
mRs.CursorLocation = adUseClient
SQL = "SELECT ID, Name, Symbol, Category, ExpiryMonth " & _
        "FROM vInstrumentDetails " & _
        IIf(where <> "", "WHERE " & where, "") & _
        "ORDER BY InstrumentClassName, ExpiryDate"


mRs.Open SQL, mConn, adOpenStatic, adLockBatchOptimistic

Set summ = New InstrumentSummaries

Do While Not mRs.EOF
    Set InstrumentSummary = New InstrumentSummary
    InstrumentSummary.id = mRs("ID")
    InstrumentSummary.name = Trim$(mRs("Name"))
    InstrumentSummary.symbol = Trim$(mRs("Symbol"))
    InstrumentSummary.expiryMonth = Nz(mRs("ExpiryMonth"), "")
    InstrumentSummary.category = Trim$(mRs("Category"))
    summ.add InstrumentSummary
    Set InstrumentSummary = Nothing
    mRs.MoveNext
Loop

mConn.Close
Set mRs.ActiveConnection = Nothing

Set searchWhere = summ

Exit Function

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "InstrumentFactory" & "::" & "searchWhere" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description


End Function










