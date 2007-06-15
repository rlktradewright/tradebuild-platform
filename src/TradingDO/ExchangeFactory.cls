VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExchangeFactory"
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
Private Const ModuleName As String = "ExchangeFactory"

'@================================================================================
' Member variables
'@================================================================================

Private mDB As TradingDB

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

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal db As TradingDB)
Set mDB = db
End Sub

Public Function loadByName(ByVal pName As String) As Exchange
Dim where As String
where = "Name='" & Replace(pName, "'", "''") & "'"
Set loadByName = loadDetails(where)
End Function

Public Function loadByID(ByVal id As Long) As Exchange
Dim where As String
where = "ID=" & id
Set loadByID = loadDetails(where)
End Function

Public Function makeNew() As Exchange
Dim SQL As String
Dim Exchange As Exchange

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

Set mConn = New ADODB.Connection
mConn.Open mDB.connectionString

failpoint = 200

Set mRs = New recordset
mRs.CursorLocation = adUseClient
SQL = "SELECT * " & _
        "FROM exchange " & _
        "WHERE 1=0"


mRs.Open SQL, mConn, adOpenStatic, adLockBatchOptimistic

Set Exchange = New Exchange
Exchange.initialise mDB
Exchange.recordset = mRs
Set makeNew = Exchange

mConn.Close
Set mRs.ActiveConnection = Nothing
Set mRs = Nothing

Exit Function

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "ExchangeFactory" & "::" & "makeNew" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function search(ByVal criterion As String) As DataObjectSummaries
Dim summ As DataObjectSummaries
Dim SQL As String
Dim where As String
Dim lDataObjectSummary As DataObjectSummary


Dim failpoint As Long
On Error GoTo Err

where = IIf(criterion = "", "", "WHERE name LIKE '" & criterion & "' ")
failpoint = 100

Set mConn = New ADODB.Connection
mConn.Open mDB.connectionString

failpoint = 200

Set mRs = New recordset
mRs.CursorLocation = adUseClient
SQL = "SELECT ID, Name " & _
        "FROM exchange " & _
        where & _
        "ORDER BY Name"


mRs.Open SQL, mConn, adOpenStatic, adLockBatchOptimistic

Set summ = New DataObjectSummaries

Do While Not mRs.EOF
    Set lDataObjectSummary = New DataObjectSummary
    lDataObjectSummary.id = mRs("ID")
    lDataObjectSummary.name = Trim$(mRs("Name"))
    summ.add lDataObjectSummary
    Set lDataObjectSummary = Nothing
    mRs.MoveNext
Loop

Set mRs.ActiveConnection = Nothing
mConn.Close

Set search = summ

Exit Function

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "ExchangeFactory" & "::" & "search" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description


End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function loadDetails(ByVal where As String) As Exchange
Dim Exchange As Exchange
Dim SQL As String
Dim failpoint As Integer

On Error GoTo Err

failpoint = 100

Set mConn = New ADODB.Connection
mConn.Open mDB.connectionString

failpoint = 200

Set mRs = New recordset
mRs.CursorLocation = adUseClient
SQL = "SELECT * " & _
        "FROM exchange " & _
        "WHERE " & where

mRs.Open SQL, mConn, adOpenStatic, adLockBatchOptimistic

failpoint = 300

If Not mRs.EOF Then
    Set Exchange = New Exchange
    
    Exchange.recordset = mRs
    Exchange.initialise mDB
    Set loadDetails = Exchange
Else
    Set loadDetails = Nothing
End If

mConn.Close
Set mRs.ActiveConnection = Nothing
Set mRs = Nothing
Set Exchange = Nothing

Exit Function

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "ExchangeFactory" & "::" & "loadDetails" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description
End Function













