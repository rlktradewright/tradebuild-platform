VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InstrumentClassFactory"
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
Private Const ModuleName As String = "InstrumentClassFactory"

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

Public Function loadByName( _
                ByVal pName As String) As InstrumentClass
Dim where As String
where = "Name='" & Replace(pName, "'", "''") & "'"
Set loadByName = loadDetails(where)
End Function

Public Function loadByID( _
                ByVal id As Long) As InstrumentClass
Dim where As String
where = "ID=" & id
Set loadByID = loadDetails(where)
End Function

Public Function loadSummaries( _
                Optional ByVal pExchange As String, _
                Optional ByVal pCategory As String, _
                Optional ByVal pCurrency As String) As DataObjectSummaries
Dim where As String
If pExchange <> "" Then where = IIf(where = "", "", where & " AND ") & "Exchange='" & pExchange & "'"
If pCategory <> "" Then where = IIf(where = "", "", where & " AND ") & "Category='" & pCategory & "'"
If pCurrency <> "" Then where = IIf(where = "", "", where & " AND ") & "Currency='" & pCurrency & "'"
Set loadSummaries = searchWhere(where)
End Function

Public Function makeNew() As InstrumentClass
Dim SQL As String
Dim lInstrumentClass As InstrumentClass

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

Set mConn = New ADODB.Connection
mConn.Open mDB.connectionString

failpoint = 200

Set mRs = New recordset
mRs.CursorLocation = adUseClient
SQL = "SELECT * " & _
        "FROM InstrumentClass " & _
        "WHERE 1=0"


mRs.Open SQL, mConn, adOpenStatic, adLockBatchOptimistic

Set lInstrumentClass = New InstrumentClass
lInstrumentClass.initialise mDB
lInstrumentClass.recordset = mRs
Set makeNew = lInstrumentClass

Set mRs.ActiveConnection = Nothing
Set mRs = Nothing
mConn.Close

Exit Function

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "InstrumentClassFactory" & "::" & "makeNew" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function loadDetails(ByVal where As String) As InstrumentClass
Dim lInstrumentClass As InstrumentClass
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
        "FROM InstrumentClass " & _
        "WHERE " & where

mRs.Open SQL, mConn, adOpenStatic, adLockBatchOptimistic

failpoint = 300

If Not mRs.EOF Then
    Set lInstrumentClass = New InstrumentClass
    
    lInstrumentClass.recordset = mRs
    lInstrumentClass.initialise mDB
    Set loadDetails = lInstrumentClass
Else
    Set loadDetails = Nothing
End If

Set mRs.ActiveConnection = Nothing
Set mRs = Nothing
mConn.Close
Set lInstrumentClass = Nothing

Exit Function

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "InstrumentClassFactory" & "::" & "loadDetails" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description
End Function

Private Function searchWhere(ByVal where As String) As DataObjectSummaries
Dim summ As DataObjectSummaries
Dim SQL As String
Dim lDataObjectSummary As DataObjectSummary


Dim failpoint As Long
On Error GoTo Err

failpoint = 100

Set mConn = New ADODB.Connection
mConn.Open mDB.connectionString

failpoint = 200

Set mRs = New recordset
mRs.CursorLocation = adUseClient
SQL = "SELECT DISTINCT InstrumentClassID, InstrumentClassName " & _
        "FROM vInstrumentDetails " & _
        IIf(where <> "", "WHERE " & where, "") & _
        " ORDER BY InstrumentClassName"


mRs.Open SQL, mConn, adOpenStatic, adLockBatchOptimistic

Set summ = New DataObjectSummaries

Do While Not mRs.EOF
    Set lDataObjectSummary = New DataObjectSummary
    lDataObjectSummary.id = mRs("InstrumentClassID")
    lDataObjectSummary.name = Trim$(mRs("InstrumentClassName"))
    summ.add lDataObjectSummary
    Set lDataObjectSummary = Nothing
    mRs.MoveNext
Loop

Set mRs.ActiveConnection = Nothing
mConn.Close

Set searchWhere = summ

Exit Function

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "InstrumentClassFactory" & "::" & "searchWhere" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description


End Function















