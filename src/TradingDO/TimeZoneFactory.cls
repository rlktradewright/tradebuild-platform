VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeZoneFactory"
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
Private Const ModuleName As String = "TimeZoneFactory"

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

Public Function loadByName(ByVal pName As String) As TimeZone
Dim where As String
where = "Name='" & Replace(pName, "'", "''") & "'"
Set loadByName = loadDetails(where)
End Function

Public Function loadByID(ByVal id As Long) As TimeZone
Dim where As String
where = "ID=" & id
Set loadByID = loadDetails(where)
End Function

Public Function makeNew() As TimeZone
Dim SQL As String
Dim lTimeZone As TimeZone

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

Set mConn = New ADODB.Connection
mConn.Open mDB.connectionString

failpoint = 200

Set mRs = New recordset
mRs.CursorLocation = adUseClient
SQL = "SELECT * " & _
        "FROM timezone " & _
        "WHERE 1=0"


mRs.Open SQL, mConn, adOpenStatic, adLockBatchOptimistic

Set lTimeZone = New TimeZone
lTimeZone.initialise mDB
lTimeZone.recordset = mRs
Set makeNew = lTimeZone

Set mRs.ActiveConnection = Nothing
Set mRs = Nothing
mConn.Close

Exit Function

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "TimeZoneFactory" & "::" & "makeNew" & "." & failpoint & _
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
        "FROM TimeZone " & _
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

mConn.Close
Set mRs.ActiveConnection = Nothing

Set search = summ

Exit Function

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "TimeZoneFactory" & "::" & "search" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description


End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function loadDetails(ByVal where As String) As TimeZone
Dim TimeZone As TimeZone
Dim SQL As String
Dim failpoint As Integer

On Error GoTo Err

failpoint = 100

Set mConn = New ADODB.Connection
mConn.Open mDB.connectionString

failpoint = 200

Set mRs = New recordset
mRs.CursorLocation = adUseClient
SQL = "SELECT TZ1.id, TZ1.name, TZ1.canonicalId, TZ2.Name AS CanonicalName " & _
        "FROM TimeZone AS TZ1 LEFT JOIN TimeZone AS TZ2 ON " & _
        "       TZ1.CanonicalId = TZ2.Id " & _
        "WHERE " & where

mRs.Open SQL, mConn, adOpenStatic, adLockBatchOptimistic

failpoint = 300

If Not mRs.EOF Then
    Set TimeZone = New TimeZone
    
    TimeZone.recordset = mRs
    TimeZone.initialise mDB
    Set loadDetails = TimeZone
Else
    Set loadDetails = Nothing
End If

mConn.Close
Set mRs.ActiveConnection = Nothing
Set mRs = Nothing
Set TimeZone = Nothing

Exit Function

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "TimeZoneFactory" & "::" & "loadDetails" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description
End Function















