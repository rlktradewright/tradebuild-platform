VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InstrumentLocalSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private mConnectionString As String
Private mConn As ADODB.Connection
Private mRs As ADODB.recordset
Private mCol As Collection
Private mInstrumentID As Long

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
Set mCol = New Collection
End Sub

'@===============================================================================
' XXXX Interface Members
'@===============================================================================

'@===============================================================================
' XXXX Event Handlers
'@===============================================================================

'@===============================================================================
' Properties
'@===============================================================================

Friend Property Let connectionString( _
                ByVal value As String)
mConnectionString = value
End Property

'@===============================================================================
' Methods
'@===============================================================================

Public Function add() As InstrumentLocalSymbol
mRs.AddNew
mRs("InstrumentID") = mInstrumentID
Set add = New InstrumentLocalSymbol
add.initialise mRs.Clone, mRs.bookmark
mCol.add add
End Function

Public Sub cancel()
mRs.CancelBatch
End Sub

Public Function count() As Long
count = mCol.count
End Function

Public Function item( _
                ByVal index As Variant) As InstrumentLocalSymbol
Set item = mCol(index)
End Function

Friend Sub load( _
                ByVal instrumentID As Long)
Dim obj As InstrumentLocalSymbol
Dim SQL As String

Dim failpoint As Long
On Error GoTo Err

If Not mRs Is Nothing Then Err.Raise 445

failpoint = 100

Set mConn = New ADODB.Connection
mConn.Open mConnectionString

failpoint = 200

Set mRs = New recordset
mRs.CursorLocation = adUseClient
SQL = "SELECT ID, ProviderKey, InstrumentID, LocalSymbol " & _
        "FROM InstrumentLocalSymbol " & _
        "WHERE InstrumentID=" & instrumentID & " " & _
        "ORDER BY ProviderKey"

mRs.Open SQL, mConn, adOpenKeyset, adLockBatchOptimistic

Do While Not mRs.EOF
    Set obj = New InstrumentLocalSymbol
    obj.initialise mRs.Clone, mRs.bookmark
    mCol.add obj
    mRs.MoveNext
Loop

mConn.Close
Set mRs.ActiveConnection = Nothing

mInstrumentID = instrumentID

Exit Sub

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "InstrumentLocalSymbols" & "::" & "load" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mCol.[_NewEnum]
End Function

Public Sub remove(index As Long)
Dim obj As InstrumentLocalSymbol
Set obj = mCol(index)
mRs.bookmark = obj.bookmark
mRs.Delete adAffectCurrent
mCol.remove index
End Sub

Public Sub removeAll()
mRs.MoveFirst
Do While Not mRs.EOF
    mRs.Delete
    mRs.MoveNext
Loop
'mRs.Delete adAffectAll
Set mCol = New Collection
End Sub

Public Sub save()
mRs.UpdateBatch
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================



