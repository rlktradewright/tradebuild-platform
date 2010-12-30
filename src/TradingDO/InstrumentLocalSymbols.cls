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

Private Const ModuleName                As String = "InstrumentLocalSymbols"

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
Private mConn As ADODB.connection
Private mRs As ADODB.Recordset
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

Friend Property Let ConnectionString( _
                ByVal value As String)
value = Trim$(value)
mConnectionString = value
End Property

'@===============================================================================
' Methods
'@===============================================================================

Public Function Add() As InstrumentLocalSymbol
Const ProcName As String = "Add"
Dim failpoint As String
On Error GoTo Err

mRs.AddNew
mRs("InstrumentID") = mInstrumentID
Set Add = New InstrumentLocalSymbol
Add.Initialise mRs.Clone, mRs.Bookmark
mCol.Add Add

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub Cancel()
Const ProcName As String = "Cancel"
Dim failpoint As String
On Error GoTo Err

mRs.CancelBatch

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function Count() As Long
Const ProcName As String = "count"
Dim failpoint As String
On Error GoTo Err

Count = mCol.Count

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function Item( _
                ByVal index As Variant) As InstrumentLocalSymbol
Const ProcName As String = "Item"
Dim failpoint As String
On Error GoTo Err

Set Item = mCol(index)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub Load( _
                ByVal instrumentID As Long)
Dim obj As InstrumentLocalSymbol
Dim SQL As String

Const ProcName As String = "Load"
Dim failpoint As String
On Error GoTo Err

If Not mRs Is Nothing Then Err.Raise 445

failpoint = 100

Set mConn = New ADODB.connection
mConn.Open mConnectionString

failpoint = 200

Set mRs = New Recordset
mRs.cursorlocation = adUseClient
SQL = "SELECT ID, ProviderKey, InstrumentID, LocalSymbol " & _
        "FROM InstrumentLocalSymbol " & _
        "WHERE InstrumentID=" & instrumentID & " " & _
        "ORDER BY ProviderKey"

mRs.Open SQL, mConn, adOpenKeyset, adLockBatchOptimistic

Set mRs.ActiveConnection = Nothing

Do While Not mRs.EOF
    Set obj = New InstrumentLocalSymbol
    obj.Initialise mRs.Clone, mRs.Bookmark
    mCol.Add obj
    mRs.MoveNext
Loop

mConn.Close

mInstrumentID = instrumentID

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mCol.[_NewEnum]
End Function

Public Sub Remove(index As Long)
Dim obj As InstrumentLocalSymbol
Const ProcName As String = "Remove"
Dim failpoint As String
On Error GoTo Err

Set obj = mCol(index)
mRs.Bookmark = obj.Bookmark
mRs.Delete adAffectCurrent
mCol.Remove index

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveAll()
Const ProcName As String = "RemoveAll"
Dim failpoint As String
On Error GoTo Err

mRs.MoveFirst
Do While Not mRs.EOF
    mRs.Delete
    mRs.MoveNext
Loop
'mRs.Delete adAffectAll
Set mCol = New Collection

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Save()
Const ProcName As String = "Save"
Dim failpoint As String
On Error GoTo Err

mRs.UpdateBatch

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================



