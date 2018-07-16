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
Private mConn As ADODB.Connection
Private mRS As ADODB.Recordset
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
                ByVal Value As String)
Value = Trim$(Value)
mConnectionString = Value
End Property

'@===============================================================================
' Methods
'@===============================================================================

Public Function Add() As InstrumentLocalSymbol
Const ProcName As String = "Add"

On Error GoTo Err

mRS.AddNew
mRS("InstrumentID") = mInstrumentID
Set Add = New InstrumentLocalSymbol
Add.Initialise mRS.Clone, mRS.Bookmark
mCol.Add Add

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Cancel()
Const ProcName As String = "Cancel"

On Error GoTo Err

mRS.CancelBatch

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Count() As Long
Const ProcName As String = "count"

On Error GoTo Err

Count = mCol.Count

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Item( _
                ByVal index As Variant) As InstrumentLocalSymbol
Const ProcName As String = "Item"

On Error GoTo Err

Set Item = mCol(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Load( _
                ByVal instrumentID As Long)
Const ProcName As String = "Load"
Dim failpoint As String
On Error GoTo Err

Dim obj As InstrumentLocalSymbol
Dim SQL As String

If Not mRS Is Nothing Then Err.Raise 445

failpoint = 100

Set mConn = New ADODB.Connection
mConn.Open mConnectionString

failpoint = 200

Set mRS = New Recordset
mRS.cursorlocation = adUseClient
SQL = "SELECT ID, ProviderKey, InstrumentID, LocalSymbol " & _
        "FROM InstrumentLocalSymbol " & _
        "WHERE InstrumentID=" & instrumentID & " " & _
        "ORDER BY ProviderKey"

mRS.Open SQL, mConn, adOpenKeyset, adLockBatchOptimistic

Set mRS.ActiveConnection = Nothing

Do While Not mRS.EOF
    Set obj = New InstrumentLocalSymbol
    obj.Initialise mRS.Clone, mRS.Bookmark
    mCol.Add obj
    mRS.MoveNext
Loop

mConn.Close

mInstrumentID = instrumentID

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mCol.[_NewEnum]
End Function

Public Sub Remove(index As Long)
Dim obj As InstrumentLocalSymbol
Const ProcName As String = "Remove"

On Error GoTo Err

Set obj = mCol(index)
mRS.Bookmark = obj.Bookmark
mRS.Delete adAffectCurrent
mCol.Remove index

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveAll()
Const ProcName As String = "RemoveAll"

On Error GoTo Err

mRS.MoveFirst
Do While Not mRS.EOF
    mRS.Delete
    mRS.MoveNext
Loop
'mRs.Delete adAffectAll
Set mCol = New Collection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Save()
Const ProcName As String = "Save"

On Error GoTo Err

mRS.UpdateBatch

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================



