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

Implements DataObjectFactory

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

Private Const ProjectName                       As String = "TradingDO26"
Private Const ModuleName                        As String = "ExchangeFactory"

'@================================================================================
' Member variables
'@================================================================================

Private mDB As TradingDB

Private mBaseFactory As AbstractDOFactory

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' DataObjectFactory Interface Members
'@================================================================================

Private Property Get DataObjectFactory_fieldNames() As String()
DataObjectFactory_fieldNames = fieldNames
End Property

Private Property Get DataObjectFactory_fieldSpecifiers() As BusObjUtils10.FieldSpecifier()
DataObjectFactory_fieldSpecifiers = fieldSpecifiers
End Property

Private Function DataObjectFactory_loadByID( _
                ByVal id As Long) As BusinessDataObject
Set DataObjectFactory_loadByID = loadByID(id)
End Function

Private Function DataObjectFactory_loadByName( _
                ByVal pName As String) As BusinessDataObject
Set DataObjectFactory_loadByName = loadByName(pName)
End Function

Private Function DataObjectFactory_loadByQuery( _
                ByVal whereCondition As String) As BusinessDataObject
Set DataObjectFactory_loadByQuery = loadByQuery(whereCondition)
End Function

Private Function DataObjectFactory_makeNew() As BusinessDataObject
Set DataObjectFactory_makeNew = makeNew
End Function

Private Function DataObjectFactory_query( _
                ByVal whereCondition As String, _
                fieldNames() As String) As DataObjectSummaries
Set DataObjectFactory_query = query(whereCondition, fieldNames)
End Function

Private Function DataObjectFactory_search( _
                ByVal criterion As String, _
                fieldNames() As String) As DataObjectSummaries
Set DataObjectFactory_search = search(criterion, fieldNames)
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get fieldNames() As String()
fieldNames = mBaseFactory.fieldNames
End Property

Public Property Get fieldSpecifiers() As FieldSpecifier()
fieldSpecifiers = mBaseFactory.fieldSpecifiers
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal db As TradingDB)

Set mDB = db
Set mBaseFactory = CreateAbstractDOFactory(mDB.connectionString, "Exchange", "vExchanges")

mBaseFactory.addField FieldNameName, ExchangeColumnName, FieldWidthExchange, FieldAlignExchange, True
mBaseFactory.addField FieldNameTimeZone, ExchangeColumnTimeZone, FieldWidthTimeZone, FieldAlignTimeZone, True
mBaseFactory.addField FieldNameNotes, ExchangeColumnNotes, FieldWidthNotes, FieldAlignNotes, True

End Sub

Public Function loadByID( _
                ByVal id As Long) As Exchange
Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset
Set rs = mBaseFactory.loadByID(id)
If Not rs Is Nothing Then
    Set loadByID = New Exchange
    loadByID.initialise mDB, rs
End If

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadByID" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function loadByName( _
                ByVal pName As String) As Exchange
Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset

pName = Trim$(pName)

Set rs = mBaseFactory.loadByName(pName)
If Not rs Is Nothing Then
    Set loadByName = New Exchange
    loadByName.initialise mDB, rs
End If

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadByName" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function loadByQuery( _
                ByVal whereCondition As String) As Exchange
Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset
Set rs = mBaseFactory.loadByQuery(whereCondition)
If Not rs Is Nothing Then
    Set loadByQuery = New Exchange
    loadByQuery.initialise mDB, rs
End If

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadByQuery" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function makeNew() As Exchange
Dim failpoint As Long
On Error GoTo Err

Set makeNew = New Exchange
makeNew.initialise mDB, mBaseFactory.makeNew

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "makeNew" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function query( _
                ByVal whereCondition As String, _
                ByRef fieldNames() As String) As DataObjectSummaries
Dim failpoint As Long
On Error GoTo Err

Set query = mBaseFactory.query(whereCondition, fieldNames)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "query" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function search( _
                ByVal criterion As String, _
                ByRef fieldNames() As String) As DataObjectSummaries
Dim failpoint As Long
On Error GoTo Err

Set search = mBaseFactory.search(criterion, fieldNames)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "search" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

'@================================================================================
' Helper Functions
'@================================================================================








