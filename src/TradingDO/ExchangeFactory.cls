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

Private Property Get DataObjectFactory_fieldSpecifiers() As BusObjUtils10.fieldSpecifiers
Set DataObjectFactory_fieldSpecifiers = fieldSpecifiers
End Property

Private Function DataObjectFactory_loadByID( _
                ByVal id As Long) As BusinessDataObject
Set DataObjectFactory_loadByID = loadByID(id)
End Function

Private Function DataObjectFactory_loadByIDAsync( _
                ByVal id As Long, _
                Optional ByVal cookie As Variant) As TWUtilities30.TaskController
Set DataObjectFactory_loadByIDAsync = loadByIDAsync(id, cookie)
End Function

Private Function DataObjectFactory_loadByName( _
                ByVal pName As String) As BusinessDataObject
Set DataObjectFactory_loadByName = loadByName(Trim$(pName))
End Function

Private Function DataObjectFactory_loadByNameAsync( _
                ByVal pName As String, _
                Optional ByVal cookie As Variant) As TWUtilities30.TaskController
Set DataObjectFactory_loadByNameAsync = loadByNameAsync(pName, cookie)
End Function

Private Function DataObjectFactory_loadByQuery( _
                ByVal whereCondition As String) As BusinessDataObject
Set DataObjectFactory_loadByQuery = loadByQuery(whereCondition)
End Function

Private Function DataObjectFactory_loadByQueryAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant) As TWUtilities30.TaskController
Set DataObjectFactory_loadByQueryAsync = loadByQueryAsync(whereCondition, cookie)
End Function

Private Function DataObjectFactory_makeNew() As BusinessDataObject
Set DataObjectFactory_makeNew = makeNew
End Function

Private Function DataObjectFactory_makeNewAsync( _
                Optional ByVal cookie As Variant) As TWUtilities30.TaskController
Set DataObjectFactory_makeNewAsync = makeNewAsync(cookie)
End Function

Private Function DataObjectFactory_makeNewFromRecordset( _
                ByVal rs As ADODB.Recordset) As BusObjUtils10.BusinessDataObject
Dim obj As New Exchange
Set DataObjectFactory_makeNewFromRecordset = obj
obj.initialise mDB, rs
End Function

Private Function DataObjectFactory_query( _
                ByVal whereCondition As String, _
                fieldNames() As String) As DataObjectSummaries
Set DataObjectFactory_query = query(whereCondition, fieldNames)
End Function

Private Function DataObjectFactory_queryAsync( _
                ByVal whereCondition As String, _
                fieldNames() As String, _
                Optional ByVal cookie As Variant) As TWUtilities30.TaskController
Set DataObjectFactory_queryAsync = queryAsync(whereCondition, fieldNames, cookie)
End Function

Private Function DataObjectFactory_queryObjects( _
                ByVal whereCondition As String) As BusObjUtils10.DataObjects
Set DataObjectFactory_queryObjects = queryObjects(whereCondition)
End Function

Private Function DataObjectFactory_queryObjectsAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant) As TWUtilities30.TaskController
Set DataObjectFactory_queryObjectsAsync = queryObjectsAsync(whereCondition, cookie)
End Function

Private Function DataObjectFactory_search( _
                ByVal criterion As String, _
                fieldNames() As String) As DataObjectSummaries
Set DataObjectFactory_search = search(criterion, fieldNames)
End Function

Private Function DataObjectFactory_searchAsync( _
                ByVal criterion As String, _
                fieldNames() As String, _
                Optional ByVal cookie As Variant) As TWUtilities30.TaskController
Set DataObjectFactory_searchAsync = searchAsync(criterion, fieldNames, cookie)
End Function

Private Function DataObjectFactory_searchObjects( _
                ByVal criterion As String) As BusObjUtils10.DataObjects
Set DataObjectFactory_searchObjects = searchObjects(criterion)
End Function

Private Function DataObjectFactory_searchObjectsAsync( _
                ByVal criterion As String, _
                Optional ByVal cookie As Variant) As TWUtilities30.TaskController
Set DataObjectFactory_searchObjectsAsync = searchObjectsAsync(criterion, cookie)
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get fieldSpecifiers() As fieldSpecifiers
Set fieldSpecifiers = mBaseFactory.fieldSpecifiers
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal db As TradingDB)

Set mDB = db
Set mBaseFactory = CreateAbstractDOFactory(Me, mDB.connectionString, "Exchange", "vExchanges")

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

Public Function loadByIDAsync( _
                ByVal id As Long, _
                Optional ByVal cookie As Variant) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set loadByIDAsync = mBaseFactory.loadByIDAsync(id, cookie)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadByIDAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function loadByName( _
                ByVal pName As String) As Exchange
Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset

pName = Trim$(pName)

Set rs = mBaseFactory.loadByName(Trim$(pName))
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

Public Function loadByNameAsync( _
                ByVal pName As String, _
                Optional ByVal cookie As Variant) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set loadByNameAsync = mBaseFactory.loadByNameAsync(Trim$(pName), cookie)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadByNameAsync" & "." & failpoint & _
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

Public Function loadByQueryAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set loadByQueryAsync = mBaseFactory.loadByQueryAsync(whereCondition, cookie)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadByQueryAsync" & "." & failpoint & _
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

Public Function makeNewAsync( _
                Optional ByVal cookie As Variant) As TaskController

Dim failpoint As Long
On Error GoTo Err

Set makeNewAsync = mBaseFactory.makeNewAsync(cookie)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "makeNewAsync" & "." & failpoint & _
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

Public Function queryAsync( _
                ByVal whereCondition As String, _
                ByRef fieldNames() As String, _
                Optional ByVal cookie As Variant) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set queryAsync = mBaseFactory.queryAsync(whereCondition, fieldNames, cookie)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "queryAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function queryObjects( _
                ByVal whereCondition As String) As DataObjects
Dim failpoint As Long
On Error GoTo Err

Set queryObjects = mBaseFactory.queryObjects(whereCondition)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "queryObjects" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function queryObjectsAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set queryObjectsAsync = mBaseFactory.queryObjectsAsync(whereCondition, cookie)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "queryObjectsAsync" & "." & failpoint & _
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

Public Function searchAsync( _
                ByVal criterion As String, _
                ByRef fieldNames() As String, _
                Optional ByVal cookie As Variant) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set searchAsync = mBaseFactory.searchAsync(criterion, fieldNames, cookie)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "searchAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function searchObjects( _
                ByVal criterion As String) As DataObjects
Dim failpoint As Long
On Error GoTo Err

Set searchObjects = mBaseFactory.searchObjects(criterion)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "searchObjects" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function searchObjectsAsync( _
                ByVal criterion As String, _
                Optional ByVal cookie As Variant) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set searchObjectsAsync = mBaseFactory.searchObjectsAsync(criterion, cookie)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "searchObjectsAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

'@================================================================================
' Helper Functions
'@================================================================================








