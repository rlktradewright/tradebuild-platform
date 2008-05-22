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
Private Const ModuleName                        As String = "TimeZoneFactory"

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
                ByVal id As Long) As TWUtilities30.TaskController
Set DataObjectFactory_loadByIDAsync = loadByIDAsync(id)
End Function

Private Function DataObjectFactory_loadByName( _
                ByVal pName As String) As BusinessDataObject
Set DataObjectFactory_loadByName = loadByName(Trim$(pName))
End Function

Private Function DataObjectFactory_loadByNameAsync( _
                ByVal pName As String) As TWUtilities30.TaskController
Set DataObjectFactory_loadByNameAsync = loadByNameAsync(pName)
End Function

Private Function DataObjectFactory_loadByQuery( _
                ByVal whereCondition As String) As BusinessDataObject
Set DataObjectFactory_loadByQuery = loadByQuery(whereCondition)
End Function

Private Function DataObjectFactory_loadByQueryAsync( _
                ByVal whereCondition As String) As TWUtilities30.TaskController
Set DataObjectFactory_loadByQueryAsync = loadByQueryAsync(whereCondition)
End Function

Private Function DataObjectFactory_makeNew() As BusinessDataObject
Set DataObjectFactory_makeNew = makeNew
End Function

Private Function DataObjectFactory_makeNewAsync() As TWUtilities30.TaskController
Set DataObjectFactory_makeNewAsync = makeNewAsync
End Function

Private Function DataObjectFactory_query( _
                ByVal whereCondition As String, _
                fieldNames() As String) As DataObjectSummaries
Set DataObjectFactory_query = query(whereCondition, fieldNames)
End Function

Private Function DataObjectFactory_queryAsync( _
                ByVal whereCondition As String, _
                fieldNames() As String) As TWUtilities30.TaskController
Set DataObjectFactory_queryAsync = queryAsync(whereCondition, fieldNames)
End Function

Private Function DataObjectFactory_search( _
                ByVal criterion As String, _
                fieldNames() As String) As DataObjectSummaries
Set DataObjectFactory_search = search(criterion, fieldNames)
End Function

Private Function DataObjectFactory_searchAsync( _
                ByVal criterion As String, _
                fieldNames() As String) As TWUtilities30.TaskController
Set DataObjectFactory_searchAsync = searchAsync(criterion, fieldNames)
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
Set mBaseFactory = CreateAbstractDOFactory(Me, mDB.connectionString, "timezone", "vTimezones")

mBaseFactory.addField FieldNameName, TimeZoneColumnName, FieldWidthTimeZone, FieldAlignTimeZone, True
mBaseFactory.addField FieldNameCanonicalName, TimeZoneColumnCanonicalName, FieldWidthTimeZone, FieldAlignTimeZone, True

End Sub

Public Function loadByID( _
                ByVal id As Long) As TimeZone
Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset
Set rs = mBaseFactory.loadByID(id)
If Not rs Is Nothing Then
    Set loadByID = New TimeZone
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
                ByVal id As Long) As TaskController
Dim newDataObj As New TimeZone
Dim failpoint As Long
On Error GoTo Err

newDataObj.initialise mDB
Set loadByIDAsync = mBaseFactory.loadByIDAsync(newDataObj, id)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadByIDAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function loadByName( _
                ByVal pName As String) As TimeZone
Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset

Set rs = mBaseFactory.loadByName(Trim$(pName))
If Not rs Is Nothing Then
    Set loadByName = New TimeZone
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
                ByVal pName As String) As TaskController
Dim newDataObj As New TimeZone
Dim failpoint As Long
On Error GoTo Err

newDataObj.initialise mDB
Set loadByNameAsync = mBaseFactory.loadByNameAsync(newDataObj, Trim$(pName))

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadByNameAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function loadByQuery( _
                ByVal whereCondition As String) As TimeZone
Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset
Set rs = mBaseFactory.loadByQuery(whereCondition)
If Not rs Is Nothing Then
    Set loadByQuery = New TimeZone
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
                ByVal whereCondition As String) As TaskController
Dim newDataObj As New TimeZone
Dim failpoint As Long
On Error GoTo Err

newDataObj.initialise mDB
Set loadByQueryAsync = mBaseFactory.loadByQueryAsync(newDataObj, whereCondition)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadByQueryAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function makeNew() As TimeZone
Dim failpoint As Long
On Error GoTo Err

Set makeNew = New TimeZone
makeNew.initialise mDB, mBaseFactory.makeNew

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "makeNew" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function makeNewAsync() As TaskController
Dim newDataObj As New instrument

Dim failpoint As Long
On Error GoTo Err

newDataObj.initialise mDB
Set makeNewAsync = mBaseFactory.makeNewAsync(newDataObj)

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
                ByRef fieldNames() As String) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set queryAsync = mBaseFactory.queryAsync(whereCondition, fieldNames)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "queryAsync" & "." & failpoint & _
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
                ByRef fieldNames() As String) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set searchAsync = mBaseFactory.searchAsync(criterion, fieldNames)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "searchAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

'@================================================================================
' Helper Functions
'@================================================================================















