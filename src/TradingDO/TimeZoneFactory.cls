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

Private Sub DataObjectFactory_delete(ByVal Id As Long)
Const ProcName As String = "DataObjectFactory_delete"

On Error GoTo Err

mBaseFactory.Delete Id

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get DataObjectFactory_FieldSpecifiers() As FieldSpecifiers
Const ProcName As String = "DataObjectFactory_FieldSpecifiers"

On Error GoTo Err

Set DataObjectFactory_FieldSpecifiers = FieldSpecifiers

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function DataObjectFactory_LoadByID( _
                ByVal Id As Long) As BusinessDataObject
Const ProcName As String = "DataObjectFactory_LoadByID"

On Error GoTo Err

Set DataObjectFactory_LoadByID = LoadByID(Id)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_LoadByIDAsync( _
                ByVal Id As Long, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "DataObjectFactory_LoadByIDAsync"

On Error GoTo Err

Set DataObjectFactory_LoadByIDAsync = LoadByIDAsync(Id, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_LoadByName( _
                ByVal pName As String) As BusinessDataObject
Const ProcName As String = "DataObjectFactory_LoadByName"

On Error GoTo Err

Set DataObjectFactory_LoadByName = LoadByName(Trim$(pName))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_LoadByNameAsync( _
                ByVal pName As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "DataObjectFactory_LoadByNameAsync"

On Error GoTo Err

Set DataObjectFactory_LoadByNameAsync = LoadByNameAsync(pName, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_LoadByQuery( _
                ByVal whereCondition As String) As BusinessDataObject
Const ProcName As String = "DataObjectFactory_LoadByQuery"

On Error GoTo Err

Set DataObjectFactory_LoadByQuery = LoadByQuery(whereCondition)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_LoadByQueryAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "DataObjectFactory_LoadByQueryAsync"

On Error GoTo Err

Set DataObjectFactory_LoadByQueryAsync = LoadByQueryAsync(whereCondition, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_MakeNew() As BusinessDataObject
Const ProcName As String = "DataObjectFactory_MakeNew"

On Error GoTo Err

Set DataObjectFactory_MakeNew = MakeNew

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_MakeNewAsync( _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "DataObjectFactory_MakeNewAsync"

On Error GoTo Err

Set DataObjectFactory_MakeNewAsync = MakeNewAsync(cookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_makeNewFromRecordset( _
                ByVal rs As ADODB.Recordset) As BusinessDataObject
Dim obj As New TimeZone
Const ProcName As String = "DataObjectFactory_makeNewFromRecordset"

On Error GoTo Err

Set DataObjectFactory_makeNewFromRecordset = obj
obj.Initialise mDB, rs

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_Query( _
                ByVal whereCondition As String, _
                fieldNames() As String) As DataObjectSummaries
Const ProcName As String = "DataObjectFactory_Query"

On Error GoTo Err

Set DataObjectFactory_Query = Query(whereCondition, fieldNames)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_QueryAsync( _
                ByVal whereCondition As String, _
                fieldNames() As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "DataObjectFactory_QueryAsync"

On Error GoTo Err

Set DataObjectFactory_QueryAsync = QueryAsync(whereCondition, fieldNames, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_QueryObjects( _
                ByVal whereCondition As String) As DataObjects
Const ProcName As String = "DataObjectFactory_QueryObjects"

On Error GoTo Err

Set DataObjectFactory_QueryObjects = QueryObjects(whereCondition)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_QueryObjectsAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "DataObjectFactory_QueryObjectsAsync"

On Error GoTo Err

Set DataObjectFactory_QueryObjectsAsync = QueryObjectsAsync(whereCondition, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub DataObjectFactory_Refresh( _
                ByVal dataObj As BusinessDataObject)
Const ProcName As String = "DataObjectFactory_Refresh"

On Error GoTo Err

Refresh dataObj

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function DataObjectFactory_Search( _
                ByVal criterion As String, _
                fieldNames() As String) As DataObjectSummaries
Const ProcName As String = "DataObjectFactory_Search"

On Error GoTo Err

Set DataObjectFactory_Search = Search(criterion, fieldNames)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_SearchAsync( _
                ByVal criterion As String, _
                fieldNames() As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "DataObjectFactory_SearchAsync"

On Error GoTo Err

Set DataObjectFactory_SearchAsync = SearchAsync(criterion, fieldNames, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_SearchObjects( _
                ByVal criterion As String) As DataObjects
Const ProcName As String = "DataObjectFactory_SearchObjects"

On Error GoTo Err

Set DataObjectFactory_SearchObjects = SearchObjects(criterion)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_SearchObjectsAsync( _
                ByVal criterion As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "DataObjectFactory_SearchObjectsAsync"

On Error GoTo Err

Set DataObjectFactory_SearchObjectsAsync = SearchObjectsAsync(criterion, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get FieldSpecifiers() As FieldSpecifiers
Const ProcName As String = "FieldSpecifiers"

On Error GoTo Err

Set FieldSpecifiers = mBaseFactory.FieldSpecifiers

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal db As TradingDB)

Const ProcName As String = "Initialise"

On Error GoTo Err

Set mDB = db
Set mBaseFactory = CreateAbstractDOFactory(Me, mDB.ConnectionString, "timezone", "vTimezones")

mBaseFactory.AddField FieldNameName, TimeZoneColumnName, FieldWidthTimeZone, FieldAlignTimeZone, True
mBaseFactory.AddField FieldNameCanonicalName, TimeZoneColumnCanonicalName, FieldWidthTimeZone, FieldAlignTimeZone, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Public Function LoadByID( _
                ByVal Id As Long) As TimeZone
Dim rs As Recordset
Const ProcName As String = "LoadByID"

On Error GoTo Err

Set rs = mBaseFactory.LoadByID(Id)
If Not rs Is Nothing Then
    Set LoadByID = New TimeZone
    LoadByID.Initialise mDB, rs
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadByIDAsync( _
                ByVal Id As Long, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "LoadByIDAsync"

On Error GoTo Err

Set LoadByIDAsync = mBaseFactory.LoadByIDAsync(Id, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadByName( _
                ByVal pName As String) As TimeZone
Dim rs As Recordset

Const ProcName As String = "LoadByName"

On Error GoTo Err

Set rs = mBaseFactory.LoadByName(Trim$(pName))
If Not rs Is Nothing Then
    Set LoadByName = New TimeZone
    LoadByName.Initialise mDB, rs
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadByNameAsync( _
                ByVal pName As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "LoadByNameAsync"

On Error GoTo Err

Set LoadByNameAsync = mBaseFactory.LoadByNameAsync(Trim$(pName), cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadByQuery( _
                ByVal whereCondition As String) As TimeZone
Dim rs As Recordset
Const ProcName As String = "LoadByQuery"

On Error GoTo Err

Set rs = mBaseFactory.LoadByQuery(whereCondition)
If Not rs Is Nothing Then
    Set LoadByQuery = New TimeZone
    LoadByQuery.Initialise mDB, rs
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadByQueryAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "LoadByQueryAsync"

On Error GoTo Err

Set LoadByQueryAsync = mBaseFactory.LoadByQueryAsync(whereCondition, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MakeNew() As TimeZone
Const ProcName As String = "MakeNew"

On Error GoTo Err

Set MakeNew = New TimeZone
MakeNew.Initialise mDB, mBaseFactory.MakeNew

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MakeNewAsync( _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "MakeNewAsync"

On Error GoTo Err

Set MakeNewAsync = mBaseFactory.MakeNewAsync(cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Query( _
                ByVal whereCondition As String, _
                ByRef fieldNames() As String) As DataObjectSummaries
Const ProcName As String = "Query"

On Error GoTo Err

Set Query = mBaseFactory.Query(whereCondition, fieldNames)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function QueryAsync( _
                ByVal whereCondition As String, _
                ByRef fieldNames() As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "QueryAsync"

On Error GoTo Err

Set QueryAsync = mBaseFactory.QueryAsync(whereCondition, fieldNames, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function QueryObjects( _
                ByVal whereCondition As String) As DataObjects
Const ProcName As String = "QueryObjects"

On Error GoTo Err

Set QueryObjects = mBaseFactory.QueryObjects(whereCondition)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function QueryObjectsAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "QueryObjectsAsync"

On Error GoTo Err

Set QueryObjectsAsync = mBaseFactory.QueryObjectsAsync(whereCondition, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Refresh( _
                ByVal pTimezone As TimeZone)
Dim rs As Recordset

Const ProcName As String = "Refresh"

On Error GoTo Err

Set rs = mBaseFactory.LoadByID(pTimezone.Id)
If Not rs Is Nothing Then
    pTimezone.Initialise mDB, rs
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Search( _
                ByVal criterion As String, _
                ByRef fieldNames() As String) As DataObjectSummaries
Const ProcName As String = "Search"

On Error GoTo Err

Set Search = mBaseFactory.Search(criterion, fieldNames)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SearchAsync( _
                ByVal criterion As String, _
                ByRef fieldNames() As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "SearchAsync"

On Error GoTo Err

Set SearchAsync = mBaseFactory.SearchAsync(criterion, fieldNames, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SearchObjects( _
                ByVal criterion As String) As DataObjects
Const ProcName As String = "SearchObjects"

On Error GoTo Err

Set SearchObjects = mBaseFactory.SearchObjects(criterion)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SearchObjectsAsync( _
                ByVal criterion As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "SearchObjectsAsync"

On Error GoTo Err

Set SearchObjectsAsync = mBaseFactory.SearchObjectsAsync(criterion, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================















