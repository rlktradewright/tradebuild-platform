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
Dim failpoint As String
On Error GoTo Err

mBaseFactory.Delete Id

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get DataObjectFactory_FieldSpecifiers() As BusObjUtils10.FieldSpecifiers
Const ProcName As String = "DataObjectFactory_FieldSpecifiers"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_FieldSpecifiers = FieldSpecifiers

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Function DataObjectFactory_LoadByID( _
                ByVal Id As Long) As BusinessDataObject
Const ProcName As String = "DataObjectFactory_LoadByID"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_LoadByID = LoadByID(Id)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_LoadByIDAsync( _
                ByVal Id As Long, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Const ProcName As String = "DataObjectFactory_LoadByIDAsync"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_LoadByIDAsync = LoadByIDAsync(Id, cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_LoadByName( _
                ByVal pName As String) As BusinessDataObject
Const ProcName As String = "DataObjectFactory_LoadByName"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_LoadByName = LoadByName(Trim$(pName))

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_LoadByNameAsync( _
                ByVal pName As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Const ProcName As String = "DataObjectFactory_LoadByNameAsync"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_LoadByNameAsync = LoadByNameAsync(pName, cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_LoadByQuery( _
                ByVal whereCondition As String) As BusinessDataObject
Const ProcName As String = "DataObjectFactory_LoadByQuery"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_LoadByQuery = LoadByQuery(whereCondition)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_LoadByQueryAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Const ProcName As String = "DataObjectFactory_LoadByQueryAsync"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_LoadByQueryAsync = LoadByQueryAsync(whereCondition, cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_MakeNew() As BusinessDataObject
Const ProcName As String = "DataObjectFactory_MakeNew"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_MakeNew = MakeNew

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_MakeNewAsync( _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Const ProcName As String = "DataObjectFactory_MakeNewAsync"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_MakeNewAsync = MakeNewAsync(cookie)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_makeNewFromRecordset( _
                ByVal rs As ADODB.Recordset) As BusObjUtils10.BusinessDataObject
Dim obj As New TimeZone
Const ProcName As String = "DataObjectFactory_makeNewFromRecordset"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_makeNewFromRecordset = obj
obj.Initialise mDB, rs

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_Query( _
                ByVal whereCondition As String, _
                fieldNames() As String) As DataObjectSummaries
Const ProcName As String = "DataObjectFactory_Query"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_Query = Query(whereCondition, fieldNames)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_QueryAsync( _
                ByVal whereCondition As String, _
                fieldNames() As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Const ProcName As String = "DataObjectFactory_QueryAsync"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_QueryAsync = QueryAsync(whereCondition, fieldNames, cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_QueryObjects( _
                ByVal whereCondition As String) As BusObjUtils10.DataObjects
Const ProcName As String = "DataObjectFactory_QueryObjects"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_QueryObjects = QueryObjects(whereCondition)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_QueryObjectsAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Const ProcName As String = "DataObjectFactory_QueryObjectsAsync"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_QueryObjectsAsync = QueryObjectsAsync(whereCondition, cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub DataObjectFactory_Refresh( _
                ByVal dataObj As BusObjUtils10.BusinessDataObject)
Const ProcName As String = "DataObjectFactory_Refresh"
Dim failpoint As String
On Error GoTo Err

Refresh dataObj

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function DataObjectFactory_Search( _
                ByVal criterion As String, _
                fieldNames() As String) As DataObjectSummaries
Const ProcName As String = "DataObjectFactory_Search"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_Search = Search(criterion, fieldNames)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_SearchAsync( _
                ByVal criterion As String, _
                fieldNames() As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Const ProcName As String = "DataObjectFactory_SearchAsync"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_SearchAsync = SearchAsync(criterion, fieldNames, cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_SearchObjects( _
                ByVal criterion As String) As BusObjUtils10.DataObjects
Const ProcName As String = "DataObjectFactory_SearchObjects"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_SearchObjects = SearchObjects(criterion)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function DataObjectFactory_SearchObjectsAsync( _
                ByVal criterion As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Const ProcName As String = "DataObjectFactory_SearchObjectsAsync"
Dim failpoint As String
On Error GoTo Err

Set DataObjectFactory_SearchObjectsAsync = SearchObjectsAsync(criterion, cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get FieldSpecifiers() As FieldSpecifiers
Const ProcName As String = "FieldSpecifiers"
Dim failpoint As String
On Error GoTo Err

Set FieldSpecifiers = mBaseFactory.FieldSpecifiers

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal db As TradingDB)

Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mDB = db
Set mBaseFactory = CreateAbstractDOFactory(Me, mDB.ConnectionString, "timezone", "vTimezones")

mBaseFactory.AddField FieldNameName, TimeZoneColumnName, FieldWidthTimeZone, FieldAlignTimeZone, True
mBaseFactory.AddField FieldNameCanonicalName, TimeZoneColumnCanonicalName, FieldWidthTimeZone, FieldAlignTimeZone, True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Public Function LoadByID( _
                ByVal Id As Long) As TimeZone
Dim rs As Recordset
Const ProcName As String = "LoadByID"
Dim failpoint As String
On Error GoTo Err

Set rs = mBaseFactory.LoadByID(Id)
If Not rs Is Nothing Then
    Set LoadByID = New TimeZone
    LoadByID.Initialise mDB, rs
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function LoadByIDAsync( _
                ByVal Id As Long, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "LoadByIDAsync"
Dim failpoint As String
On Error GoTo Err

Set LoadByIDAsync = mBaseFactory.LoadByIDAsync(Id, cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function LoadByName( _
                ByVal pName As String) As TimeZone
Dim rs As Recordset

Const ProcName As String = "LoadByName"
Dim failpoint As String
On Error GoTo Err

Set rs = mBaseFactory.LoadByName(Trim$(pName))
If Not rs Is Nothing Then
    Set LoadByName = New TimeZone
    LoadByName.Initialise mDB, rs
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function LoadByNameAsync( _
                ByVal pName As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "LoadByNameAsync"
Dim failpoint As String
On Error GoTo Err

Set LoadByNameAsync = mBaseFactory.LoadByNameAsync(Trim$(pName), cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function LoadByQuery( _
                ByVal whereCondition As String) As TimeZone
Dim rs As Recordset
Const ProcName As String = "LoadByQuery"
Dim failpoint As String
On Error GoTo Err

Set rs = mBaseFactory.LoadByQuery(whereCondition)
If Not rs Is Nothing Then
    Set LoadByQuery = New TimeZone
    LoadByQuery.Initialise mDB, rs
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function LoadByQueryAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "LoadByQueryAsync"
Dim failpoint As String
On Error GoTo Err

Set LoadByQueryAsync = mBaseFactory.LoadByQueryAsync(whereCondition, cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function MakeNew() As TimeZone
Const ProcName As String = "MakeNew"
Dim failpoint As String
On Error GoTo Err

Set MakeNew = New TimeZone
MakeNew.Initialise mDB, mBaseFactory.MakeNew

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function MakeNewAsync( _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "MakeNewAsync"
Dim failpoint As String
On Error GoTo Err

Set MakeNewAsync = mBaseFactory.MakeNewAsync(cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function Query( _
                ByVal whereCondition As String, _
                ByRef fieldNames() As String) As DataObjectSummaries
Const ProcName As String = "Query"
Dim failpoint As String
On Error GoTo Err

Set Query = mBaseFactory.Query(whereCondition, fieldNames)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function QueryAsync( _
                ByVal whereCondition As String, _
                ByRef fieldNames() As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "QueryAsync"
Dim failpoint As String
On Error GoTo Err

Set QueryAsync = mBaseFactory.QueryAsync(whereCondition, fieldNames, cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function QueryObjects( _
                ByVal whereCondition As String) As DataObjects
Const ProcName As String = "QueryObjects"
Dim failpoint As String
On Error GoTo Err

Set QueryObjects = mBaseFactory.QueryObjects(whereCondition)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function QueryObjectsAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "QueryObjectsAsync"
Dim failpoint As String
On Error GoTo Err

Set QueryObjectsAsync = mBaseFactory.QueryObjectsAsync(whereCondition, cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub Refresh( _
                ByVal pTimezone As TimeZone)
Dim rs As Recordset

Const ProcName As String = "Refresh"
Dim failpoint As String
On Error GoTo Err

Set rs = mBaseFactory.LoadByID(pTimezone.Id)
If Not rs Is Nothing Then
    pTimezone.Initialise mDB, rs
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function Search( _
                ByVal criterion As String, _
                ByRef fieldNames() As String) As DataObjectSummaries
Const ProcName As String = "Search"
Dim failpoint As String
On Error GoTo Err

Set Search = mBaseFactory.Search(criterion, fieldNames)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function SearchAsync( _
                ByVal criterion As String, _
                ByRef fieldNames() As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "SearchAsync"
Dim failpoint As String
On Error GoTo Err

Set SearchAsync = mBaseFactory.SearchAsync(criterion, fieldNames, cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function SearchObjects( _
                ByVal criterion As String) As DataObjects
Const ProcName As String = "SearchObjects"
Dim failpoint As String
On Error GoTo Err

Set SearchObjects = mBaseFactory.SearchObjects(criterion)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function SearchObjectsAsync( _
                ByVal criterion As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "SearchObjectsAsync"
Dim failpoint As String
On Error GoTo Err

Set SearchObjectsAsync = mBaseFactory.SearchObjectsAsync(criterion, cookie, taskName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' Helper Functions
'@================================================================================















