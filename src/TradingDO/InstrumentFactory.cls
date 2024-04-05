VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InstrumentFactory"
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


Private Const ModuleName As String = "InstrumentFactory"

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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get DataObjectFactory_FieldSpecifiers() As FieldSpecifiers
Const ProcName As String = "DataObjectFactory_FieldSpecifiers"
On Error GoTo Err

Set DataObjectFactory_FieldSpecifiers = FieldSpecifiers

Exit Property

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Property

Private Function DataObjectFactory_LoadByID( _
                ByVal Id As Long) As BusinessDataObject
Const ProcName As String = "DataObjectFactory_LoadByID"
On Error GoTo Err

Set DataObjectFactory_LoadByID = LoadByID(Id)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_LoadByName( _
                ByVal pName As String) As BusinessDataObject
Const ProcName As String = "DataObjectFactory_LoadByName"
On Error GoTo Err

Set DataObjectFactory_LoadByName = LoadByName(Trim$(pName))

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_LoadByQuery( _
                ByVal whereCondition As String) As BusinessDataObject
Const ProcName As String = "DataObjectFactory_LoadByQuery"
On Error GoTo Err

Set DataObjectFactory_LoadByQuery = LoadByQuery(whereCondition)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_MakeNew() As BusinessDataObject
Const ProcName As String = "DataObjectFactory_MakeNew"
On Error GoTo Err

Set DataObjectFactory_MakeNew = MakeNew

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_MakeNewAsync( _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "DataObjectFactory_MakeNewAsync"
On Error GoTo Err

Set DataObjectFactory_MakeNewAsync = MakeNewAsync(cookie)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_makeNewFromRecordset( _
                ByVal rs As ADODB.Recordset) As BusinessDataObject
Const ProcName As String = "DataObjectFactory_makeNewFromRecordset"
On Error GoTo Err

Dim obj As New instrument
obj.Initialise mDB, rs
Set DataObjectFactory_makeNewFromRecordset = obj

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_Query( _
                ByVal whereCondition As String, _
                fieldNames() As String) As DataObjectSummaries
Const ProcName As String = "DataObjectFactory_Query"
On Error GoTo Err

Set DataObjectFactory_Query = Query(whereCondition, fieldNames)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_QueryObjects( _
                ByVal whereCondition As String) As DataObjects
Const ProcName As String = "DataObjectFactory_QueryObjects"
On Error GoTo Err

Set DataObjectFactory_QueryObjects = QueryObjects(whereCondition)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub DataObjectFactory_Refresh( _
                ByVal dataObj As BusinessDataObject)
Const ProcName As String = "DataObjectFactory_Refresh"
On Error GoTo Err

Refresh dataObj

Exit Sub

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function DataObjectFactory_Search( _
                ByVal criterion As String, _
                fieldNames() As String) As DataObjectSummaries
Const ProcName As String = "DataObjectFactory_Search"
On Error GoTo Err

Set DataObjectFactory_Search = Search(criterion, fieldNames)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function DataObjectFactory_SearchObjects( _
                ByVal criterion As String) As DataObjects
Const ProcName As String = "DataObjectFactory_SearchObjects"
On Error GoTo Err

Set DataObjectFactory_SearchObjects = SearchObjects(criterion)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal db As TradingDB)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mDB = db

Set mBaseFactory = CreateAbstractDOFactory(Me, mDB.ConnectionString, "instrument", "vinstrumentdetails")

mBaseFactory.AddField FieldNameName, InstrumentColumnName, FieldWidthInstrument, FieldAlignInstrument, True
mBaseFactory.AddField FieldNameShortName, InstrumentColumnShortName, FieldWidthShortName, FieldAlignShortName, True
mBaseFactory.AddField FieldNameExchange, InstrumentColumnExchangeName, FieldWidthExchange, FieldAlignExchange, False
mBaseFactory.AddField FieldNameSymbol, InstrumentColumnSymbol, FieldWidthSymbol, FieldAlignSymbol, True
mBaseFactory.AddField FieldNameSecType, InstrumentColumnSecType, FieldWidthSecType, FieldAlignSecType, False
mBaseFactory.AddField FieldNameExpiry, InstrumentColumnExpiry, FieldWidthExpiry, FieldAlignExpiry, True
mBaseFactory.AddField FieldNameCurrency, InstrumentColumnCurrency, FieldWidthCurrency, FieldAlignCurrency, False
mBaseFactory.AddField FieldNameStrikePrice, InstrumentColumnStrikePrice, FieldWidthStrikePrice, FieldAlignStrikePrice, True
mBaseFactory.AddField FieldNameOptionRight, InstrumentColumnOptionRight, FieldWidthOptionRight, FieldAlignOptionRight, True
mBaseFactory.AddField FieldNameTickSize, InstrumentColumnTickSize, FieldWidthTickSize, FieldAlignTickSize, False
mBaseFactory.AddField FieldNameTickValue, InstrumentColumnTickValue, FieldWidthTickValue, FieldAlignTickValue, False
mBaseFactory.AddField FieldNameSessionStartTime, InstrumentColumnSessionStartTime, FieldWidthSessionStartTime, FieldAlignSessionStartTime, False
mBaseFactory.AddField FieldNameSessionEndTime, InstrumentColumnSessionEndTime, FieldWidthSessionEndTime, FieldAlignSessionEndTime, False
mBaseFactory.AddField FieldNameTimeZone, InstrumentColumnTimeZoneName, FieldWidthTimeZone, FieldAlignTimeZone, False
mBaseFactory.AddField FieldNameNotes, InstrumentColumnNotes, FieldWidthNotes, FieldAlignNotes, True

Exit Sub

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName

End Sub

Public Function LoadByID( _
                ByVal Id As Long) As instrument
Const ProcName As String = "LoadByID"
On Error GoTo Err

Dim rs As Recordset
Set rs = mBaseFactory.LoadByID(Id)
If Not rs Is Nothing Then
    Set LoadByID = New instrument
    LoadByID.Initialise mDB, rs
End If

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadByName( _
                ByVal pName As String) As instrument
Const ProcName As String = "LoadByName"
On Error GoTo Err

Dim rs As Recordset
Set rs = mBaseFactory.LoadByQuery(getNameWhereCondition(Trim$(pName)))
If Not rs Is Nothing Then
    Set LoadByName = New instrument
    LoadByName.Initialise mDB, rs
End If

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadByNameAsync( _
                ByVal pName As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "LoadByNameAsync"
On Error GoTo Err

Set LoadByNameAsync = mBaseFactory.LoadByQueryAsync(getNameWhereCondition(Trim$(pName)), cookie, taskName)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadByQuery( _
                ByVal whereCondition As String) As instrument
Const ProcName As String = "LoadByQuery"
On Error GoTo Err

Dim rs As Recordset
Set rs = mBaseFactory.LoadByQuery(whereCondition)
If Not rs Is Nothing Then
    Set LoadByQuery = New instrument
    LoadByQuery.Initialise mDB, rs
End If

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadBySpecifier( _
                ByVal contractSpec As IContractSpecifier) As instrument
Const ProcName As String = "LoadBySpecifier"
On Error GoTo Err

Dim summs As DataObjectSummaries
Dim fieldNames() As String
Set summs = LoadSummariesBySpec(fieldNames, contractSpec)

If summs.Count <> 1 Then Exit Function

Dim rs As Recordset
Set rs = mBaseFactory.LoadByID(summs(1).Id)
If Not rs Is Nothing Then
    Set LoadBySpecifier = New instrument
    LoadBySpecifier.Initialise mDB, rs
End If

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadBySpecifierAsync( _
                ByVal contractSpec As IContractSpecifier, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "LoadBySpecifierAsync"
On Error GoTo Err

Set LoadBySpecifierAsync = mBaseFactory.LoadByQueryAsync(conditionStringFromSpec(contractSpec), cookie, taskName)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadSummariesByClass( _
                ByRef fieldNames() As String, _
                ByVal ExchangeName As String, _
                ByVal InstrumentClassName As String) As DataObjectSummaries
Const ProcName As String = "LoadSummariesByClass"
On Error GoTo Err

Dim cond As New SimpleConditionBuilder
cond.addTerm InstrumentColumnExchangeName, CondOpEqual, ExchangeName
cond.addTerm InstrumentColumnInstrumentClassName, CondOpEqual, InstrumentClassName, LogicalOpAND
Set LoadSummariesByClass = mBaseFactory.Query(cond.conditionString, fieldNames)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadSummariesByClassAsync( _
                ByRef fieldNames() As String, _
                ByVal ExchangeName As String, _
                ByVal InstrumentClassName As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "LoadSummariesByClassAsync"
On Error GoTo Err

Dim cond As New SimpleConditionBuilder
cond.addTerm InstrumentColumnExchangeName, CondOpEqual, ExchangeName
cond.addTerm InstrumentColumnInstrumentClassName, CondOpEqual, InstrumentClassName, LogicalOpAND
Set LoadSummariesByClassAsync = mBaseFactory.QueryAsync(cond.conditionString, fieldNames, cookie, taskName)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadSummariesBySpec( _
                ByRef fieldNames() As String, _
                ByVal contractSpec As IContractSpecifier) As DataObjectSummaries
Const ProcName As String = "LoadSummariesBySpec"
On Error GoTo Err

Set LoadSummariesBySpec = mBaseFactory.Query(conditionStringFromSpec(contractSpec), fieldNames)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadSummariesBySpecAsync( _
                ByRef fieldNames() As String, _
                ByVal contractSpec As IContractSpecifier, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "LoadSummariesBySpecAsync"
On Error GoTo Err

Set LoadSummariesBySpecAsync = mBaseFactory.QueryAsync(conditionStringFromSpec(contractSpec), fieldNames, cookie, taskName)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function MakeNew() As instrument
Const ProcName As String = "MakeNew"
On Error GoTo Err

Set MakeNew = New instrument
MakeNew.Initialise mDB, mBaseFactory.MakeNew

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function MakeNewAsync( _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "MakeNewAsync"
On Error GoTo Err

Set MakeNewAsync = mBaseFactory.MakeNewAsync(cookie, taskName)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function Query( _
                ByVal whereCondition As String, _
                ByRef fieldNames() As String) As DataObjectSummaries
Const ProcName As String = "Query"
On Error GoTo Err

Set Query = mBaseFactory.Query(whereCondition, fieldNames)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function QueryObjects( _
                ByVal whereCondition As String) As DataObjects
Const ProcName As String = "QueryObjects"
On Error GoTo Err

Set QueryObjects = mBaseFactory.QueryObjects(whereCondition)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function QueryObjectsBySpec( _
                ByVal contractSpec As IContractSpecifier) As DataObjects
Const ProcName As String = "QueryObjectsBySpec"
On Error GoTo Err

Set QueryObjectsBySpec = mBaseFactory.QueryObjects(conditionStringFromSpec(contractSpec))

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function QueryObjectsBySpecAsync( _
                ByVal contractSpec As IContractSpecifier, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "QueryObjectsBySpecAsync"
On Error GoTo Err

Set QueryObjectsBySpecAsync = mBaseFactory.QueryObjectsAsync(conditionStringFromSpec(contractSpec), cookie, taskName)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Refresh( _
                ByVal pInstrument As instrument)
Const ProcName As String = "Refresh"
On Error GoTo Err

Dim rs As Recordset
Set rs = mBaseFactory.LoadByID(pInstrument.Id)
If Not rs Is Nothing Then
    pInstrument.Initialise mDB, rs
End If

Exit Sub

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Search( _
                ByVal criterion As String, _
                ByRef fieldNames() As String) As DataObjectSummaries
Const ProcName As String = "Search"
On Error GoTo Err

Set Search = mBaseFactory.Search(criterion, fieldNames)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function SearchObjects( _
                ByVal criterion As String) As DataObjects
Const ProcName As String = "SearchObjects"
On Error GoTo Err

Set SearchObjects = mBaseFactory.SearchObjects(criterion)

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addTerm( _
            ByVal cond As SimpleConditionBuilder, _
            ByVal columnName As String, _
            ByVal op As ConditionalOperators, _
            ByVal Value As String)
Const ProcName As String = "addTerm"
On Error GoTo Err

If Value <> "" Then
    If cond.expectingLogicalOperator Then
        cond.addTerm columnName, op, Value, LogicalOpAND
    Else
        cond.addTerm columnName, op, Value
    End If
End If

Exit Sub

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function conditionStringFromSpec( _
            ByVal contractSpec As IContractSpecifier) As String
Const ProcName As String = "conditionStringFromSpec"
On Error GoTo Err

Dim cond As New SimpleConditionBuilder
With contractSpec
    addTerm cond, InstrumentColumnShortName, CondOpEqual, .LocalSymbol
    addTerm cond, InstrumentColumnSymbol, CondOpEqual, .Symbol
    addTerm cond, InstrumentColumnExchangeName, CondOpEqual, .Exchange
    addTerm cond, InstrumentColumnSecType, CondOpEqual, CategoryToString(CategoryFromSecType(.SecType))
    addTerm cond, InstrumentColumnCurrencyE, CondOpEqual, .CurrencyCode
    If .Multiplier <> 0 Then addTerm cond, InstrumentColumnTickValueE & "/" & InstrumentColumnTickSizeE, CondOpEqual, .Multiplier
    
    Select Case .SecType
    Case SecTypeFuture
        If Len(.Expiry) >= 6 Then addTerm cond, InstrumentColumnExpiryMonth, CondOpLike, .Expiry & "%"
    Case SecTypeOption, SecTypeFuturesOption
        If Len(.Expiry) >= 6 Then addTerm cond, InstrumentColumnExpiryMonth, CondOpLike, .Expiry & "%"
        If .Strike <> 0# Then addTerm cond, InstrumentColumnStrikePrice, CondOpEqual, .Strike
        If .Right <> OptNone Then addTerm cond, InstrumentColumnOptionRight, CondOpEqual, Left$(OptionRightToString(.Right), 1)
    End Select
End With

conditionStringFromSpec = cond.conditionString

GTDO.Logger.Log "Query condition = " & conditionStringFromSpec, ProcName, ModuleName, LogLevelDetail

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName

End Function

Private Function getNameWhereCondition( _
                ByVal pName As String) As String
Const ProcName As String = "getNameWhereCondition"
On Error GoTo Err

pName = Trim$(pName)

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(pName, "/")

AssertArgument clp.NumberOfArgs = 3, "Name must be of the form exchangename/instrumentclassname/instrumentname"

Dim instrClass As InstrumentClass
Set instrClass = mDB.InstrumentClassFactory.LoadByName(clp.Arg(0) & "/" & clp.Arg(1))
AssertArgument Not instrClass Is Nothing, clp.Arg(0) & "/" & clp.Arg(1) & " is not a valid instrumentclass Name"

Dim cond As New SimpleConditionBuilder
cond.addTerm InstrumentColumnInstrumentClassID, CondOpEqual, instrClass.Id
cond.addTerm InstrumentColumnName, CondOpEqual, clp.Arg(2), LogicalOpAND

getNameWhereCondition = cond.conditionString

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function









