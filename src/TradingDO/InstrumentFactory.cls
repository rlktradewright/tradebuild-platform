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

Private Const ProjectName As String = "TradingDO26"
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

Private Sub DataObjectFactory_delete(ByVal id As Long)
mBaseFactory.Delete id
End Sub

Private Property Get DataObjectFactory_fieldSpecifiers() As BusObjUtils10.fieldSpecifiers
Set DataObjectFactory_fieldSpecifiers = fieldSpecifiers
End Property

Private Function DataObjectFactory_loadByID( _
                ByVal id As Long) As BusinessDataObject
Set DataObjectFactory_loadByID = loadByID(id)
End Function

Private Function DataObjectFactory_loadByIDAsync( _
                ByVal id As Long, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Set DataObjectFactory_loadByIDAsync = loadByIDAsync(id, cookie, taskName)
End Function

Private Function DataObjectFactory_loadByName( _
                ByVal pName As String) As BusinessDataObject
Set DataObjectFactory_loadByName = loadByName(Trim$(pName))
End Function

Private Function DataObjectFactory_loadByNameAsync( _
                ByVal pName As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Set DataObjectFactory_loadByNameAsync = loadByNameAsync(pName, cookie, taskName)
End Function

Private Function DataObjectFactory_loadByQuery( _
                ByVal whereCondition As String) As BusinessDataObject
Set DataObjectFactory_loadByQuery = loadByQuery(whereCondition)
End Function

Private Function DataObjectFactory_loadByQueryAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Set DataObjectFactory_loadByQueryAsync = loadByQueryAsync(whereCondition, cookie, taskName)
End Function

Private Function DataObjectFactory_makeNew() As BusinessDataObject
Set DataObjectFactory_makeNew = makeNew
End Function

Private Function DataObjectFactory_makeNewAsync( _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Set DataObjectFactory_makeNewAsync = makeNewAsync(cookie)
End Function

Private Function DataObjectFactory_makeNewFromRecordset( _
                ByVal rs As ADODB.Recordset) As BusObjUtils10.BusinessDataObject
Dim obj As New instrument
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
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Set DataObjectFactory_queryAsync = queryAsync(whereCondition, fieldNames, cookie, taskName)
End Function

Private Function DataObjectFactory_queryObjects( _
                ByVal whereCondition As String) As BusObjUtils10.DataObjects
Set DataObjectFactory_queryObjects = queryObjects(whereCondition)
End Function

Private Function DataObjectFactory_queryObjectsAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Set DataObjectFactory_queryObjectsAsync = queryObjectsAsync(whereCondition, cookie, taskName)
End Function

Private Sub DataObjectFactory_refresh( _
                ByVal dataObj As BusObjUtils10.BusinessDataObject)
refresh dataObj
End Sub

Private Function DataObjectFactory_search( _
                ByVal criterion As String, _
                fieldNames() As String) As DataObjectSummaries
Set DataObjectFactory_search = search(criterion, fieldNames)
End Function

Private Function DataObjectFactory_searchAsync( _
                ByVal criterion As String, _
                fieldNames() As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Set DataObjectFactory_searchAsync = searchAsync(criterion, fieldNames, cookie, taskName)
End Function

Private Function DataObjectFactory_searchObjects( _
                ByVal criterion As String) As BusObjUtils10.DataObjects
Set DataObjectFactory_searchObjects = searchObjects(criterion)
End Function

Private Function DataObjectFactory_searchObjectsAsync( _
                ByVal criterion As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TWUtilities30.TaskController
Set DataObjectFactory_searchObjectsAsync = searchObjectsAsync(criterion, cookie, taskName)
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

Set mBaseFactory = CreateAbstractDOFactory(Me, mDB.connectionString, "instrument", "vinstrumentdetails")

mBaseFactory.addField FieldNameName, InstrumentColumnName, FieldWidthInstrument, FieldAlignInstrument, True
mBaseFactory.addField FieldNameShortName, InstrumentColumnShortName, FieldWidthShortName, FieldAlignShortName, True
mBaseFactory.addField FieldNameExchange, InstrumentColumnExchangeName, FieldWidthExchange, FieldAlignExchange, False
mBaseFactory.addField FieldNameSymbol, InstrumentColumnSymbol, FieldWidthSymbol, FieldAlignSymbol, True
mBaseFactory.addField FieldNameSecType, InstrumentColumnSecType, FieldWidthSecType, FieldAlignSecType, False
mBaseFactory.addField FieldNameExpiry, InstrumentColumnExpiry, FieldWidthExpiry, FieldAlignExpiry, True
mBaseFactory.addField FieldNameCurrency, InstrumentColumnCurrency, FieldWidthCurrency, FieldAlignCurrency, False
mBaseFactory.addField FieldNameStrikePrice, InstrumentColumnStrikePrice, FieldWidthStrikePrice, FieldAlignStrikePrice, True
mBaseFactory.addField FieldNameOptionRight, InstrumentColumnOptionRight, FieldWidthOptionRight, FieldAlignOptionRight, True
mBaseFactory.addField FieldNameTickSize, InstrumentColumnTickSize, FieldWidthTickSize, FieldAlignTickSize, False
mBaseFactory.addField FieldNameTickValue, InstrumentColumnTickValue, FieldWidthTickValue, FieldAlignTickValue, False
mBaseFactory.addField FieldNameSessionStartTime, InstrumentColumnSessionStartTime, FieldWidthSessionStartTime, FieldAlignSessionStartTime, False
mBaseFactory.addField FieldNameSessionEndTime, InstrumentColumnSessionEndTime, FieldWidthSessionEndTime, FieldAlignSessionEndTime, False
mBaseFactory.addField FieldNameTimeZone, InstrumentColumnTimeZoneName, FieldWidthTimeZone, FieldAlignTimeZone, False
mBaseFactory.addField FieldNameNotes, InstrumentColumnNotes, FieldWidthNotes, FieldAlignNotes, True

End Sub

Public Function loadByID( _
                ByVal id As Long) As instrument
Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset
Set rs = mBaseFactory.loadByID(id)
If Not rs Is Nothing Then
    Set loadByID = New instrument
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
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set loadByIDAsync = mBaseFactory.loadByIDAsync(id, cookie, taskName)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadByIDAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function loadByName( _
                ByVal pName As String) As instrument
Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset

Set rs = mBaseFactory.loadByQuery(getNameWhereCondition(Trim$(pName)))
If Not rs Is Nothing Then
    Set loadByName = New instrument
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
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set loadByNameAsync = mBaseFactory.loadByQueryAsync(getNameWhereCondition(Trim$(pName)), cookie, taskName)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadByNameAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function loadByQuery( _
                ByVal whereCondition As String) As instrument
Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset
Set rs = mBaseFactory.loadByQuery(whereCondition)
If Not rs Is Nothing Then
    Set loadByQuery = New instrument
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
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set loadByQueryAsync = mBaseFactory.loadByQueryAsync(whereCondition, cookie, taskName)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadByQueryAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function loadBySpecifier( _
                ByVal contractSpec As ContractSpecifier) As instrument

Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset
Dim summs As DataObjectSummaries
Dim fieldNames() As String

Set summs = loadSummariesBySpec(fieldNames, contractSpec)

If summs.count <> 1 Then Exit Function

Set rs = mBaseFactory.loadByID(summs(1).id)
If Not rs Is Nothing Then
    Set loadBySpecifier = New instrument
    loadBySpecifier.initialise mDB, rs
End If

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadBySpec" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function loadBySpecifierAsync( _
                ByVal contractSpec As ContractSpecifier, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set loadBySpecifierAsync = mBaseFactory.loadByQueryAsync(conditionStringFromSpec(contractSpec), cookie, taskName)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadBySpecifierAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function loadSummariesByClass( _
                ByRef fieldNames() As String, _
                ByVal exchangeName As String, _
                ByVal instrumentClassName As String) As DataObjectSummaries
Dim cond As New SimpleConditionBuilder

Dim failpoint As Long
On Error GoTo Err

cond.addTerm InstrumentColumnExchangeName, CondOpEqual, exchangeName
cond.addTerm InstrumentColumnInstrumentClassName, CondOpEqual, instrumentClassName, LogicalOpAND
Set loadSummariesByClass = mBaseFactory.query(cond.conditionString, fieldNames)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadSummariesByClass" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function loadSummariesByClassAsync( _
                ByRef fieldNames() As String, _
                ByVal exchangeName As String, _
                ByVal instrumentClassName As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim cond As New SimpleConditionBuilder

Dim failpoint As Long
On Error GoTo Err

cond.addTerm InstrumentColumnExchangeName, CondOpEqual, exchangeName
cond.addTerm InstrumentColumnInstrumentClassName, CondOpEqual, instrumentClassName, LogicalOpAND
Set loadSummariesByClassAsync = mBaseFactory.queryAsync(cond.conditionString, fieldNames, cookie, taskName)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadSummariesByClassAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function loadSummariesBySpec( _
                ByRef fieldNames() As String, _
                ByVal contractSpec As ContractSpecifier) As DataObjectSummaries

Dim failpoint As Long
On Error GoTo Err

Set loadSummariesBySpec = mBaseFactory.query(conditionStringFromSpec(contractSpec), fieldNames)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadSummariesBySpec" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function loadSummariesBySpecAsync( _
                ByRef fieldNames() As String, _
                ByVal contractSpec As ContractSpecifier, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController

Dim failpoint As Long
On Error GoTo Err

Set loadSummariesBySpecAsync = mBaseFactory.queryAsync(conditionStringFromSpec(contractSpec), fieldNames, cookie, taskName)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadSummariesBySpecAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function makeNew() As instrument

Dim failpoint As Long
On Error GoTo Err

Set makeNew = New instrument
makeNew.initialise mDB, mBaseFactory.makeNew

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "makeNew" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function makeNewAsync( _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController

Dim failpoint As Long
On Error GoTo Err

Set makeNewAsync = mBaseFactory.makeNewAsync(cookie, taskName)

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
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set queryAsync = mBaseFactory.queryAsync(whereCondition, fieldNames, cookie, taskName)

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
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set queryObjectsAsync = mBaseFactory.queryObjectsAsync(whereCondition, cookie, taskName)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "queryObjectsAsync" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function queryObjectsBySpec( _
                ByVal contractSpec As ContractSpecifier) As DataObjects
Dim failpoint As Long
On Error GoTo Err

Set queryObjectsBySpec = mBaseFactory.queryObjects(conditionStringFromSpec(contractSpec))

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "queryObjectsBySpec" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function queryObjectsBySpecAsync( _
                ByVal contractSpec As ContractSpecifier, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set queryObjectsBySpecAsync = mBaseFactory.queryObjectsAsync(conditionStringFromSpec(contractSpec), cookie, taskName)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "queryObjectsBySpecAsync" & "." & failpoint & _
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

Friend Sub refresh( _
                ByVal pInstrument As instrument)
Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset
Set rs = mBaseFactory.loadByID(pInstrument.id)
If Not rs Is Nothing Then
    pInstrument.initialise mDB, rs
End If

Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "reFresh" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Sub

Public Function searchAsync( _
                ByVal criterion As String, _
                ByRef fieldNames() As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set searchAsync = mBaseFactory.searchAsync(criterion, fieldNames, cookie, taskName)

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
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim failpoint As Long
On Error GoTo Err

Set searchObjectsAsync = mBaseFactory.searchObjectsAsync(criterion, cookie, taskName)

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

Private Sub addTerm( _
            ByVal cond As SimpleConditionBuilder, _
            ByVal columnName As String, _
            ByVal op As ConditionalOperators, _
            ByVal value As String)
If value <> "" Then
    If cond.expectingLogicalOperator Then
        cond.addTerm columnName, op, value, LogicalOpAND
    Else
        cond.addTerm columnName, op, value
    End If
End If
End Sub

Private Function conditionStringFromSpec( _
            ByVal contractSpec As ContractSpecifier) As String

Dim cond As New SimpleConditionBuilder
With contractSpec
    addTerm cond, InstrumentColumnShortName, CondOpEqual, .localSymbol
    addTerm cond, InstrumentColumnSymbol, CondOpEqual, .symbol
    addTerm cond, InstrumentColumnExchangeName, CondOpEqual, .Exchange
    addTerm cond, InstrumentColumnSecType, CondOpEqual, gCategoryToString(gCategoryFromSecType(.secType))
    addTerm cond, InstrumentColumnCurrencyE, CondOpEqual, .currencyCode
    
    Select Case .secType
    Case SecTypeFuture
        addTerm cond, InstrumentColumnExpiryMonth, CondOpLike, .expiry & "%"
    Case SecTypeOption, SecTypeFuturesOption
        addTerm cond, InstrumentColumnExpiryMonth, CondOpLike, .expiry & "%"
        If .strike <> 0# Then addTerm cond, InstrumentColumnStrikePrice, CondOpEqual, .strike
        If .Right <> OptNone Then addTerm cond, InstrumentColumnOptionRight, CondOpEqual, Left$(OptionRightToString(.Right), 1)
    End Select
End With

conditionStringFromSpec = cond.conditionString

End Function

Private Function getNameWhereCondition( _
                ByVal pName As String) As String
Dim instrClass As InstrumentClass
Dim cond As New SimpleConditionBuilder
Dim clp As CommandLineParser

pName = Trim$(pName)

Set clp = CreateCommandLineParser(pName, "/")

If clp.NumberOfArgs <> 3 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "getNameWhereCondition", _
            "Name must be of the form exchangename/instrumentclassname/instrumentname"
End If

Set instrClass = mDB.InstrumentClassFactory.loadByName("""" & clp.Arg(0) & """/""" & clp.Arg(1) & """")
If instrClass Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "getNameWhereCondition", _
            clp.Arg(0) & "/" & clp.Arg(1) & " is not a valid instrumentclass name"
End If

cond.addTerm InstrumentColumnInstrumentClassID, CondOpEqual, instrClass.id
cond.addTerm InstrumentColumnName, CondOpEqual, clp.Arg(2), LogicalOpAND

getNameWhereCondition = cond.conditionString
End Function









