VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InstrumentClassFactory"
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
Private Const ModuleName As String = "InstrumentClassFactory"

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

Set mBaseFactory = CreateAbstractDOFactory(mDB.connectionString, "InstrumentClass", "vinstrumentclasses")

mBaseFactory.addField FieldNameName, InstrumentClassColumnName, FieldWidthInstrumentClass, FieldAlignInstrumentClass, True
mBaseFactory.addField FieldNameSecType, InstrumentClassColumnSecType, FieldWidthSecType, FieldAlignSecType, True
mBaseFactory.addField FieldNameExchange, InstrumentClassColumnExchange, FieldWidthExchange, FieldAlignExchange, False
mBaseFactory.addField FieldNameCurrency, InstrumentClassColumnCurrency, FieldWidthCurrency, FieldAlignCurrency, True
mBaseFactory.addField FieldNameTickSize, InstrumentClassColumnTickSize, FieldWidthTickSize, FieldAlignTickSize, True
mBaseFactory.addField FieldNameTickValue, InstrumentClassColumnTickValue, FieldWidthTickValue, FieldAlignTickValue, True
mBaseFactory.addField FieldNameSwitchDays, InstrumentClassColumnSwitchDays, FieldWidthSwitchDays, FieldAlignSwitchDays, True
mBaseFactory.addField FieldNameSessionStartTime, InstrumentClassColumnSessionStartTime, FieldWidthSessionStartTime, FieldAlignSessionStartTime, True
mBaseFactory.addField FieldNameSessionEndTime, InstrumentClassColumnSessionEndTime, FieldWidthSessionEndTime, FieldAlignSessionEndTime, True
mBaseFactory.addField FieldNameTimeZone, InstrumentClassColumnTimeZone, FieldWidthTimeZone, FieldAlignTimeZone, False
mBaseFactory.addField FieldNameNotes, InstrumentClassColumnNotes, FieldWidthNotes, FieldAlignNotes, True

End Sub

Public Function loadByID( _
                ByVal id As Long) As InstrumentClass
Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset
Set rs = mBaseFactory.loadByID(id)
If Not rs Is Nothing Then
    Set loadByID = New InstrumentClass
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
                ByVal pName As String) As InstrumentClass
Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset
Dim exchg As Exchange
Dim cond As New SimpleConditionBuilder
Dim clp As CommandLineParser

pName = Trim$(pName)

Set clp = CreateCommandLineParser(pName, "/")

If clp.NumberOfArgs <> 2 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "loadByName", _
            "Name must be of the form exchange/instrumentclassname"
End If

Set exchg = mDB.ExchangeFactory.loadByName(clp.Arg(0))
If exchg Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "loadByName", _
            clp.Arg(0) & " is not a valid exchange name"
End If

cond.addTerm InstrumentClassColumnExchangeID, CondOpEqual, exchg.id
cond.addTerm InstrumentClassColumnName, CondOpEqual, clp.Arg(1), LogicalOpAND

Set rs = mBaseFactory.loadByQuery(cond.conditionString)
If Not rs Is Nothing Then
    Set loadByName = New InstrumentClass
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
                ByVal whereCondition As String) As InstrumentClass
Dim failpoint As Long
On Error GoTo Err

Dim rs As Recordset
Set rs = mBaseFactory.loadByQuery(whereCondition)
If Not rs Is Nothing Then
    Set loadByQuery = New InstrumentClass
    loadByQuery.initialise mDB, rs
End If

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadByQuery" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function loadSummaries( _
                ByRef fieldNames() As String, _
                Optional ByVal pExchange As String, _
                Optional ByVal pSecType As SecurityTypes, _
                Optional ByVal pCurrency As String) As DataObjectSummaries
Dim cond As New SimpleConditionBuilder

Dim failpoint As Long
On Error GoTo Err

addTerm cond, InstrumentClassColumnExchange, pExchange
addTerm cond, InstrumentClassColumnSecType, gCategoryToString(gCategoryFromSecType(pSecType))
addTerm cond, InstrumentClassColumnCurrency, pCurrency

Set loadSummaries = mBaseFactory.query(cond.conditionString, fieldNames)

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "loadSummaries" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Public Function makeNew() As InstrumentClass
Dim failpoint As Long
On Error GoTo Err

Set makeNew = New InstrumentClass
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

Private Sub addTerm( _
            ByVal cond As SimpleConditionBuilder, _
            ByVal columnName As String, _
            ByVal value As String)
If value <> "" Then
    If cond.expectingLogicalOperator Then
        cond.addTerm columnName, CondOpEqual, value, LogicalOpAND
    Else
        cond.addTerm columnName, CondOpEqual, value
    End If
End If
End Sub














