VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InstrumentClass"
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

Implements BusinessDataObject

'@===============================================================================
' Events
'@===============================================================================

Public Event Clean()
Public Event Deleted()
Public Event Dirty()
Public Event Invalid()
Public Event Valid()

'@===============================================================================
' Constants
'@===============================================================================

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

' Generic declarations -----------------------------------------------------------
Private WithEvents mDataObj                         As AbstractDataObject
Attribute mDataObj.VB_VarHelpID = -1

Private mDB                                         As TradingDB

' Custom declarations ------------------------------------------------------------

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get currencyCode() As String
currencyCode = mDataObj.getColumnValue(InstrumentClassColumnCurrency, "")
End Property

Public Property Let currencyCode(ByVal value As String)
value = Trim$(value)
mDataObj.ruleBroken BusRuleInstrumentClassCurrencyCodeValid, _
                    Len(value) = 0 Or Len(value) > mDataObj.columnMaxLength(InstrumentClassColumnCurrency) Or _
                    Not IsValidCurrencyCode(value)
mDataObj.columnValue(InstrumentClassColumnCurrency) = Left$(value, mDataObj.columnMaxLength(InstrumentClassColumnCurrency))
End Property

Public Property Get daysBeforeExpiryToSwitch() As Long
daysBeforeExpiryToSwitch = mDataObj.getColumnValue(InstrumentClassColumnSwitchDays, 0)
End Property

Public Property Let daysBeforeExpiryToSwitch(ByVal value As Long)
mDataObj.ruleBroken BusRuleInstrumentClassDaysBeforeExpiryValid, value < 0
mDataObj.columnValue(InstrumentClassColumnSwitchDays) = value
End Property

Public Property Let daysBeforeExpiryToSwitchString(ByVal value As String)
If value = "" Then Exit Property
If mDataObj.ruleBroken(BusRuleInstrumentClassDaysBeforeExpiryValid, Not IsInteger(value)) Then
    daysBeforeExpiryToSwitch = 999
Else
    daysBeforeExpiryToSwitch = CLng(value)
End If
End Property

Public Property Let Exchange( _
                ByVal value As Exchange)
Dim instrCl As InstrumentClass
If mDataObj.ruleBroken(BusRuleInstrumentClassExchangeValid, value Is Nothing) Then Exit Property
    
If name <> "" Then
    Set instrCl = mDB.InstrumentClassFactory.loadByName(value.name & "/" & name)
    If Not instrCl Is Nothing Then mDataObj.ruleBroken getBusRuleNameValidId, instrCl.id <> id
End If
mDataObj.columnValue(InstrumentClassColumnExchangeID) = value.id
End Property

Public Property Get Exchange() As Exchange
Set Exchange = mDB.ExchangeFactory.loadByID(mDataObj.getColumnValue(InstrumentClassColumnExchangeID, 0))
End Property

Public Property Let exchangeName( _
                ByVal value As String)
Exchange = mDB.ExchangeFactory.loadByName(value)
End Property

Public Property Get exchangeName() As String
If mDataObj.IsNew Then
    If CLng(mDataObj.getColumnValue(InstrumentClassColumnExchangeID, 0)) <> 0 Then
        exchangeName = mDB.ExchangeFactory.loadByID(mDataObj.getColumnValue(InstrumentClassColumnExchangeID)).name
    End If
Else
    exchangeName = mDataObj.getColumnValue(InstrumentClassColumnExchange)
End If
End Property

Public Property Get name() As String
name = getName
End Property

Public Property Let name(ByVal value As String)
' overrides generic property
value = Trim$(value)
If value = getName Then Exit Property
If Not mDataObj.ruleBroken(getBusRuleNameValidId, _
                            Len(value) = 0 Or _
                            Len(value) > mDataObj.columnMaxLength(GenericColumnName)) Then
    If exchangeName = "" Then
        mDataObj.columnValue(GenericColumnName) = Left$(value, mDataObj.columnMaxLength(GenericColumnName))
    ElseIf Not mDataObj.ruleBroken(getBusRuleNameValidId, _
                                    Not mDB.InstrumentClassFactory.loadByName(exchangeName & "/" & name) Is Nothing) Then
        mDataObj.columnValue(GenericColumnName) = Left$(value, mDataObj.columnMaxLength(GenericColumnName))
    End If
End If

End Property

Public Property Get notes() As String
notes = mDataObj.getColumnValue(InstrumentClassColumnNotes, "")
End Property

Public Property Let notes(ByVal val As String)
mDataObj.columnValue(InstrumentClassColumnNotes) = val
End Property

Public Property Let secType(ByVal value As SecurityTypes)
Dim catid As InstrumentCategories

If Not mDataObj.ruleBroken(BusRuleInstrumentClassSecTypeValid, Not IsValidSecType(value)) Then
    catid = gCategoryFromSecType(value)
    mDataObj.columnValue(InstrumentClassColumnSecTypeId) = catid
Else
    mDataObj.ruleBroken BusRuleInstrumentClassSecTypeValid, True
End If
End Property

Public Property Get secType() As SecurityTypes
secType = gCategoryToSecType(mDataObj.getColumnValue(InstrumentClassColumnSecTypeId, 0))
End Property

Public Property Let secTypeString(ByVal value As String)
value = Trim$(value)
If mDataObj.ruleBroken(BusRuleInstrumentClassSecTypeValid, _
                    (SecTypeFromString(value) = SecTypeNone)) Then
    secType = SecTypeNone
Else
    secType = SecTypeFromString(value)
End If
End Property

Public Property Get secTypeString() As String
secTypeString = SecTypeToString(secType)
End Property

Public Property Get sessionEndTime() As Date
sessionEndTime = mDataObj.getColumnValue(InstrumentClassColumnSessionEndTime, CDate(0))
sessionEndTime = sessionEndTime - Int(sessionEndTime)
End Property

Public Property Let sessionEndTime( _
                ByVal value As Date)
mDataObj.ruleBroken BusRuleInstrumentClassSessionEndTimeValid, value >= 1
mDataObj.columnValue(InstrumentClassColumnSessionEndTime) = value
End Property

Public Property Let sessionEndTimeString( _
                ByVal value As String)
If Not mDataObj.ruleBroken(BusRuleInstrumentClassSessionEndTimeValid, Not IsDate(value)) Then
    sessionEndTime = CDate(value)
End If
End Property

Public Property Get sessionStartTime() As Date
sessionStartTime = mDataObj.getColumnValue(InstrumentClassColumnSessionStartTime, CDate(0))
sessionStartTime = sessionStartTime - Int(sessionStartTime)
End Property

Public Property Let sessionStartTime( _
                ByVal value As Date)
mDataObj.ruleBroken BusRuleInstrumentClassSessionStartTimeValid, value >= 1
mDataObj.columnValue(InstrumentClassColumnSessionStartTime) = value
End Property

Public Property Let sessionStartTimeString( _
                ByVal value As String)
If Not mDataObj.ruleBroken(BusRuleInstrumentClassSessionStartTimeValid, Not IsDate(value)) Then
    sessionStartTime = CDate(value)
End If
End Property

Public Property Let tickSize( _
                ByVal value As Double)
mDataObj.ruleBroken BusRuleInstrumentClassTickSizeValid, value <= 0
mDataObj.columnValue(InstrumentClassColumnTickSize) = value
End Property

Public Property Get tickSize() As Double
tickSize = mDataObj.getColumnValue(InstrumentClassColumnTickSize, 0)
End Property

Public Property Let tickSizeString( _
                ByVal value As String)
If mDataObj.ruleBroken(BusRuleInstrumentClassTickSizeValid, Not IsNumeric(value)) Then
    tickSize = 999.999
Else
    tickSize = CDbl(value)
End If
End Property

Public Property Let tickValue(ByVal value As Double)
mDataObj.ruleBroken BusRuleInstrumentClassTickValueValid, value <= 0
mDataObj.columnValue(InstrumentClassColumnTickValue) = value
End Property

Public Property Get tickValue() As Double
tickValue = mDataObj.getColumnValue(InstrumentClassColumnTickValue, 0)
End Property

Public Property Let tickValueString( _
                ByVal value As String)
If mDataObj.ruleBroken(BusRuleInstrumentClassTickValueValid, Not IsNumeric(value)) Then
    tickValue = 999.999
Else
    tickValue = CDbl(value)
End If
End Property

Public Property Get TimeZoneName() As String
If Not mDataObj.IsNew Then
    TimeZoneName = mDataObj.getColumnValue(InstrumentClassColumnTimeZone, "")
End If
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal db As TradingDB, _
                Optional ByVal rs As ADODB.Recordset)
genericInitialise db, rs
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getBusRuleNameValidId() As Long
getBusRuleNameValidId = BusRuleInstrumentClassNameValid
End Function

Private Function getFactory() As DataObjectFactory
Set getFactory = mDB.InstrumentClassFactory
End Function

Private Sub setInitialErrors()
' set errors for fields that are invalid when a new record is created
' Note that this is done automatically by the AbstractDataObject for
' the name field
mDataObj.ruleBroken BusRuleInstrumentClassExchangeValid, True
mDataObj.ruleBroken BusRuleInstrumentClassSecTypeValid, True
mDataObj.ruleBroken BusRuleInstrumentClassCurrencyCodeValid, True
mDataObj.ruleBroken BusRuleInstrumentClassTickSizeValid, True
mDataObj.ruleBroken BusRuleInstrumentClassTickValueValid, True
mDataObj.ruleBroken BusRuleInstrumentClassSessionStartTimeValid, True
mDataObj.ruleBroken BusRuleInstrumentClassSessionEndTimeValid, True
End Sub


'@********************************************************************************
'@********************************************************************************
' Beyond this point everything is generic
'@********************************************************************************
'@********************************************************************************


'@================================================================================
' BusinessDataObject Interface Members
'@================================================================================

Private Sub BusinessDataObject_ApplyEdit()
ApplyEdit
End Sub

Private Sub BusinessDataObject_CancelEdit()
CancelEdit
End Sub

Private Sub BusinessDataObject_Delete()
Delete
End Sub

Private Property Get BusinessDataObject_ErrorList() As ErrorList
Set BusinessDataObject_ErrorList = mDataObj.ErrorList
End Property

Private Property Get BusinessDataObject_EventSource() As BusinessObjEventSource
Set BusinessDataObject_EventSource = mDataObj.eventSource
End Property

Private Property Get BusinessDataObject_ID() As Long
BusinessDataObject_ID = id
End Property

Private Property Get BusinessDataObject_IsDeleted() As Boolean
BusinessDataObject_IsDeleted = IsDeleted
End Property

Private Property Get BusinessDataObject_IsDirty() As Boolean
BusinessDataObject_IsDirty = IsDirty
End Property

Private Property Get BusinessDataObject_IsNew() As Boolean
BusinessDataObject_IsNew = IsNew
End Property

Private Property Get BusinessDataObject_IsValid() As Boolean
BusinessDataObject_IsValid = IsValid
End Property

Private Property Let BusinessDataObject_Name(ByVal RHS As String)
name = RHS
End Property

Private Property Get BusinessDataObject_Name() As String
BusinessDataObject_Name = name
End Property

Private Sub BusinessDataObject_Refresh()
refresh
End Sub

'@===============================================================================
' mDataObj Event Handlers
'@===============================================================================

Private Sub mDataObj_Clean()
RaiseEvent Clean
End Sub

Private Sub mDataObj_Deleted()
RaiseEvent Deleted
End Sub

Private Sub mDataObj_Dirty()
RaiseEvent Dirty
End Sub

Private Sub mDataObj_Invalid()
RaiseEvent Invalid
End Sub

Private Sub mDataObj_Valid()
RaiseEvent Valid
End Sub

'@================================================================================
' Generic Properties
'@================================================================================

Public Property Get ErrorList() As ErrorList
Set ErrorList = mDataObj.ErrorList
End Property

Public Property Get id() As Long
id = mDataObj.id
End Property

Public Property Get IsValid() As Boolean
IsValid = mDataObj.IsValid
End Property

Public Property Get IsDeleted() As Boolean
IsDeleted = mDataObj.IsDeleted
End Property

Public Property Get IsDirty() As Boolean
IsDirty = mDataObj.IsDirty
End Property

Public Property Get IsNew() As Boolean
IsNew = mDataObj.IsNew
End Property

'@================================================================================
' Generic Methods
'@================================================================================

Public Sub ApplyEdit()
mDataObj.ApplyEdit
End Sub

Public Sub CancelEdit()
mDataObj.CancelEdit
If mDataObj.IsNew Then setInitialErrors
End Sub

Public Sub Delete()
mDataObj.Delete
End Sub

Public Sub genericInitialise( _
                ByVal db As Variant, _
                Optional ByVal rs As ADODB.Recordset)
Set mDB = db
If Not rs Is Nothing Then
    Set mDataObj = CreateAbstractDataObject(Me, getBusRuleNameValidId, getFactory, mDB.connectionString)
    mDataObj.initialiseRs rs
    If mDataObj.IsNew Then setInitialErrors
End If
End Sub

Public Function getName() As String
getName = mDataObj.name
End Function

Public Sub refresh()
mDataObj.refresh
End Sub

Public Sub setName(ByVal value As String)
mDataObj.name = value
End Sub


