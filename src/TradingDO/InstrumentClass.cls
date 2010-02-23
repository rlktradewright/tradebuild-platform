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

Private Const ModuleName                As String = "InstrumentClass"

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

Public Property Get CurrencyCode() As String
Const ProcName As String = "CurrencyCode"
Dim failpoint As String
On Error GoTo Err

CurrencyCode = mDataObj.GetColumnValue(InstrumentClassColumnCurrency, "")

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let CurrencyCode(ByVal value As String)
Const ProcName As String = "CurrencyCode"
Dim failpoint As String
On Error GoTo Err

value = Trim$(value)
mDataObj.RuleBroken BusRuleInstrumentClassCurrencyCodeValid, _
                    Len(value) = 0 Or Len(value) > mDataObj.ColumnMaxLength(InstrumentClassColumnCurrency) Or _
                    Not IsValidCurrencyCode(value)
mDataObj.ColumnValue(InstrumentClassColumnCurrency) = Left$(value, mDataObj.ColumnMaxLength(InstrumentClassColumnCurrency))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DaysBeforeExpiryToSwitch() As Long
Const ProcName As String = "DaysBeforeExpiryToSwitch"
Dim failpoint As String
On Error GoTo Err

DaysBeforeExpiryToSwitch = mDataObj.GetColumnValue(InstrumentClassColumnSwitchDays, 0)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let DaysBeforeExpiryToSwitch(ByVal value As Long)
Const ProcName As String = "DaysBeforeExpiryToSwitch"
Dim failpoint As String
On Error GoTo Err

mDataObj.RuleBroken BusRuleInstrumentClassDaysBeforeExpiryValid, value < 0
mDataObj.ColumnValue(InstrumentClassColumnSwitchDays) = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let DaysBeforeExpiryToSwitchString(ByVal value As String)
Const ProcName As String = "DaysBeforeExpiryToSwitchString"
Dim failpoint As String
On Error GoTo Err

If value = "" Then Exit Property
If mDataObj.RuleBroken(BusRuleInstrumentClassDaysBeforeExpiryValid, Not IsInteger(value)) Then
    DaysBeforeExpiryToSwitch = 999
Else
    DaysBeforeExpiryToSwitch = CLng(value)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Exchange( _
                ByVal value As Exchange)
Dim instrCl As InstrumentClass
Const ProcName As String = "Exchange"
Dim failpoint As String
On Error GoTo Err

If mDataObj.RuleBroken(BusRuleInstrumentClassExchangeValid, value Is Nothing) Then Exit Property
    
If Name <> "" Then
    Set instrCl = mDB.InstrumentClassFactory.LoadByName(value.Name & "/" & Name)
    If Not instrCl Is Nothing Then mDataObj.RuleBroken getBusRuleNameValidId, instrCl.Id <> Id
End If
mDataObj.ColumnValue(InstrumentClassColumnExchangeID) = value.Id

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Exchange() As Exchange
Const ProcName As String = "Exchange"
Dim failpoint As String
On Error GoTo Err

Set Exchange = mDB.ExchangeFactory.LoadByID(mDataObj.GetColumnValue(InstrumentClassColumnExchangeID, 0))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ExchangeName( _
                ByVal value As String)
Const ProcName As String = "ExchangeName"
Dim failpoint As String
On Error GoTo Err

Exchange = mDB.ExchangeFactory.LoadByName(value)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ExchangeName() As String
Const ProcName As String = "ExchangeName"
Dim failpoint As String
On Error GoTo Err

If mDataObj.IsNew Then
    If CLng(mDataObj.GetColumnValue(InstrumentClassColumnExchangeID, 0)) <> 0 Then
        ExchangeName = mDB.ExchangeFactory.LoadByID(mDataObj.GetColumnValue(InstrumentClassColumnExchangeID)).Name
    End If
Else
    ExchangeName = mDataObj.GetColumnValue(InstrumentClassColumnExchange)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Name() As String
Name = GetName
End Property

Public Property Let Name(ByVal value As String)
' overrides generic property
Const ProcName As String = "Name"
Dim failpoint As String
On Error GoTo Err

value = Trim$(value)
If value = GetName Then Exit Property
If Not mDataObj.RuleBroken(getBusRuleNameValidId, _
                            Len(value) = 0 Or _
                            Len(value) > mDataObj.ColumnMaxLength(GenericColumnName)) Then
    If ExchangeName = "" Then
        mDataObj.ColumnValue(GenericColumnName) = Left$(value, mDataObj.ColumnMaxLength(GenericColumnName))
    ElseIf Not mDataObj.RuleBroken(getBusRuleNameValidId, _
                                    Not mDB.InstrumentClassFactory.LoadByName(ExchangeName & "/" & Name) Is Nothing) Then
        mDataObj.ColumnValue(GenericColumnName) = Left$(value, mDataObj.ColumnMaxLength(GenericColumnName))
    End If
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Property

Public Property Get Notes() As String
Const ProcName As String = "Notes"
Dim failpoint As String
On Error GoTo Err

Notes = mDataObj.GetColumnValue(InstrumentClassColumnNotes, "")

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Notes(ByVal val As String)
Const ProcName As String = "Notes"
Dim failpoint As String
On Error GoTo Err

mDataObj.ColumnValue(InstrumentClassColumnNotes) = val

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let SecType(ByVal value As SecurityTypes)
Dim catid As InstrumentCategories

Const ProcName As String = "SecType"
Dim failpoint As String
On Error GoTo Err

If Not mDataObj.RuleBroken(BusRuleInstrumentClassSecTypeValid, Not IsValidSecType(value)) Then
    catid = gCategoryFromSecType(value)
    mDataObj.ColumnValue(InstrumentClassColumnSecTypeId) = catid
Else
    mDataObj.RuleBroken BusRuleInstrumentClassSecTypeValid, True
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get SecType() As SecurityTypes
SecType = gCategoryToSecType(mDataObj.GetColumnValue(InstrumentClassColumnSecTypeId, 0))
End Property

Public Property Let SecTypeString(ByVal value As String)
Const ProcName As String = "SecTypeString"
Dim failpoint As String
On Error GoTo Err

value = Trim$(value)
If mDataObj.RuleBroken(BusRuleInstrumentClassSecTypeValid, _
                    (SecTypeFromString(value) = SecTypeNone)) Then
    SecType = SecTypeNone
Else
    SecType = SecTypeFromString(value)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get SecTypeString() As String
Const ProcName As String = "SecTypeString"
Dim failpoint As String
On Error GoTo Err

SecTypeString = SecTypeToString(SecType)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get SessionEndTime() As Date
Const ProcName As String = "SessionEndTime"
Dim failpoint As String
On Error GoTo Err

SessionEndTime = mDataObj.GetColumnValue(InstrumentClassColumnSessionEndTime, CDate(0))
SessionEndTime = SessionEndTime - Int(SessionEndTime)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let SessionEndTime( _
                ByVal value As Date)
Const ProcName As String = "SessionEndTime"
Dim failpoint As String
On Error GoTo Err

mDataObj.RuleBroken BusRuleInstrumentClassSessionEndTimeValid, value >= 1
mDataObj.ColumnValue(InstrumentClassColumnSessionEndTime) = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let SessionEndTimeString( _
                ByVal value As String)
Const ProcName As String = "SessionEndTimeString"
Dim failpoint As String
On Error GoTo Err

If Not mDataObj.RuleBroken(BusRuleInstrumentClassSessionEndTimeValid, Not IsDate(value)) Then
    SessionEndTime = CDate(value)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get SessionStartTime() As Date
SessionStartTime = mDataObj.GetColumnValue(InstrumentClassColumnSessionStartTime, CDate(0))
SessionStartTime = SessionStartTime - Int(SessionStartTime)
End Property

Public Property Let SessionStartTime( _
                ByVal value As Date)
Const ProcName As String = "SessionStartTime"
Dim failpoint As String
On Error GoTo Err

mDataObj.RuleBroken BusRuleInstrumentClassSessionStartTimeValid, value >= 1
mDataObj.ColumnValue(InstrumentClassColumnSessionStartTime) = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let SessionStartTimeString( _
                ByVal value As String)
If Not mDataObj.RuleBroken(BusRuleInstrumentClassSessionStartTimeValid, Not IsDate(value)) Then
    SessionStartTime = CDate(value)
End If
End Property

Public Property Let TickSize( _
                ByVal value As Double)
Const ProcName As String = "TickSize"
Dim failpoint As String
On Error GoTo Err

mDataObj.RuleBroken BusRuleInstrumentClassTickSizeValid, value <= 0
mDataObj.ColumnValue(InstrumentClassColumnTickSize) = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TickSize() As Double
TickSize = mDataObj.GetColumnValue(InstrumentClassColumnTickSize, 0)
End Property

Public Property Let TickSizeString( _
                ByVal value As String)
Const ProcName As String = "TickSizeString"
Dim failpoint As String
On Error GoTo Err

If mDataObj.RuleBroken(BusRuleInstrumentClassTickSizeValid, Not IsNumeric(value)) Then
    TickSize = 999.999
Else
    TickSize = CDbl(value)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let TickValue(ByVal value As Double)
Const ProcName As String = "TickValue"
Dim failpoint As String
On Error GoTo Err

mDataObj.RuleBroken BusRuleInstrumentClassTickValueValid, value <= 0
mDataObj.ColumnValue(InstrumentClassColumnTickValue) = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TickValue() As Double
Const ProcName As String = "TickValue"
Dim failpoint As String
On Error GoTo Err

TickValue = mDataObj.GetColumnValue(InstrumentClassColumnTickValue, 0)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let TickValueString( _
                ByVal value As String)
Const ProcName As String = "TickValueString"
Dim failpoint As String
On Error GoTo Err

If mDataObj.RuleBroken(BusRuleInstrumentClassTickValueValid, Not IsNumeric(value)) Then
    TickValue = 999.999
Else
    TickValue = CDbl(value)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TimeZoneName() As String
Const ProcName As String = "TimeZoneName"
Dim failpoint As String
On Error GoTo Err

If Not mDataObj.IsNew Then
    TimeZoneName = mDataObj.GetColumnValue(InstrumentClassColumnTimeZone, "")
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                Optional ByVal rs As ADODB.Recordset)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

GenericInitialise db, rs

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getBusRuleNameValidId() As Long
getBusRuleNameValidId = BusRuleInstrumentClassNameValid
End Function

Private Function getFactory() As DataObjectFactory
Const ProcName As String = "getFactory"
Dim failpoint As String
On Error GoTo Err

Set getFactory = mDB.InstrumentClassFactory

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub setInitialErrors()
' set errors for fields that are invalid when a new record is created
' Note that this is done automatically by the AbstractDataObject for
' the Name field
Const ProcName As String = "setInitialErrors"
Dim failpoint As String
On Error GoTo Err

mDataObj.RuleBroken BusRuleInstrumentClassExchangeValid, True
mDataObj.RuleBroken BusRuleInstrumentClassSecTypeValid, True
mDataObj.RuleBroken BusRuleInstrumentClassCurrencyCodeValid, True
mDataObj.RuleBroken BusRuleInstrumentClassTickSizeValid, True
mDataObj.RuleBroken BusRuleInstrumentClassTickValueValid, True
mDataObj.RuleBroken BusRuleInstrumentClassSessionStartTimeValid, True
mDataObj.RuleBroken BusRuleInstrumentClassSessionEndTimeValid, True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "BusinessDataObject_ApplyEdit"
Dim failpoint As String
On Error GoTo Err

ApplyEdit

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub BusinessDataObject_CancelEdit()
Const ProcName As String = "BusinessDataObject_CancelEdit"
Dim failpoint As String
On Error GoTo Err

CancelEdit

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub BusinessDataObject_Delete()
Const ProcName As String = "BusinessDataObject_Delete"
Dim failpoint As String
On Error GoTo Err

Delete

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get BusinessDataObject_ErrorList() As ErrorList
Set BusinessDataObject_ErrorList = mDataObj.ErrorList
End Property

Private Property Get BusinessDataObject_EventSource() As BusinessObjEventSource
Const ProcName As String = "BusinessDataObject_EventSource"
Dim failpoint As String
On Error GoTo Err

Set BusinessDataObject_EventSource = mDataObj.eventSource

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BusinessDataObject_ID() As Long
Const ProcName As String = "BusinessDataObject_ID"
Dim failpoint As String
On Error GoTo Err

BusinessDataObject_ID = Id

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BusinessDataObject_IsDeleted() As Boolean
Const ProcName As String = "BusinessDataObject_IsDeleted"
Dim failpoint As String
On Error GoTo Err

BusinessDataObject_IsDeleted = IsDeleted

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BusinessDataObject_IsDirty() As Boolean
Const ProcName As String = "BusinessDataObject_IsDirty"
Dim failpoint As String
On Error GoTo Err

BusinessDataObject_IsDirty = IsDirty

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BusinessDataObject_IsNew() As Boolean
Const ProcName As String = "BusinessDataObject_IsNew"
Dim failpoint As String
On Error GoTo Err

BusinessDataObject_IsNew = IsNew

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BusinessDataObject_IsValid() As Boolean
Const ProcName As String = "BusinessDataObject_IsValid"
Dim failpoint As String
On Error GoTo Err

BusinessDataObject_IsValid = IsValid

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Let BusinessDataObject_Name(ByVal RHS As String)
Const ProcName As String = "BusinessDataObject_Name"
Dim failpoint As String
On Error GoTo Err

Name = RHS

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BusinessDataObject_Name() As String
Const ProcName As String = "BusinessDataObject_Name"
Dim failpoint As String
On Error GoTo Err

BusinessDataObject_Name = Name

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub BusinessDataObject_Refresh()
Const ProcName As String = "BusinessDataObject_Refresh"
Dim failpoint As String
On Error GoTo Err

Refresh

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "ErrorList"
Dim failpoint As String
On Error GoTo Err

Set ErrorList = mDataObj.ErrorList

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Id() As Long
Const ProcName As String = "Id"
Dim failpoint As String
On Error GoTo Err

Id = mDataObj.Id

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IsValid() As Boolean
Const ProcName As String = "IsValid"
Dim failpoint As String
On Error GoTo Err

IsValid = mDataObj.IsValid

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IsDeleted() As Boolean
Const ProcName As String = "IsDeleted"
Dim failpoint As String
On Error GoTo Err

IsDeleted = mDataObj.IsDeleted

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IsDirty() As Boolean
Const ProcName As String = "IsDirty"
Dim failpoint As String
On Error GoTo Err

IsDirty = mDataObj.IsDirty

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IsNew() As Boolean
Const ProcName As String = "IsNew"
Dim failpoint As String
On Error GoTo Err

IsNew = mDataObj.IsNew

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Generic Methods
'@================================================================================

Public Sub ApplyEdit()
Const ProcName As String = "ApplyEdit"
Dim failpoint As String
On Error GoTo Err

mDataObj.ApplyEdit

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub CancelEdit()
Const ProcName As String = "CancelEdit"
Dim failpoint As String
On Error GoTo Err

mDataObj.CancelEdit
If mDataObj.IsNew Then setInitialErrors

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub Delete()
Const ProcName As String = "Delete"
Dim failpoint As String
On Error GoTo Err

mDataObj.Delete

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub GenericInitialise( _
                ByVal db As Variant, _
                Optional ByVal rs As ADODB.Recordset)
Const ProcName As String = "GenericInitialise"
Dim failpoint As String
On Error GoTo Err

Set mDB = db
If Not rs Is Nothing Then
    Set mDataObj = CreateAbstractDataObject(Me, getBusRuleNameValidId, getFactory, mDB.ConnectionString)
    mDataObj.InitialiseRs rs
    If mDataObj.IsNew Then setInitialErrors
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function GetName() As String
Const ProcName As String = "GetName"
Dim failpoint As String
On Error GoTo Err

GetName = mDataObj.Name

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub Refresh()
Const ProcName As String = "Refresh"
Dim failpoint As String
On Error GoTo Err

mDataObj.Refresh

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub SetName(ByVal value As String)
Const ProcName As String = "SetName"
Dim failpoint As String
On Error GoTo Err

mDataObj.Name = value

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub


