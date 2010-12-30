VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Exchange"
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

Implements BusinessDataObject

'@================================================================================
' Events
'@================================================================================

Public Event Dirty()
Public Event Clean()
Public Event Deleted()
Public Event Invalid()
Public Event Valid()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                            As String = "Exchange"

'@================================================================================
' Member variables
'@================================================================================

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

'@===============================================================================
' XXXX Event Handlers
'@===============================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Name() As String
Name = GetName
End Property

Public Property Let Name(ByVal value As String)
SetName value
End Property

Public Property Get Notes() As String
Const ProcName As String = "Notes"
On Error GoTo Err

Notes = mDataObj.GetColumnValue(ExchangeColumnNotes, "")

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Notes(ByVal val As String)
Const ProcName As String = "Notes"
On Error GoTo Err

mDataObj.ColumnValue(ExchangeColumnNotes) = val

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let TimeZone( _
                ByVal value As TimeZone)
Const ProcName As String = "TimeZone"
On Error GoTo Err

mDataObj.RuleBroken BusRuleExchangeTimezoneValid, value Is Nothing
mDataObj.ColumnValue(ExchangeColumnTimeZoneID) = value.Id

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get TimeZone() As TimeZone
Const ProcName As String = "TimeZone"
On Error GoTo Err

If mDataObj.GetColumnValue(ExchangeColumnTimeZoneID, 0) = 0 Then Exit Property
Set TimeZone = mDB.TimeZoneFactory.LoadByID(mDataObj.GetColumnValue(ExchangeColumnTimeZoneID, 0))

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let TimeZoneName( _
                ByVal value As String)
Dim tz As TimeZone

Const ProcName As String = "TimeZoneName"
On Error GoTo Err

Set tz = mDB.TimeZoneFactory.LoadByName(value)
If tz Is Nothing Then
    mDataObj.RuleBroken BusRuleExchangeTimezoneValid, True
    If Not mDataObj.IsNew Then
        ' this column is not available for a new object
        mDataObj.ColumnValue(ExchangeColumnTimeZoneName) = Left$(value, mDataObj.ColumnMaxLength(ExchangeColumnTimeZoneName))
    End If
Else
    TimeZone = tz
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get TimeZoneName() As String
Const ProcName As String = "TimeZoneName"
On Error GoTo Err

If Not mDataObj.IsNew Then
    TimeZoneName = mDataObj.GetColumnValue(ExchangeColumnTimeZoneName, "")
ElseIf CLng(mDataObj.GetColumnValue(ExchangeColumnTimeZoneID, 0)) <> 0 Then
    TimeZoneName = mDB.TimeZoneFactory.LoadByID(mDataObj.GetColumnValue(ExchangeColumnTimeZoneID, 0)).Name
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                Optional ByVal rs As ADODB.Recordset)
Const ProcName As String = "Initialise"
On Error GoTo Err

GenericInitialise db, rs

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getBusRuleNameValidId() As Long
Const ProcName As String = "getBusRuleNameValidId"
On Error GoTo Err

getBusRuleNameValidId = BusRuleExchangeNameValid

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getFactory() As DataObjectFactory
Const ProcName As String = "getFactory"
On Error GoTo Err

Set getFactory = mDB.ExchangeFactory

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub setInitialErrors()
' set errors for fields that are invalid when a new record is created
' Note that this is done automatically by the AbstractDataObject for
' the Name field
Const ProcName As String = "setInitialErrors"
On Error GoTo Err

mDataObj.RuleBroken BusRuleExchangeTimezoneValid, True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
On Error GoTo Err

ApplyEdit

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub BusinessDataObject_CancelEdit()
Const ProcName As String = "BusinessDataObject_CancelEdit"
On Error GoTo Err

CancelEdit

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub BusinessDataObject_Delete()
Const ProcName As String = "BusinessDataObject_Delete"
On Error GoTo Err

Delete

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get BusinessDataObject_ErrorList() As ErrorList
Const ProcName As String = "BusinessDataObject_ErrorList"
On Error GoTo Err

Set BusinessDataObject_ErrorList = mDataObj.ErrorList

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get BusinessDataObject_EventSource() As BusinessObjEventSource
Const ProcName As String = "BusinessDataObject_EventSource"
On Error GoTo Err

Set BusinessDataObject_EventSource = mDataObj.EventSource

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get BusinessDataObject_ID() As Long
Const ProcName As String = "BusinessDataObject_ID"
On Error GoTo Err

BusinessDataObject_ID = Id

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get BusinessDataObject_IsDeleted() As Boolean
Const ProcName As String = "BusinessDataObject_IsDeleted"
On Error GoTo Err

BusinessDataObject_IsDeleted = IsDeleted

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get BusinessDataObject_IsDirty() As Boolean
Const ProcName As String = "BusinessDataObject_IsDirty"
On Error GoTo Err

BusinessDataObject_IsDirty = IsDirty

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get BusinessDataObject_IsNew() As Boolean
Const ProcName As String = "BusinessDataObject_IsNew"
On Error GoTo Err

BusinessDataObject_IsNew = IsNew

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get BusinessDataObject_IsValid() As Boolean
Const ProcName As String = "BusinessDataObject_IsValid"
On Error GoTo Err

BusinessDataObject_IsValid = IsValid

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Let BusinessDataObject_Name(ByVal RHS As String)
Const ProcName As String = "BusinessDataObject_Name"
On Error GoTo Err

Name = RHS

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get BusinessDataObject_Name() As String
Const ProcName As String = "BusinessDataObject_Name"
On Error GoTo Err

BusinessDataObject_Name = Name

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub BusinessDataObject_Refresh()
Const ProcName As String = "BusinessDataObject_Refresh"
On Error GoTo Err

Refresh

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@===============================================================================
' mDataObj Event Handlers
'@===============================================================================

Private Sub mDataObj_Clean()
Const ProcName As String = "mDataObj_Clean"
On Error GoTo Err

RaiseEvent Clean

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mDataObj_Deleted()
Const ProcName As String = "mDataObj_Deleted"
On Error GoTo Err

RaiseEvent Deleted

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mDataObj_Dirty()
Const ProcName As String = "mDataObj_Dirty"
On Error GoTo Err

RaiseEvent Dirty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mDataObj_Invalid()
Const ProcName As String = "mDataObj_Invalid"
On Error GoTo Err

RaiseEvent Invalid

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mDataObj_Valid()
Const ProcName As String = "mDataObj_Valid"
On Error GoTo Err

RaiseEvent Valid

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Generic Properties
'@================================================================================

Public Property Get ErrorList() As ErrorList
Const ProcName As String = "ErrorList"
On Error GoTo Err

Set ErrorList = mDataObj.ErrorList

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Id() As Long
Const ProcName As String = "Id"
On Error GoTo Err

Id = mDataObj.Id

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsValid() As Boolean
Const ProcName As String = "IsValid"
On Error GoTo Err

IsValid = mDataObj.IsValid

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsDeleted() As Boolean
Const ProcName As String = "IsDeleted"
On Error GoTo Err

IsDeleted = mDataObj.IsDeleted

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsDirty() As Boolean
Const ProcName As String = "IsDirty"
On Error GoTo Err

IsDirty = mDataObj.IsDirty

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsNew() As Boolean
Const ProcName As String = "IsNew"
On Error GoTo Err

IsNew = mDataObj.IsNew

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Generic Methods
'@================================================================================

Public Sub ApplyEdit()
Const ProcName As String = "ApplyEdit"
On Error GoTo Err

mDataObj.ApplyEdit

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub CancelEdit()
Const ProcName As String = "CancelEdit"
On Error GoTo Err

mDataObj.CancelEdit
If mDataObj.IsNew Then setInitialErrors

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Delete()
Const ProcName As String = "Delete"
On Error GoTo Err

mDataObj.Delete

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub GenericInitialise( _
                ByVal db As Variant, _
                Optional ByVal rs As ADODB.Recordset)
Const ProcName As String = "GenericInitialise"
On Error GoTo Err

Set mDB = db
If Not rs Is Nothing Then
    Set mDataObj = CreateAbstractDataObject(Me, getBusRuleNameValidId, getFactory, mDB.ConnectionString)
    mDataObj.InitialiseRs rs
    If mDataObj.IsNew Then setInitialErrors
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function GetName() As String
Const ProcName As String = "GetName"
On Error GoTo Err

GetName = mDataObj.Name

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub Refresh()
Const ProcName As String = "Refresh"
On Error GoTo Err

mDataObj.Refresh

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub SetName(ByVal value As String)
Const ProcName As String = "SetName"
On Error GoTo Err

mDataObj.Name = value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub




