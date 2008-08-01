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

Private Const ProjectName                           As String = "TradingDO26"
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

Public Property Get name() As String
name = getName
End Property

Public Property Let name(ByVal value As String)
setName value
End Property

Public Property Get notes() As String
notes = mDataObj.getColumnValue(ExchangeColumnNotes, "")
End Property

Public Property Let notes(ByVal val As String)
mDataObj.columnValue(ExchangeColumnNotes) = val
End Property

Public Property Let TimeZone( _
                ByVal value As TimeZone)
mDataObj.ruleBroken BusRuleExchangeTimezoneValid, value Is Nothing

If Not value Is Nothing Then mDataObj.columnValue(ExchangeColumnTimeZoneID) = value.id
'If Not mDataObj.IsNew Then
'    mDataObj.columnValue(ExchangeColumnTimeZoneName) = value.canonicalName
'End If
End Property

Public Property Get TimeZone() As TimeZone
If mDataObj.getColumnValue(ExchangeColumnTimeZoneID, 0) = 0 Then Exit Property
Set TimeZone = mDB.TimeZoneFactory.loadByID(mDataObj.getColumnValue(ExchangeColumnTimeZoneID, 0))
End Property

Public Property Let TimeZoneName( _
                ByVal value As String)
Dim tz As TimeZone
Set tz = mDB.TimeZoneFactory.loadByName(value)
If tz Is Nothing Then
    mDataObj.ruleBroken BusRuleExchangeTimezoneValid, True
    If Not mDataObj.IsNew Then
        ' this column is not available for a new object
        mDataObj.columnValue(ExchangeColumnTimeZoneName) = Left$(value, mDataObj.columnMaxLength(ExchangeColumnTimeZoneName))
    End If
Else
    TimeZone = tz
End If
End Property

Public Property Get TimeZoneName() As String
If Not mDataObj.IsNew Then
    TimeZoneName = mDataObj.getColumnValue(ExchangeColumnTimeZoneName, "")
ElseIf CLng(mDataObj.getColumnValue(ExchangeColumnTimeZoneID, 0)) <> 0 Then
    TimeZoneName = mDB.TimeZoneFactory.loadByID(mDataObj.getColumnValue(ExchangeColumnTimeZoneID, 0)).name
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
getBusRuleNameValidId = BusRuleExchangeNameValid
End Function

Private Function getFactory() As DataObjectFactory
Set getFactory = mDB.ExchangeFactory
End Function

Private Sub setInitialErrors()
' set errors for fields that are invalid when a new record is created
' Note that this is done automatically by the AbstractDataObject for
' the name field
mDataObj.ruleBroken BusRuleExchangeTimezoneValid, True
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




