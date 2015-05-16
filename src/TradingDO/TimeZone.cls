VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeZone"
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

Private Const ModuleName                            As String = "TimeZone"

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

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get canonicalID() As Long
Const ProcName As String = "canonicalID"

On Error GoTo Err

canonicalID = mDataObj.GetColumnValue(TimeZoneColumnCanonicalId, 0)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get canonicalName() As String
Const ProcName As String = "canonicalName"

On Error GoTo Err

If mDataObj.IsNew Then
    canonicalName = ""
Else
    canonicalName = mDataObj.GetColumnValue(TimeZoneColumnCanonicalName)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let canonicalName(ByVal value As String)
Dim lTimeZone As TimeZone

Const ProcName As String = "canonicalName"

On Error GoTo Err

value = Trim$(value)
mDataObj.RuleBroken BusRuleTimeZoneCanonicalNameValid, _
                    Len(value) = 0 Or Len(value) > mDataObj.ColumnMaxLength(TimeZoneColumnName)

Set lTimeZone = mDB.TimeZoneFactory.LoadByName(value)

mDataObj.RuleBroken BusRuleTimeZoneCanonicalNameValid, _
                    lTimeZone Is Nothing

If Not lTimeZone Is Nothing Then
    mDataObj.RuleBroken BusRuleTimeZoneCanonicalNameValid, _
                        lTimeZone.canonicalID <> 0
    mDataObj.ColumnValue(TimeZoneColumnCanonicalId) = lTimeZone.Id
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Name() As String
Const ProcName As String = "Name"

On Error GoTo Err

Name = GetName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Name(ByVal value As String)
Const ProcName As String = "Name"

On Error GoTo Err

SetName value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getBusRuleNameValidId() As Long
getBusRuleNameValidId = BusRuleTimeZoneNameValid
End Function

Private Function getFactory() As DataObjectFactory
Const ProcName As String = "getFactory"

On Error GoTo Err

Set getFactory = mDB.TimeZoneFactory

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setInitialErrors()
' set errors for fields that are invalid when a new record is created
' Note that this is done automatically by the AbstractDataObject for
' the Name field
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub BusinessDataObject_CancelEdit()
Const ProcName As String = "BusinessDataObject_CancelEdit"

On Error GoTo Err

CancelEdit

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub BusinessDataObject_Delete()
Const ProcName As String = "BusinessDataObject_Delete"

On Error GoTo Err

Delete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get BusinessDataObject_ErrorList() As ErrorList
Const ProcName As String = "BusinessDataObject_ErrorList"

On Error GoTo Err

Set BusinessDataObject_ErrorList = mDataObj.ErrorList

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get BusinessDataObject_EventSource() As BusinessObjEventSource
Const ProcName As String = "BusinessDataObject_EventSource"

On Error GoTo Err

Set BusinessDataObject_EventSource = mDataObj.EventSource

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get BusinessDataObject_ID() As Long
Const ProcName As String = "BusinessDataObject_ID"

On Error GoTo Err

BusinessDataObject_ID = Id

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get BusinessDataObject_IsDeleted() As Boolean
Const ProcName As String = "BusinessDataObject_IsDeleted"

On Error GoTo Err

BusinessDataObject_IsDeleted = IsDeleted

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get BusinessDataObject_IsDirty() As Boolean
Const ProcName As String = "BusinessDataObject_IsDirty"

On Error GoTo Err

BusinessDataObject_IsDirty = IsDirty

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get BusinessDataObject_IsNew() As Boolean
Const ProcName As String = "BusinessDataObject_IsNew"

On Error GoTo Err

BusinessDataObject_IsNew = IsNew

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get BusinessDataObject_IsValid() As Boolean
Const ProcName As String = "BusinessDataObject_IsValid"

On Error GoTo Err

BusinessDataObject_IsValid = IsValid

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Let BusinessDataObject_Name(ByVal RHS As String)
Const ProcName As String = "BusinessDataObject_Name"

On Error GoTo Err

Name = RHS

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get BusinessDataObject_Name() As String
Const ProcName As String = "BusinessDataObject_Name"

On Error GoTo Err

BusinessDataObject_Name = Name

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub BusinessDataObject_Refresh()
Const ProcName As String = "BusinessDataObject_Refresh"

On Error GoTo Err

Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

On Error GoTo Err

Set ErrorList = mDataObj.ErrorList

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Id() As Long
Const ProcName As String = "Id"

On Error GoTo Err

Id = mDataObj.Id

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsValid() As Boolean
Const ProcName As String = "IsValid"

On Error GoTo Err

IsValid = mDataObj.IsValid

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsDeleted() As Boolean
Const ProcName As String = "IsDeleted"

On Error GoTo Err

IsDeleted = mDataObj.IsDeleted

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsDirty() As Boolean
Const ProcName As String = "IsDirty"

On Error GoTo Err

IsDirty = mDataObj.IsDirty

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsNew() As Boolean
Const ProcName As String = "IsNew"

On Error GoTo Err

IsNew = mDataObj.IsNew

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub CancelEdit()
Const ProcName As String = "CancelEdit"

On Error GoTo Err

mDataObj.CancelEdit
If mDataObj.IsNew Then setInitialErrors

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Delete()
Const ProcName As String = "Delete"

On Error GoTo Err

mDataObj.Delete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetName() As String
Const ProcName As String = "GetName"

On Error GoTo Err

GetName = mDataObj.Name

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Refresh()
Const ProcName As String = "Refresh"

On Error GoTo Err

mDataObj.Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetName(ByVal value As String)
Const ProcName As String = "SetName"

On Error GoTo Err

mDataObj.Name = value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub








