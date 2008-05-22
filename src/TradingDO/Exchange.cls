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

Private Const ProjectName As String = "TradingDO26"
Private Const ModuleName As String = "Exchange"

'@================================================================================
' Member variables
'@================================================================================

Private mDB As TradingDB

Private mConn As ADODB.Connection

Private mRs As ADODB.Recordset
Private mNew As Boolean

Private mDeleted As Boolean
Private mDirty As Boolean


Private WithEvents mErrorList As ErrorList
Attribute mErrorList.VB_VarHelpID = -1

Private mEventRaiser As BusinessObjEventRaiser

Private mMaxNameLength As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mErrorList = New ErrorList
Set mEventRaiser = New BusinessObjEventRaiser
End Sub

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
Set BusinessDataObject_ErrorList = mErrorList
End Property

Private Property Get BusinessDataObject_EventSource() As BusinessObjEventSource
Set BusinessDataObject_EventSource = mEventRaiser.EventSource
End Property

Private Property Get BusinessDataObject_ID() As Long
BusinessDataObject_ID = id
End Property

Private Sub BusinessDataObject_InitialiseFromRecordset(ByVal rs As ADODB.Recordset)
initialiseRs rs
End Sub

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

'@===============================================================================
' mErrorList Event Handlers
'@===============================================================================

Private Sub mErrorList_ItemAdded( _
                ByVal item As ErrorItem)
RaiseEvent Invalid
mEventRaiser.Invalid
End Sub

Private Sub mErrorList_NoItems()
RaiseEvent Valid
mEventRaiser.Valid
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get ErrorList() As ErrorList
Set ErrorList = mErrorList
End Property

Public Property Get id() As Long
id = mRs("ID")
End Property

Public Property Get IsValid() As Boolean
    IsValid = (mErrorList.count = 0)
End Property


Public Property Get IsDeleted() As Boolean
  IsDeleted = mDeleted
End Property

Public Property Get IsDirty() As Boolean
IsDirty = mDirty
End Property

Public Property Get IsNew() As Boolean
  IsNew = mNew
End Property

Public Property Get name() As String
name = Trim$(mRs(ExchangeColumnName))
End Property

Public Property Let name(ByVal value As String)
value = Trim$(value)
ruleBroken BusRuleExchangeNameValid, _
    Len(value) = 0 Or Len(value) > mMaxNameLength
mRs("NAME") = Left$(value, mMaxNameLength)
Dirty
End Property

Public Property Get notes() As String
notes = Nz(mRs(ExchangeColumnNotes), "")
End Property

Public Property Let notes(ByVal val As String)
mRs(ExchangeColumnNotes) = val
Dirty
End Property

Public Property Let TimeZone( _
                ByVal value As TimeZone)
ruleBroken BusRuleExchangeTimezoneValid, value Is Nothing

If Not value Is Nothing Then mRs("timezoneid") = value.id
Dirty
End Property

Public Property Get TimeZone() As TimeZone
If IsNull(mRs("timezoneid")) Then Exit Property
Set TimeZone = mDB.TimeZoneFactory.loadByID(CLng(mRs("timezoneid")))
End Property

Public Property Get timeZoneCanonicalName() As String
If Not TimeZone Is Nothing Then timeZoneCanonicalName = TimeZone.canonicalName
End Property

Public Property Let TimeZoneName( _
                ByVal value As String)
Dim tz As TimeZone
Set tz = mDB.TimeZoneFactory.loadByName(value)
If tz Is Nothing Then
    ruleBroken BusRuleExchangeTimezoneValid, True
    Dirty
Else
    TimeZone = tz
End If
End Property

Public Property Get TimeZoneName() As String
If TimeZone Is Nothing Then Exit Property
TimeZoneName = TimeZone.name
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ApplyEdit()
If mDeleted Then
    mRs.Delete
    save
    Set mRs = Nothing
    RaiseEvent Deleted
    mEventRaiser.Deleted
ElseIf mDirty Or mNew Then
    If Not IsValid Then Err.Raise 445
    save
    mNew = False
End If
Clean
End Sub

Public Sub CancelEdit()
mRs.CancelUpdate
mDeleted = False
If mRs.EOF Then
    mRs.AddNew
    mNew = True
    setInitialErrors
End If
Clean
End Sub

Public Sub Delete()
mDeleted = True
End Sub

Friend Sub initialise( _
                ByVal db As TradingDB, _
                Optional ByVal rs As ADODB.Recordset)
Set mDB = db

If Not rs Is Nothing Then initialiseRs rs
End Sub

Friend Sub initialiseRs( _
                ByVal rs As ADODB.Recordset)
Set mRs = rs

mMaxNameLength = mRs.Fields(ExchangeColumnName).DefinedSize

If mRs.EOF Then
    mRs.AddNew
    mNew = True
    setInitialErrors
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub Clean()
mDirty = False
RaiseEvent Clean
mEventRaiser.Clean
mErrorList.Clear
If mNew Then
    setInitialErrors
End If
End Sub

Private Sub Dirty()
If Not mDirty Then
    mDirty = True
    RaiseEvent Dirty
    mEventRaiser.Dirty
End If
End Sub

Private Function ruleBroken( _
                ByVal ruleId As BusinessRuleIds, _
                ByVal isBroken As Boolean) As Boolean
ruleBroken = mErrorList.AssertRule(Not isBroken, ruleId)
End Function

Private Sub save()
Dim failpoint As Long
On Error GoTo Err

failpoint = 100

Set mConn = New ADODB.Connection
mConn.Open mDB.connectionString

failpoint = 200

Set mRs.ActiveConnection = mConn
mRs.UpdateBatch
Set mRs.ActiveConnection = Nothing
mConn.Close

Exit Sub

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "Exchange" & "::" & "save" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Sub

Private Sub setInitialErrors()
' sets errors for fields that are invalid when a new record is created
ruleBroken BusRuleExchangeNameValid, True
ruleBroken BusRuleExchangeTimezoneValid, True
End Sub



