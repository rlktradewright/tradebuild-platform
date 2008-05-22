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

Private mDB As TradingDB

Private mConn As ADODB.Connection

Private mNew As Boolean

Private mDeleted As Boolean
Private mDirty As Boolean

Private WithEvents mErrorList As ErrorList
Attribute mErrorList.VB_VarHelpID = -1

Private mRs As ADODB.Recordset

Private mEventRaiser As BusinessObjEventRaiser

Private mMaxNameLength As Long
Private mMaxCurrencyCodeLength As Long

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
Set mErrorList = New ErrorList
Set mEventRaiser = New BusinessObjEventRaiser
End Sub

'@===============================================================================
' BusinessDataObject Interface Members
'@===============================================================================

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

'@===============================================================================
' Properties
'@===============================================================================

Public Property Get currencyCode() As String
currencyCode = mRs!Currency
End Property

Public Property Let currencyCode(ByVal value As String)
value = Trim$(value)
ruleBroken BusRuleInstrumentClassCurrencyCodeValid, _
    Len(value) = 0 Or Len(value) > mMaxCurrencyCodeLength Or Not IsValidCurrencyCode(value)
mRs!Currency = Left$(value, mMaxCurrencyCodeLength)
Dirty
End Property

Public Property Get daysBeforeExpiryToSwitch() As Long
daysBeforeExpiryToSwitch = Nz(mRs("DaysBeforeExpiryToSwitch"), 0)
End Property

Public Property Let daysBeforeExpiryToSwitch(ByVal value As Long)
ruleBroken BusRuleInstrumentClassDaysBeforeExpiryValid, _
    value < 0
mRs("DaysBeforeExpiryToSwitch") = value
Dirty
End Property

Public Property Let daysBeforeExpiryToSwitchString(ByVal value As String)
If value = "" Then Exit Property
If IsInteger(value) Then
    daysBeforeExpiryToSwitch = CLng(value)
Else
    ruleBroken BusRuleInstrumentClassDaysBeforeExpiryValid, True
    Dirty
End If
End Property

Public Property Get ErrorList() As ErrorList
Set ErrorList = mErrorList
End Property

Public Property Let Exchange( _
                ByVal value As Exchange)
ruleBroken BusRuleInstrumentClassExchangeValid, _
    value Is Nothing
If Not value Is Nothing Then
    mRs("exchangeid") = value.id
End If
Dirty
End Property

Public Property Get Exchange() As Exchange
Set Exchange = mDB.ExchangeFactory.loadByID(mRs("exchangeid"))
End Property

Public Property Let exchangeName( _
                ByVal value As String)
Exchange = mDB.ExchangeFactory.loadByName(value)
End Property

Public Property Get exchangeName() As String
exchangeName = Trim$(mRs(InstrumentClassColumnExchange))
End Property

Public Property Get id() As Long
id = mRs("ID")
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

Public Property Get IsValid() As Boolean
IsValid = (mErrorList.count = 0)
End Property

Public Property Get name() As String
name = Trim$(mRs("NAME"))
End Property

Public Property Let name(ByVal value As String)
value = Trim$(value)
ruleBroken BusRuleInstrumentClassNameValid, _
    Len(value) = 0 Or Len(value) > mMaxNameLength
mRs("NAME") = Left$(value, mMaxNameLength)
Dirty
End Property

Public Property Get notes() As String
notes = Nz(mRs("notes"), "")
End Property

Public Property Let notes(ByVal val As String)
mRs("notes") = val
Dirty
End Property

Public Property Let secType(ByVal value As SecurityTypes)
Dim catid As InstrumentCategories

If IsValidSecType(value) Then
    ruleBroken BusRuleInstrumentClassSecTypeValid, False
    catid = gCategoryFromSecType(value)
    mRs("instrumentcategoryid") = catid
Else
    ruleBroken BusRuleInstrumentClassSecTypeValid, True
End If

Dirty
End Property

Public Property Get secType() As SecurityTypes
secType = gCategoryToSecType(mRs("instrumentcategoryid"))
End Property

Public Property Let secTypeString(ByVal value As String)
value = Trim$(value)
secType = SecTypeFromString(value)
End Property

Public Property Get secTypeString() As String
secTypeString = SecTypeToString(secType)
End Property

Public Property Get sessionEndTime() As Date
sessionEndTime = Nz(mRs("sessionEndTime"), CDate(0))
sessionEndTime = sessionEndTime - Int(sessionEndTime)
End Property

Public Property Let sessionEndTime( _
                ByVal value As Date)
ruleBroken BusRuleInstrumentClassSessionEndTimeValid, _
    value >= 1
mRs("sessionEndTime") = value
Dirty
End Property

Public Property Let sessionEndTimeString( _
                ByVal value As String)
If Not IsDate(value) Then
    ruleBroken BusRuleInstrumentClassSessionEndTimeValid, True
    Dirty
Else
    sessionEndTime = CDate(value)
End If
End Property

Public Property Get sessionStartTime() As Date
sessionStartTime = Nz(mRs("sessionStartTime"), CDate(0))
sessionStartTime = sessionStartTime - Int(sessionStartTime)
End Property

Public Property Let sessionStartTime( _
                ByVal value As Date)
ruleBroken BusRuleInstrumentClassSessionStartTimeValid, _
    value >= 1
mRs("sessionStartTime") = value
Dirty
End Property

Public Property Let sessionStartTimeString( _
                ByVal value As String)
If Not IsDate(value) Then
    ruleBroken BusRuleInstrumentClassSessionStartTimeValid, True
    Dirty
Else
    sessionStartTime = CDate(value)
End If
End Property

Public Property Let tickSize( _
                ByVal value As Double)
ruleBroken BusRuleInstrumentClassTickSizeValid, _
    value <= 0
mRs("TickSize") = value
Dirty
End Property

Public Property Get tickSize() As Double
tickSize = mRs("TickSize")
End Property

Public Property Let tickSizeString( _
                ByVal value As String)
If Not IsNumeric(value) Then
    ruleBroken BusRuleInstrumentClassTickSizeValid, True
    Dirty
Else
    tickSize = CDbl(value)
End If
End Property

Public Property Let tickValue(ByVal value As Double)
ruleBroken BusRuleInstrumentClassTickValueValid, _
    value <= 0
mRs("TickValue") = value
Dirty
End Property

Public Property Get tickValue() As Double
tickValue = mRs("TickValue")
End Property

Public Property Let tickValueString( _
                ByVal value As String)
If Not IsNumeric(value) Then
    ruleBroken BusRuleInstrumentClassTickValueValid, True
    Dirty
Else
    tickValue = CDbl(value)
End If
End Property

Public Property Get timeZoneCanonicalName() As String
timeZoneCanonicalName = Trim$(mRs(InstrumentClassColumnTimeZone))
End Property

Public Property Get TimeZoneName() As String
TimeZoneName = Trim$(mRs("TimeZone"))
End Property

'@===============================================================================
' Methods
'@===============================================================================

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

mMaxNameLength = mRs.Fields(InstrumentClassColumnName).DefinedSize
mMaxCurrencyCodeLength = mRs.Fields(InstrumentClassColumnCurrency).DefinedSize

If mRs.EOF Then
    mRs.AddNew
    mNew = True
    setInitialErrors
End If
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

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
        "TradingDO26" & "." & "Instrument" & "::" & "save" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Sub

Private Sub setInitialErrors()
' sets errors for fields that are invalid when a new record is created
ruleBroken BusRuleInstrumentClassNameValid, True
ruleBroken BusRuleInstrumentClassExchangeValid, True
ruleBroken BusRuleInstrumentClassSecTypeValid, True
ruleBroken BusRuleInstrumentClassCurrencyCodeValid, True
ruleBroken BusRuleInstrumentClassTickSizeValid, True
ruleBroken BusRuleInstrumentClassTickValueValid, True
ruleBroken BusRuleInstrumentClassSessionStartTimeValid, True
ruleBroken BusRuleInstrumentClassSessionEndTimeValid, True
End Sub




