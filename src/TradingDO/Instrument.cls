VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Instrument"
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

Public Event Loaded()
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

Private mRs As ADODB.recordset

Private mLocalSymbols As InstrumentLocalSymbols

Private mEventRaiser As BusinessObjEventRaiser

Private mMaxNameLength As Long
Private mMaxShortNameLength As Long
Private mMaxSymbolLength As Long
Private mMaxCurrencyCodeLength As Long
Private mMaxOptRightLength As Long

Private mWriters As Collection

Private mWriterCounter As Long

Private mInstrumentClass As InstrumentClass

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
Set mErrorList = New ErrorList
Set mEventRaiser = New BusinessObjEventRaiser
Set mWriters = New Collection
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

Public Property Get category() As String
category = gCategoryToString(mRs("InstrumentCategoryid"))
End Property

Public Property Get categoryid() As InstrumentCategories
categoryid = mRs("InstrumentCategoryid")
End Property

Public Property Get currencyCode() As String
currencyCode = Trim$(mRs!Currency)
End Property

Public Property Get daysBeforeExpiryToSwitch() As Long
daysBeforeExpiryToSwitch = Nz(mRs("DaysBeforeExpiryToSwitch"), 0)
End Property

Public Property Get ErrorList() As ErrorList
Set ErrorList = mErrorList
End Property

Public Property Get Exchange() As String
Exchange = Trim$(Nz(mRs("Exchange"), ""))
End Property

Public Property Get ExpiryDate() As Date
ExpiryDate = Nz(mRs("ExpiryDate"), 0#)
End Property

Public Property Let ExpiryDate(ByVal value As Date)
mRs("ExpiryDate") = value
Dirty
End Property

Public Property Get id() As Long
id = mRs("ID")
End Property

Public Property Get InstrumentClass() As InstrumentClass
If mInstrumentClass Is Nothing Then
    Set mInstrumentClass = mDB.InstrumentClassFactory.loadByID(mRs("instrumentclassid"))
End If
Set InstrumentClass = mInstrumentClass
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

Public Property Get localSymbols() As InstrumentLocalSymbols
If mLocalSymbols Is Nothing Then
    Set mLocalSymbols = New InstrumentLocalSymbols
    mLocalSymbols.connectionString = mDB.connectionString
    mLocalSymbols.load id
End If
Set localSymbols = mLocalSymbols
End Property

Public Property Get month() As String
month = Trim$(Nz(mRs("ExpiryMonth"), ""))
End Property

Public Property Get name() As String
name = Trim$(mRs("NAME"))
End Property

Public Property Let name(ByVal value As String)
ruleBroken BusRuleInstrumentNameValid, _
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

Public Property Get optionRight() As String
optionRight = Nz(mRs("optRight"), "")
End Property

Public Property Let optionRight(ByVal value As String)
validateOptionRight value
mRs("optRight") = Left$(value, mMaxOptRightLength)
Dirty
End Property

Friend Property Let recordset(value As ADODB.recordset)
Set mRs = value

mMaxNameLength = mRs.Fields("NAME").DefinedSize
mMaxShortNameLength = mRs.Fields("SHORTNAME").DefinedSize
mMaxSymbolLength = mRs.Fields("SYMBOL").DefinedSize
mMaxOptRightLength = mRs.Fields("OptRight").DefinedSize
If mRs.EOF Then
    mRs.AddNew
    mNew = True
    setInitialErrors
End If
End Property

Public Property Get sessionEndTime() As Date
sessionEndTime = Nz(mRs("sessionEndTime"), CDate(0))
End Property

Public Property Get sessionStartTime() As Date
sessionStartTime = Nz(mRs("sessionStartTime"), CDate(0))
End Property

Public Property Get shortName() As String
shortName = Trim$(mRs("SHORTNAME"))
End Property

Public Property Let shortName(ByVal value As String)
ruleBroken BusRuleInstrumentShortNameValid, _
    Len(value) = 0 Or Len(value) > mMaxShortNameLength
mRs("SHORTNAME") = Left$(value, mMaxShortNameLength)
Dirty
End Property

Public Property Get strikePrice() As Double
strikePrice = Nz(mRs("strikePrice"), 0#)
End Property

Public Property Let strikePrice(ByVal value As Double)
validateStrikePrice value
mRs("strikePrice") = value
Dirty
End Property

Public Property Let symbol(ByVal value As String)
ruleBroken BusRuleInstrumentSymbolValid, _
    Len(value) = 0 Or Len(value) > mMaxSymbolLength
mRs("SYMBOL") = Left$(value, mMaxSymbolLength)
Dirty
End Property

Public Property Get symbol() As String
symbol = Trim$(mRs("Symbol"))
End Property

Public Property Get TickSize() As Currency
TickSize = mRs("TickSize")
End Property

Public Property Get TickValue() As Currency
TickValue = mRs("TickValue")
End Property

Public Property Get TimeZoneCanonicalName() As String
TimeZoneCanonicalName = Trim$(Nz(mRs("CanonicalTimeZoneName"), ""))
End Property

Public Property Get TimeZoneName() As String
TimeZoneName = Trim$(Nz(mRs("TimeZoneName"), ""))
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

Public Function CreateBarDataReader() As BarDataReader
Set CreateBarDataReader = New BarDataReader
CreateBarDataReader.initialise mDB.connectionString, id
End Function

Public Function CreateBarDataWriter() As BarDataWriter
Set CreateBarDataWriter = New BarDataWriter
CreateBarDataWriter.initialise mDB.connectionString, id, Me
mWriters.add CreateBarDataWriter
End Function

Public Function CreateTickDataReader() As TickDataReader
Set CreateTickDataReader = New TickDataReader
CreateTickDataReader.initialise mDB.connectionString, id
End Function

Public Function CreateTickDataWriter() As TickDataWriter
Set CreateTickDataWriter = New TickDataWriter
mWriterCounter = mWriterCounter + 1
CreateTickDataWriter.initialise mDB.connectionString, id, Me, shortName & "/" & mWriterCounter
mWriters.add CreateTickDataWriter
End Function

Public Sub Delete()
mDeleted = True
End Sub

Friend Sub initialise( _
                ByVal db As TradingDB)
Set mDB = db
End Sub

Friend Sub releaseWriter( _
                ByVal writer As Object)
Dim i As Long
Dim obj As Object

For i = mWriters.count To 1 Step -1
    Set obj = mWriters.item(i)
    If obj Is writer Then
        mWriters.remove i
        Exit Sub
    End If
Next

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
mRs.Update
mConn.Close
Set mRs.ActiveConnection = Nothing

Exit Sub

Err:
Err.Raise Err.Number, _
        "TradingDO26" & "." & "Instrument" & "::" & "save" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Sub

Private Sub setInitialErrors()
' sets errors for fields that are invalid when a new record is created
ruleBroken BusRuleInstrumentNameValid, True
ruleBroken BusRuleInstrumentShortNameValid, True
ruleBroken BusRuleInstrumentSymbolValid, True
End Sub

Private Sub validateOptionRight( _
                ByVal value As String)

If Len(value) > mMaxOptRightLength Then
    ruleBroken BusRuleInstrumentOptionRightvalid, True
    Exit Sub
End If

ruleBroken BusRuleInstrumentOptionRightvalid, _
    ((categoryid = InstrumentCategoryOption Or _
        categoryid = InstrumentCategoryFuturesOption) And _
    (value <> "C" And value <> "P"))

End Sub

Private Sub validateStrikePrice( _
                ByVal value As Double)
ruleBroken BusRuleInstrumentStrikePriceValid, _
    (categoryid = InstrumentCategoryOption Or _
        categoryid = InstrumentCategoryFuturesOption) And _
    value <= 0

End Sub



