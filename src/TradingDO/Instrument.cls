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

Private mTradingDB As tradingDB
Private mConnectionString As String

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
Private mMaxExpiryMonthLength As Long
Private mMaxCurrencyCodeLength As Long
Private mMaxOptRightLength As Long

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
                ByVal item As TWUtilities.ErrorItem)
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

Public Property Let category(ByVal value As String)
Dim catid As InstrumentCategories

catid = gCategoryFromString(value)
ruleBroken BusRuleInstrumentCategoryInvalid, catid = 0
mRs("categoryid") = catid

Select Case catid
Case InstrumentCategoryStock
    
Case InstrumentCategoryFuture
    ValidateMonth month
Case InstrumentCategoryOption
    ValidateMonth month
    validateOptionRight optionRight
    validateStrikePrice strikePrice
Case InstrumentCategoryCash

Case InstrumentCategoryFuturesOption
    ValidateMonth month
    validateOptionRight optionRight
    validateStrikePrice strikePrice
Case InstrumentCategoryIndex

End Select
Dirty
End Property

Public Property Get category() As String
category = gCategoryToString(mRs("Categoryid"))
End Property

Public Property Get categoryid() As InstrumentCategories
categoryid = mRs("Categoryid")
End Property

Friend Property Let connectionString( _
                ByVal value As String)
mConnectionString = value
End Property

'Public Property Get contractID() As Long
'contractID = Nz(mRs("contractID"), 0)
'End Property
'
'Public Property Let contractID(ByVal value As Long)
'mRs("contractID") = value
'Dirty
'End Property

Public Property Get currencyCode() As String
currencyCode = mRs!Currency
End Property

Public Property Let currencyCode(ByVal value As String)
ruleBroken BusRuleCurrencyCodeInvalid, _
    Len(value) = 0 Or Len(value) > mMaxCurrencyCodeLength
mRs!Currency = Left$(value, mMaxCurrencyCodeLength)
Dirty
End Property

Public Property Get daysBeforeExpiryToSwitch() As Long
daysBeforeExpiryToSwitch = Nz(mRs("DaysBeforeExpiryToSwitch"), 0)
End Property

Public Property Let daysBeforeExpiryToSwitch(ByVal value As Long)
mRs("DaysBeforeExpiryToSwitch") = value
Dirty
End Property

'Public Property Get defaultTradeQuantity() As Long
'defaultTradeQuantity = mRs("defaultTradeQuantity")
'End Property
'
'Public Property Let defaultTradeQuantity(ByVal value As Long)
'ruleBroken BR_INSTRUMENT_DEFAULT_QUANTITY_INVALID, value <= 0
'mRs("defaultTradeQuantity") = value
'Dirty
'End Property

Public Property Get ErrorList() As ErrorList
Set ErrorList = mErrorList
End Property

Public Property Let Exchange(ByVal value As String)
Static exchangeFactory As exchangeFactory
Dim Exchange As Exchange

If exchangeFactory Is Nothing Then Set exchangeFactory = New exchangeFactory

Set Exchange = exchangeFactory.loadByName(value)

ruleBroken BusRuleInstrumentExchangeInvalid, _
    Exchange Is Nothing
If Not Exchange Is Nothing Then mRs("exchangeid") = Exchange.id
Dirty
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
    mLocalSymbols.connectionString = mConnectionString
    mLocalSymbols.load id
End If
Set localSymbols = mLocalSymbols
End Property

Public Property Let month(ByVal value As String)

ValidateMonth value
mRs("expirymonth") = Left$(value, mMaxExpiryMonthLength)
Dirty
End Property

Public Property Get month() As String
month = Trim$(Nz(mRs("ExpiryMonth"), ""))
End Property

Public Property Get name() As String
name = Trim$(mRs("NAME"))
End Property

Public Property Let name(ByVal value As String)
ruleBroken BusRuleInstrumentNameInvalid, _
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
mMaxExpiryMonthLength = mRs.Fields("EXPIRYMONTH").DefinedSize
mMaxCurrencyCodeLength = mRs.Fields("Currency").DefinedSize
mMaxOptRightLength = mRs.Fields("OptRight").DefinedSize
If mRs.EOF Then
    mRs.AddNew
    mNew = True
    setInitialErrors
End If
End Property

Public Property Get sessionEndTime() As Date
sessionEndTime = Nz(mRs("sessionEndTime"), CDate("31/12/2999"))
End Property

Public Property Let sessionEndTime(value As Date)
ruleBroken BusRuleInstrumentSessionStartAndEndEqual, _
    value = sessionStartTime
mRs("sessionEndTime") = value
Dirty
End Property

Public Property Get sessionStartTime() As Date
sessionStartTime = Nz(mRs("sessionStartTime"), CDate("01/01/1900"))
End Property

Public Property Let sessionStartTime(value As Date)
ruleBroken BusRuleInstrumentSessionStartAndEndEqual, _
    value = sessionEndTime
mRs("sessionStartTime") = value
Dirty
End Property

Public Property Get shortName() As String
shortName = Trim$(mRs("SHORTNAME"))
End Property

Public Property Let shortName(ByVal value As String)
ruleBroken BusRuleInstrumentShortNameInvalid, _
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
ruleBroken BusRuleInstrumentSymbolInvalid, _
    Len(value) = 0 Or Len(value) > mMaxSymbolLength
mRs("SYMBOL") = Left$(value, mMaxSymbolLength)
Dirty
End Property

Public Property Get symbol() As String
symbol = Trim$(mRs("Symbol"))
End Property

Public Property Let TickSize(ByVal value As Currency)
ruleBroken BusRuleInstrumentTickSizeInvalid, _
    value <= 0
mRs("TickSize") = value
Dirty
End Property

Public Property Get TickSize() As Currency
TickSize = mRs("TickSize")
End Property

Public Property Let TickValue(ByVal value As Currency)
ruleBroken BusRuleInstrumentTickValueInvalid, _
    value <= 0
mRs("TickValue") = value
Dirty
End Property

Public Property Get TickValue() As Currency
TickValue = mRs("TickValue")
End Property

'Public Property Get TradingDayOffsetMinutes() As Integer
'TradingDayOffsetMinutes = mRs("tradingDayOffsetMinutes")
'End Property
'
'Public Property Let TradingDayOffsetMinutes(value As Integer)
'ruleBroken BR_INSTRUMENT_TRADING_DAY_OFFSET_INVALID, _
'    value < 0 Or value > 59
'mRs("tradingDayOffsetMinutes") = value
'Dirty
'End Property

Friend Property Let tradingDB( _
                ByVal value As tradingDB)
Set mTradingDB = value
End Property

'Public Property Get tradingEndTime() As Date
'tradingEndTime = Nz(mRs("tradingEndTime"), CDate("31/12/2999"))
'End Property
'
'Public Property Let tradingEndTime(value As Date)
'ruleBroken BR_INSTRUMENT_TRADING_START_AND_END_EQUAL, _
'    value = tradingStartTime
'mRs("tradingEndTime") = value
'Dirty
'End Property

'Public Property Get tradingStartTime() As Date
'tradingStartTime = Nz(mRs("tradingStartTime"), CDate("01/01/1900"))
'End Property
'
'Public Property Let tradingStartTime(value As Date)
'ruleBroken BR_INSTRUMENT_TRADING_START_AND_END_EQUAL, _
'    value = tradingEndTime
'mRs("tradingStartTime") = value
'Dirty
'End Property

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

Public Function CreateTickDataReader() As TickDataReader
Set CreateTickDataReader = New TickDataReader
CreateTickDataReader.initialise mConnectionString, id
End Function

Public Function CreateTickDataWriter() As TickDataWriter
Set CreateTickDataWriter = New TickDataWriter
CreateTickDataWriter.initialise mConnectionString, id
End Function

Public Function CreateBarDataReader() As BarDataReader
Set CreateBarDataReader = New BarDataReader
CreateBarDataReader.initialise mConnectionString, id
End Function

Public Function CreateBarDataWriter() As BarDataWriter
Set CreateBarDataWriter = New BarDataWriter
CreateBarDataWriter.initialise mConnectionString, id
End Function

Public Sub Delete()
mDeleted = True
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
mConn.Open mConnectionString

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
ruleBroken BusRuleInstrumentNameInvalid, True
ruleBroken BusRuleInstrumentShortNameInvalid, True
ruleBroken BusRuleInstrumentSymbolInvalid, True
'ruleBroken BusRuleInstrumentExpiryMonthInvalid, True
ruleBroken BusRuleInstrumentTickSizeInvalid, True
ruleBroken BusRuleInstrumentTickValueInvalid, True
ruleBroken BusRuleInstrumentCategoryInvalid, True
ruleBroken BusRuleInstrumentExchangeInvalid, True
ruleBroken BusRuleInstrumentSessionStartAndEndEqual, True
ruleBroken BusRuleCurrencyCodeInvalid, True
End Sub

Private Sub ValidateMonth( _
                ByVal value As String)
Dim theYear As String
Dim theMonth As String

If Len(value) > mMaxExpiryMonthLength Then
    ruleBroken BusRuleInstrumentExpiryMonthInvalid, True
    Exit Sub
End If

If categoryid = InstrumentCategoryFuture Or _
    categoryid = InstrumentCategoryFuturesOption Or _
    categoryid = InstrumentCategoryOption _
Then
    If Len(value) = mMaxExpiryMonthLength And IsNumeric(value) Then
        theYear = Left$(value, 4)
        theMonth = Right$(value, 2)
        ruleBroken BusRuleInstrumentExpiryMonthInvalid, _
                    theMonth <= 0 Or theMonth > 12
        ruleBroken BusRuleInstrumentExpiryMonthOutOfRange, _
                    theYear < 1990 Or theYear > Year(Date) + 2
    Else
        ruleBroken BusRuleInstrumentExpiryMonthInvalid, True
    End If
End If

End Sub

Public Sub validateOptionRight( _
                ByVal value As String)

If Len(value) > mMaxOptRightLength Then
    ruleBroken BusRuleOptionRightInvalid, True
    Exit Sub
End If

ruleBroken BusRuleOptionRightInvalid, _
    ((categoryid = InstrumentCategoryOption Or _
        categoryid = InstrumentCategoryFuturesOption) And _
    (value <> "C" And value <> "P"))

End Sub

Public Sub validateStrikePrice( _
                ByVal value As Double)
ruleBroken BusRuleStrikePriceInvalid, _
    (categoryid = InstrumentCategoryOption Or _
        categoryid = InstrumentCategoryFuturesOption) And _
    value <= 0

End Sub



