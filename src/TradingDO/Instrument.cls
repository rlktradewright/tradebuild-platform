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

Public Event Clean()
Public Event Deleted()
Public Event Dirty()
Public Event Invalid()
Public Event Valid()

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName               As String = "TradingDO26"
Private Const ModuleName                As String = "Instrument"

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

Private mLocalSymbols As InstrumentLocalSymbols

Private mEventRaiser As BusinessObjEventRaiser

Private mMaxNameLength As Long
Private mMaxShortNameLength As Long
Private mMaxSymbolLength As Long
Private mMaxOptRightLength As Long
Private mMaxCurrencyCodeLength As Long

Private mWriters As Collection

Private mWriterCounter As Long

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

Public Property Let currencyCode( _
                ByVal value As String)
value = Trim$(value)
ruleBroken BusRuleInstrumentCurrencyCodeValid, _
    Len(value) > mMaxCurrencyCodeLength
If Len(value) = 0 Then
    mRs(InstrumentColumnCurrency) = Null
Else
    ruleBroken BusRuleInstrumentCurrencyCodeValid, _
        Not IsValidCurrencyCode(value)
    mRs(InstrumentColumnCurrency) = Left$(value, mMaxCurrencyCodeLength)
End If
Dirty
End Property

Public Property Get currencyCode() As String
currencyCode = Trim$(mRs(InstrumentColumnCurrency))
End Property

Public Property Get currencyCodeInheritedFromClass() As Boolean
If IsNull(mRs(InstrumentColumnCurrency)) Then
    currencyCodeInheritedFromClass = True
End If
End Property

Public Property Get daysBeforeExpiryToSwitch() As Long
daysBeforeExpiryToSwitch = mRs("daysBeforeExpiryToSwitch")
End Property

Public Property Get ErrorList() As ErrorList
Set ErrorList = mErrorList
End Property

Public Property Get exchangeName() As String
exchangeName = Trim$(mRs(InstrumentColumnExchange))
End Property

Public Property Get expiryDate() As Date
expiryDate = Nz(mRs("ExpiryDate"), 0#)
End Property

Public Property Let expiryDate(ByVal value As Date)
mRs("ExpiryDate") = Int(value)
Dirty
End Property

Public Property Let expiryDateString(ByVal value As String)
If IsDate(value) Then
    expiryDate = CDate(value)
Else
    ruleBroken BusinessRuleIds.BusRuleInstrumentExpiryDateValid, Not IsDate(value)
    Dirty
End If
End Property

Public Property Get hasBarData() As Boolean
Dim rs As Recordset

Set rs = New Recordset
rs.Open "select count(*) from perioddata where instrumentid=" & id, mDB.connectionString, adOpenForwardOnly
If rs.Fields(0).value <> 0 Then hasBarData = True
rs.Close
End Property

Public Property Get hasTickData() As Boolean
Dim rs As Recordset

Set rs = New Recordset
rs.Open "select count(*) as count from instrumentdata where instrumentid=" & id, mDB.connectionString, adOpenForwardOnly
If rs.Fields(0).value <> 0 Then hasTickData = True
rs.Close
End Property

Public Property Get id() As Long
id = mRs("ID")
End Property

Public Property Let InstrumentClass( _
                ByVal value As InstrumentClass)
ruleBroken BusRuleInstrumentInstrumentClassValid, value Is Nothing
If Not value Is Nothing Then
    mRs("instrumentclassid") = value.id
End If
Dirty
End Property

Public Property Get InstrumentClass() As InstrumentClass
Set InstrumentClass = mDB.InstrumentClassFactory.loadByID(mRs("instrumentclassid"))
End Property

Public Property Get InstrumentClassId() As Long
InstrumentClassId = mRs("instrumentclassid")
End Property

Public Property Let instrumentClassName( _
                ByVal value As String)
InstrumentClass = mDB.InstrumentClassFactory.loadByName(value)
End Property

Public Property Get instrumentClassName() As String
instrumentClassName = Trim$(mRs("instrumentclassname"))
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
month = format(expiryDate, "yyyymm")
End Property

Public Property Get name() As String
name = Trim$(mRs("NAME"))
End Property

Public Property Let name(ByVal value As String)
value = Trim$(value)
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

Public Property Get optionRight() As OptionRights
optionRight = OptionRightFromString(Nz(mRs("optRight"), ""))
End Property

Public Property Let optionRight(ByVal value As OptionRights)
validateOptionRight value
mRs("optRight") = Left$(OptionRightToString(value), mMaxOptRightLength)
Dirty
End Property

Public Property Get secType() As SecurityTypes
secType = gCategoryToSecType(mRs("instrumentcategoryid"))
End Property

Public Property Get sessionEndTime() As Date
sessionEndTime = mRs("sessionEndTime")
End Property

Public Property Get sessionStartTime() As Date
sessionStartTime = mRs("sessionstartTime")
End Property

Public Property Get shortName() As String
shortName = Trim$(mRs("SHORTNAME"))
End Property

Public Property Let shortName(ByVal value As String)
value = Trim$(value)
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

Public Property Let strikePriceString(ByVal value As String)
If IsNumeric(value) Then
    strikePrice = CDbl(value)
Else
    ruleBroken BusRuleInstrumentStrikePriceValid, IsNumeric(value)
    Dirty
End If
End Property

Public Property Let symbol(ByVal value As String)
value = Trim$(value)
ruleBroken BusRuleInstrumentSymbolValid, _
    Len(value) = 0 Or Len(value) > mMaxSymbolLength
mRs("SYMBOL") = Left$(value, mMaxSymbolLength)
Dirty
End Property

Public Property Get symbol() As String
symbol = Trim$(mRs("Symbol"))
End Property

Public Property Let tickSize( _
                ByVal value As Double)
ruleBroken BusRuleInstrumentTickSizeValid, _
    value < 0
If value = 0 Then
    mRs("TickSize") = Null
Else
    mRs("TickSize") = value
End If
Dirty
End Property

Public Property Get tickSize() As Double
tickSize = mRs(InstrumentColumnTickSize)
End Property

Public Property Get tickSizeInheritedFromClass() As Boolean
If IsNull(mRs(InstrumentColumnTickSize)) Then
    tickSizeInheritedFromClass = True
End If
End Property

Public Property Let tickSizeString( _
                ByVal value As String)
If value = "" Then
    tickSize = 0#
ElseIf Not IsNumeric(value) Then
    ruleBroken BusRuleInstrumentTickSizeValid, True
    Dirty
Else
    tickSize = CDbl(value)
End If
End Property

Public Property Let tickValue( _
                ByVal value As Double)
ruleBroken BusRuleInstrumentTickValueValid, _
    value < 0
If value = 0 Then
    mRs("TickValue") = Null
Else
    mRs("TickValue") = value
End If
Dirty
End Property

Public Property Get tickValue() As Double
tickValue = mRs(InstrumentColumnTickValue)
End Property

Public Property Get tickValueInheritedFromClass() As Boolean
If IsNull(mRs(InstrumentColumnTickValue)) Then
    tickValueInheritedFromClass = True
End If
End Property

Public Property Let tickValueString( _
                ByVal value As String)
If value = "" Then
    tickValue = 0#
ElseIf Not IsNumeric(value) Then
    ruleBroken BusRuleInstrumentTickValueValid, True
    Dirty
Else
    tickValue = CDbl(value)
End If
End Property

Public Property Get timeZoneCanonicalName() As String
timeZoneCanonicalName = Trim$(mRs(InstrumentColumnTimeZone))
End Property

Public Property Get TimeZoneName() As String
TimeZoneName = Trim$(mRs("TimeZoneName"))
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
CreateBarDataReader.initialise mDB.connectionString, id, timeZoneCanonicalName
End Function

Public Function CreateBarDataWriter() As BarDataWriter
Set CreateBarDataWriter = New BarDataWriter
CreateBarDataWriter.initialise mDB.connectionString, id, Me
mWriters.add CreateBarDataWriter
End Function

Public Function CreateTickDataReader() As TickDataReader
Set CreateTickDataReader = New TickDataReader
CreateTickDataReader.initialise mDB.connectionString, id, timeZoneCanonicalName
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

Public Function FetchBars( _
                ByVal barLength As Long, _
                ByVal barLengthUnits As TimePeriodUnits, _
                ByVal maxNumberOfBars As Long, _
                Optional ByVal fromDate As Date, _
                Optional ByVal toDate As Date = MaxDateValue, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal barType As BarTypes = BarTypeTrade) As bars
Dim bf As BarFetcher
Dim specifier As BarDataSpecifier

Select Case barLengthUnits
Case TimePeriodSecond, _
    TimePeriodMinute, _
    TimePeriodHour, _
    TimePeriodDay, _
    TimePeriodWeek, _
    TimePeriodMonth, _
    TimePeriodYear, _
    TimePeriodVolume, _
    TimePeriodTickMovement
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "FetchBars", _
            "Bar time unit not supported"
End Select

Set specifier = New BarDataSpecifier
specifier.barLength = barLength
specifier.barLengthUnits = barLengthUnits
specifier.barType = barType
specifier.fromDate = fromDate
specifier.includeBarsOutsideSession = includeBarsOutsideSession
specifier.maxNumberOfBars = maxNumberOfBars
specifier.toDate = toDate
specifier.setCustomSessionTimes customSessionStartTime, customSessionEndTime

Set bf = New BarFetcher
bf.initialise mDB, Me

Set FetchBars = bf.FetchBars(specifier)
Set bf = Nothing
End Function

Public Function fetchbarsAsync( _
                ByVal barLength As Long, _
                ByVal barLengthUnits As TimePeriodUnits, _
                ByVal maxNumberOfBars As Long, _
                Optional ByVal fromDate As Date, _
                Optional ByVal toDate As Date = MaxDateValue, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal barType As BarTypes = BarTypeTrade) As TaskController
Dim bf As BarFetcher
Dim specifier As BarDataSpecifier

Select Case barLengthUnits
Case TimePeriodSecond, _
    TimePeriodMinute, _
    TimePeriodHour, _
    TimePeriodDay, _
    TimePeriodWeek, _
    TimePeriodMonth, _
    TimePeriodYear, _
    TimePeriodVolume, _
    TimePeriodTickMovement
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "FetchBarsAsync", _
            "Bar time unit not supported"
End Select

Set specifier = New BarDataSpecifier
specifier.barLength = barLength
specifier.barLengthUnits = barLengthUnits
specifier.barType = barType
specifier.fromDate = fromDate
specifier.includeBarsOutsideSession = includeBarsOutsideSession
specifier.maxNumberOfBars = maxNumberOfBars
specifier.toDate = toDate
specifier.setCustomSessionTimes customSessionStartTime, customSessionEndTime

Set bf = New BarFetcher
bf.initialise mDB, Me

Set fetchbarsAsync = bf.fetchbarsAsync(specifier)
Set bf = Nothing
End Function

Friend Sub initialise( _
                ByVal db As TradingDB, _
                Optional ByVal rs As ADODB.Recordset)
Set mDB = db

If Not rs Is Nothing Then initialiseRs rs
End Sub

Friend Sub initialiseRs( _
                ByVal rs As ADODB.Recordset)
Set mRs = rs

mMaxNameLength = mRs.Fields(InstrumentColumnName).DefinedSize
mMaxShortNameLength = mRs.Fields(InstrumentColumnShortName).DefinedSize
mMaxSymbolLength = mRs.Fields(InstrumentColumnSymbol).DefinedSize
mMaxOptRightLength = mRs.Fields(InstrumentColumnOptionRight).DefinedSize
mMaxCurrencyCodeLength = mRs.Fields(InstrumentColumnCurrency).DefinedSize

If mRs.EOF Then
    mRs.AddNew
    mNew = True
    setInitialErrors
End If
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
ruleBroken BusRuleInstrumentNameValid, True
ruleBroken BusRuleInstrumentShortNameValid, True
ruleBroken BusRuleInstrumentSymbolValid, True
End Sub

Private Sub validateOptionRight( _
                ByVal value As OptionRights)

Select Case value
Case OptNone
Case OptCall
Case OptPut
Case Else
    ruleBroken BusRuleInstrumentOptionRightvalid, True
    Exit Sub
End Select

ruleBroken BusRuleInstrumentOptionRightvalid, _
    ((secType = SecTypeOption Or _
        secType = SecTypeFuturesOption) And _
    value = OptNone)

End Sub

Private Sub validateStrikePrice( _
                ByVal value As Double)
ruleBroken BusRuleInstrumentStrikePriceValid, _
    (secType = SecTypeOption Or _
        secType = SecTypeFuturesOption) And _
    value <= 0

End Sub



