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

' Generic declarations -----------------------------------------------------------
Private WithEvents mDataObj                         As AbstractDataObject
Attribute mDataObj.VB_VarHelpID = -1

Private mDB                                         As TradingDB

' Custom declarations ------------------------------------------------------------

Private mMaxNameLength As Long
Private mMaxShortNameLength As Long
Private mMaxSymbolLength As Long
Private mMaxOptRightLength As Long
Private mMaxCurrencyCodeLength As Long

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

Public Property Let currencyCode( _
                ByVal value As String)
value = Trim$(value)
mDataObj.ruleBroken BusRuleInstrumentCurrencyCodeValid, _
    Len(value) > mMaxCurrencyCodeLength
If Len(value) = 0 Then
    mDataObj.columnValue(InstrumentColumnCurrency) = Null
Else
    mDataObj.ruleBroken BusRuleInstrumentCurrencyCodeValid, _
        Not IsValidCurrencyCode(value)
    mDataObj.columnValue(InstrumentColumnCurrency) = Left$(value, mMaxCurrencyCodeLength)
End If
End Property

Public Property Get currencyCode() As String
If mDataObj.IsNew Then
    currencyCode = mDataObj.getColumnValue(InstrumentColumnCurrency, "")
Else
    currencyCode = mDataObj.getColumnValue(InstrumentColumnCurrencyE, "")
End If
End Property

Public Property Get currencyCodeInheritedFromClass() As Boolean
currencyCodeInheritedFromClass = (mDataObj.getColumnValue(InstrumentColumnCurrency, "") = "")
End Property

Public Property Get daysBeforeExpiryToSwitch() As Long
daysBeforeExpiryToSwitch = mDataObj.getColumnValue(InstrumentColumnSwitchDay, 0)
End Property

Public Property Get exchangeName() As String
If Not mDataObj.IsNew Then
    exchangeName = mDataObj.getColumnValue(InstrumentColumnExchangeName)
Else
    If CLng(mDataObj.getColumnValue(InstrumentColumnInstrumentClassID, 0)) <> 0 Then
        exchangeName = mDB.InstrumentClassFactory.loadByID( _
                                    mDataObj.getColumnValue(InstrumentColumnInstrumentClassID)).exchangeName
    End If
End If
End Property

Public Property Get expiryDate() As Date
expiryDate = mDataObj.getColumnValue(InstrumentColumnExpiry, 0#)
End Property

Public Property Let expiryDate(ByVal value As Date)
mDataObj.columnValue(InstrumentColumnExpiry) = Int(value)
End Property

Public Property Let expiryDateString(ByVal value As String)
If IsDate(value) Then
    expiryDate = CDate(value)
Else
    mDataObj.ruleBroken BusinessRuleIds.BusRuleInstrumentExpiryDateValid, Not IsDate(value)
    expiryDate = CDate(0)
End If
End Property

Public Property Get hasBarData() As Boolean
If mDataObj.IsNew Then
    hasBarData = False
Else
    hasBarData = CBool(mDataObj.getColumnValue(InstrumentColumnHasBarData, 0))
End If
End Property

Public Property Get hasTickData() As Boolean
If mDataObj.IsNew Then
    hasTickData = False
Else
    hasTickData = CBool(mDataObj.getColumnValue(InstrumentColumnHasTickData, 0))
End If
End Property

Public Property Let InstrumentClass( _
                ByVal value As InstrumentClass)
Dim instrmnt As instrument
If mDataObj.ruleBroken(BusRuleInstrumentInstrumentClassValid, value Is Nothing) Then Exit Property

If name <> "" Then
    Set instrmnt = mDB.InstrumentFactory.loadByName(value.exchangeName & "/" & _
                                                    value.name & "/" & _
                                                    name)
    If Not instrmnt Is Nothing Then mDataObj.ruleBroken getBusRuleNameValidId, instrmnt.id <> id
End If
mDataObj.columnValue(InstrumentColumnInstrumentClassID) = value.id
End Property

Public Property Get InstrumentClass() As InstrumentClass
Set InstrumentClass = mDB.InstrumentClassFactory.loadByID(mDataObj.getColumnValue(InstrumentColumnInstrumentClassID, 0))
End Property

Public Property Get InstrumentClassId() As Long
InstrumentClassId = mDataObj.getColumnValue(InstrumentColumnInstrumentClassID, 0)
End Property

Public Property Let instrumentClassName( _
                ByVal value As String)
InstrumentClass = mDB.InstrumentClassFactory.loadByName(value)
End Property

Public Property Get instrumentClassName() As String
If Not mDataObj.IsNew Then
    instrumentClassName = mDataObj.getColumnValue(InstrumentColumnInstrumentClassName)
Else
    If CLng(mDataObj.getColumnValue(InstrumentColumnInstrumentClassID, 0)) <> 0 Then
        instrumentClassName = mDB.InstrumentClassFactory.loadByID( _
                                    mDataObj.getColumnValue(InstrumentColumnInstrumentClassID)).name
    End If
End If

End Property

Public Property Get localSymbols() As InstrumentLocalSymbols
'If mLocalSymbols Is Nothing Then
'    Set mLocalSymbols = New InstrumentLocalSymbols
'    mLocalSymbols.connectionString = mDB.connectionString
'    mLocalSymbols.load id
'End If
'Set localSymbols = mLocalSymbols
End Property

'Public Property Get month() As String
'month = format(expiryDate, "yyyymm")
'End Property

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
    If instrumentClassName = "" Then
        mDataObj.columnValue(GenericColumnName) = Left$(value, mDataObj.columnMaxLength(GenericColumnName))
    ElseIf Not mDataObj.ruleBroken(getBusRuleNameValidId, _
                                    Not mDB.InstrumentFactory.loadByName(exchangeName & "/" & instrumentClassName & "/" & name) Is Nothing) Then
        mDataObj.columnValue(GenericColumnName) = Left$(value, mDataObj.columnMaxLength(GenericColumnName))
    End If
End If

End Property

Public Property Get notes() As String
notes = mDataObj.getColumnValue(InstrumentColumnNotes, "")
End Property

Public Property Let notes(ByVal val As String)
mDataObj.columnValue(InstrumentColumnNotes) = val
End Property

Public Property Get optionRight() As OptionRights
optionRight = OptionRightFromString(mDataObj.getColumnValue(InstrumentColumnOptionRight, ""))
End Property

Public Property Let optionRight(ByVal value As OptionRights)
validateOptionRight value
mDataObj.columnValue(InstrumentColumnOptionRight) = Left$(OptionRightToString(value), mMaxOptRightLength)
End Property

Public Property Get secType() As SecurityTypes
secType = gCategoryToSecType(mDataObj.getColumnValue(InstrumentColumnInstrumentCategoryId))
End Property

Public Property Get sessionEndTime() As Date
If Not mDataObj.IsNew Then
    sessionEndTime = mDataObj.getColumnValue(InstrumentColumnSessionEndTime, 0)
End If
End Property

Public Property Get sessionStartTime() As Date
If Not mDataObj.IsNew Then
    sessionStartTime = mDataObj.getColumnValue(InstrumentColumnSessionStartTime)
End If
End Property

Public Property Get shortName() As String
shortName = mDataObj.getColumnValue(InstrumentColumnShortName, "")
End Property

Public Property Let shortName(ByVal value As String)
value = Trim$(value)
mDataObj.ruleBroken BusRuleInstrumentShortNameValid, _
    Len(value) = 0 Or Len(value) > mMaxShortNameLength
mDataObj.columnValue(InstrumentColumnShortName) = Left$(value, mMaxShortNameLength)
End Property

Public Property Get strikePrice() As Double
strikePrice = mDataObj.getColumnValue(InstrumentColumnStrikePrice, 0#)
End Property

Public Property Let strikePrice(ByVal value As Double)
validateStrikePrice value
mDataObj.columnValue(InstrumentColumnStrikePrice) = value
End Property

Public Property Let strikePriceString(ByVal value As String)
If IsNumeric(value) Then
    strikePrice = CDbl(value)
Else
    mDataObj.ruleBroken BusRuleInstrumentStrikePriceValid, Not IsNumeric(value)
    strikePrice = 999.999
End If
End Property

Public Property Let symbol(ByVal value As String)
value = Trim$(value)
mDataObj.ruleBroken BusRuleInstrumentSymbolValid, _
                    Len(value) = 0 Or Len(value) > mMaxSymbolLength
mDataObj.columnValue(InstrumentColumnSymbol) = Left$(value, mMaxSymbolLength)
End Property

Public Property Get symbol() As String
symbol = mDataObj.getColumnValue(InstrumentColumnSymbol, "")
End Property

Public Property Let tickSize( _
                ByVal value As Double)
mDataObj.ruleBroken BusRuleInstrumentTickSizeValid, _
                    value < 0
If value = 0 Then
    mDataObj.columnValue(InstrumentColumnTickSize) = Null
Else
    mDataObj.columnValue(InstrumentColumnTickSize) = value
End If
End Property

Public Property Get tickSize() As Double
If mDataObj.IsNew Then
    tickSize = mDataObj.getColumnValue(InstrumentColumnTickSize, 0)
Else
    tickSize = mDataObj.getColumnValue(InstrumentColumnTickSizeE)
End If
End Property

Public Property Get tickSizeInheritedFromClass() As Boolean
If mDataObj.getColumnValue(InstrumentColumnTickSize, 0) = 0 Then
    tickSizeInheritedFromClass = True
End If
End Property

Public Property Let tickSizeString( _
                ByVal value As String)
If value = "" Then
    tickSize = 0#
ElseIf Not mDataObj.ruleBroken(BusRuleInstrumentTickSizeValid, Not IsNumeric(value)) Then
    tickSize = 999.999
Else
    tickSize = CDbl(value)
End If
End Property

Public Property Let tickValue( _
                ByVal value As Double)
mDataObj.ruleBroken BusRuleInstrumentTickValueValid, value < 0
If value = 0 Then
    mDataObj.columnValue(InstrumentColumnTickValue) = Null
Else
    mDataObj.columnValue(InstrumentColumnTickValue) = value
End If
End Property

Public Property Get tickValue() As Double
If mDataObj.IsNew Then
    tickValue = mDataObj.getColumnValue(InstrumentColumnTickValue, 0)
Else
    tickValue = mDataObj.getColumnValue(InstrumentColumnTickValueE)
End If
End Property

Public Property Get tickValueInheritedFromClass() As Boolean
If mDataObj.getColumnValue(InstrumentColumnTickValue, 0) = 0 Then
    tickValueInheritedFromClass = True
End If
End Property

Public Property Let tickValueString( _
                ByVal value As String)
If value = "" Then
    tickValue = 0#
ElseIf Not mDataObj.ruleBroken(BusRuleInstrumentTickValueValid, Not IsNumeric(value)) Then
    tickValue = 999.999
Else
    tickValue = CDbl(value)
End If
End Property

Public Property Get TimeZoneName() As String
If mDataObj.IsNew Then
    TimeZoneName = ""
Else
    TimeZoneName = mDataObj.getColumnValue(InstrumentColumnTimeZoneName)
End If
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function createTickStream( _
                ByVal fromDate As Date, _
                ByVal toDate As Date) As TickStream
Set createTickStream = New TickStream
createTickStream.initialise mDB, id, fromDate, toDate
End Function

Public Function createTickWriter( _
                ByVal useAsynchronousWrites As Boolean) As TickWriter
Set createTickWriter = New TickWriter
createTickWriter.initialise mDB, Me, useAsynchronousWrites
End Function

Public Function fetchBars( _
                ByVal barLength As Long, _
                ByVal barLengthUnits As TimePeriodUnits, _
                ByVal maxNumberOfBars As Long, _
                Optional ByVal fromDate As Date, _
                Optional ByVal toDate As Date = MaxDateValue, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal barType As BarTypes = BarTypeTrade) As Bars
Dim bf As BarFetcher
Dim specifier As BarDataSpecifier

Dim failpoint As Long
On Error GoTo Err

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
            ProjectName & "." & ModuleName & ":" & "fetchBars", _
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

Set fetchBars = bf.fetchBars(specifier)
Set bf = Nothing

Exit Function

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "fetchBars" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

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
                Optional ByVal barType As BarTypes = BarTypeTrade, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
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

Set fetchbarsAsync = bf.fetchbarsAsync(specifier, cookie, taskName)
Set bf = Nothing
End Function

Friend Sub initialise( _
                ByVal db As TradingDB, _
                Optional ByVal rs As ADODB.Recordset)
genericInitialise db, rs

mMaxNameLength = mDataObj.columnMaxLength(InstrumentColumnName)
mMaxShortNameLength = mDataObj.columnMaxLength(InstrumentColumnShortName)
mMaxSymbolLength = mDataObj.columnMaxLength(InstrumentColumnSymbol)
mMaxOptRightLength = mDataObj.columnMaxLength(InstrumentColumnOptionRight)
mMaxCurrencyCodeLength = mDataObj.columnMaxLength(InstrumentColumnCurrency)

End Sub

Public Sub WriteBar( _
                ByVal timestamp As Date, _
                ByVal barType As Long, _
                ByVal barLength As Long, _
                ByVal openPrice As Double, _
                ByVal highPrice As Double, _
                ByVal lowPrice As Double, _
                ByVal closePrice As Double, _
                ByVal volume As Long, _
                ByVal tickVolume As Long, _
                ByVal openInterest As Long)
Dim lBarWriter As New BarWriter

Dim failpoint As Long
On Error GoTo Err

lBarWriter.initialise mDB, id, shortName

lBarWriter.WriteBar timestamp, _
                    barType, _
                    barLength, _
                    openPrice, _
                    highPrice, _
                    lowPrice, _
                    closePrice, _
                    volume, _
                    tickVolume, _
                    openInterest

Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "WriteBar" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Sub

Public Function WriteBarAsync( _
                ByVal timestamp As Date, _
                ByVal barType As Long, _
                ByVal barLength As Long, _
                ByVal openPrice As Double, _
                ByVal highPrice As Double, _
                ByVal lowPrice As Double, _
                ByVal closePrice As Double, _
                ByVal volume As Long, _
                ByVal tickVolume As Long, _
                ByVal openInterest As Long, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim lBarWriter As New BarWriter

lBarWriter.initialise mDB, id, shortName

Set WriteBarAsync = lBarWriter.WriteBarAsync(timestamp, _
                                            barType, _
                                            barLength, _
                                            openPrice, _
                                            highPrice, _
                                            lowPrice, _
                                            closePrice, _
                                            volume, _
                                            tickVolume, _
                                            openInterest, _
                                            cookie, _
                                            taskName)
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getBusRuleNameValidId() As Long
getBusRuleNameValidId = BusRuleInstrumentNameValid
End Function

Private Function getFactory() As DataObjectFactory
Set getFactory = mDB.InstrumentFactory
End Function

Private Sub setInitialErrors()
' set errors for fields that are invalid when a new record is created
' Note that this is done automatically by the AbstractDataObject for
' the name field
mDataObj.ruleBroken BusRuleInstrumentShortNameValid, True
mDataObj.ruleBroken BusRuleInstrumentSymbolValid, True
End Sub

Private Sub validateOptionRight( _
                ByVal value As OptionRights)

Select Case value
Case OptNone
Case OptCall
Case OptPut
Case Else
    mDataObj.ruleBroken BusRuleInstrumentOptionRightvalid, True
    Exit Sub
End Select

mDataObj.ruleBroken BusRuleInstrumentOptionRightvalid, _
                    ((secType = SecTypeOption Or _
                        secType = SecTypeFuturesOption) And _
                    value = OptNone)

End Sub

Private Sub validateStrikePrice( _
                ByVal value As Double)
mDataObj.ruleBroken BusRuleInstrumentStrikePriceValid, _
                    (secType = SecTypeOption Or _
                        secType = SecTypeFuturesOption) And _
                    value <= 0

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









