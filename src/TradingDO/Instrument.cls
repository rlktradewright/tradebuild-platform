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
Private mMaxTradingClassLength As Long

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

Public Property Let CurrencyCode( _
                ByVal Value As String)
Const ProcName As String = "CurrencyCode"
On Error GoTo Err

Value = Trim$(Value)
mDataObj.RuleBroken BusRuleInstrumentCurrencyCodeValid, _
    Len(Value) > mMaxCurrencyCodeLength
If Len(Value) = 0 Then
    mDataObj.ColumnValue(InstrumentColumnCurrency) = Null
Else
    mDataObj.RuleBroken BusRuleInstrumentCurrencyCodeValid, _
        Not IsValidCurrencyCode(Value)
    mDataObj.ColumnValue(InstrumentColumnCurrency) = Left$(Value, mMaxCurrencyCodeLength)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrencyCode() As String
Const ProcName As String = "CurrencyCode"
On Error GoTo Err

If mDataObj.IsNew Then
    CurrencyCode = mDataObj.GetColumnValue(InstrumentColumnCurrency, "")
Else
    CurrencyCode = mDataObj.GetColumnValue(InstrumentColumnCurrencyE, "")
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrencyCodeInheritedFromClass() As Boolean
Const ProcName As String = "CurrencyCodeInheritedFromClass"
On Error GoTo Err

CurrencyCodeInheritedFromClass = (mDataObj.GetColumnValue(InstrumentColumnCurrency, "") = "")

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DaysBeforeExpiryToSwitch() As Long
Const ProcName As String = "DaysBeforeExpiryToSwitch"
On Error GoTo Err

DaysBeforeExpiryToSwitch = mDataObj.GetColumnValue(InstrumentColumnSwitchDay, 0)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ExchangeName() As String
Const ProcName As String = "ExchangeName"
On Error GoTo Err

If Not mDataObj.IsNew Then
    ExchangeName = mDataObj.GetColumnValue(InstrumentColumnExchangeName)
Else
    If CLng(mDataObj.GetColumnValue(InstrumentColumnInstrumentClassID, 0)) <> 0 Then
        ExchangeName = mDB.InstrumentClassFactory.LoadByID( _
                                    mDataObj.GetColumnValue(InstrumentColumnInstrumentClassID)).ExchangeName
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ExpiryDate() As Date
Const ProcName As String = "ExpiryDate"
On Error GoTo Err

ExpiryDate = mDataObj.GetColumnValue(InstrumentColumnExpiry, 0#)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ExpiryDate(ByVal Value As Date)
Const ProcName As String = "ExpiryDate"
On Error GoTo Err

mDataObj.ColumnValue(InstrumentColumnExpiry) = Int(Value)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ExpiryDateString(ByVal Value As String)
Const ProcName As String = "ExpiryDateString"
On Error GoTo Err

If IsDate(Value) Then
    ExpiryDate = CDate(Value)
Else
    mDataObj.RuleBroken BusinessRuleIds.BusRuleInstrumentExpiryDateValid, Not IsDate(Value)
    ExpiryDate = CDate(0)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HasBarData() As Boolean
Const ProcName As String = "HasBarData"
On Error GoTo Err

If mDataObj.IsNew Then
    HasBarData = False
Else
    HasBarData = CBool(mDataObj.GetColumnValue(InstrumentColumnHasBarData, 0))
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HasTickData() As Boolean
Const ProcName As String = "HasTickData"
On Error GoTo Err

If mDataObj.IsNew Then
    HasTickData = False
Else
    HasTickData = CBool(mDataObj.GetColumnValue(InstrumentColumnHasTickData, 0))
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let InstrumentClass( _
                ByVal Value As InstrumentClass)
Const ProcName As String = "InstrumentClass"
On Error GoTo Err

If mDataObj.RuleBroken(BusRuleInstrumentInstrumentClassValid, Value Is Nothing) Then Exit Property

If Name <> "" Then
    Dim instrmnt As instrument
    Set instrmnt = mDB.InstrumentFactory.LoadByName(Value.ExchangeName & "/" & _
                                                    Value.Name & "/" & _
                                                    Name)
    If Not instrmnt Is Nothing Then mDataObj.RuleBroken getBusRuleNameValidId, instrmnt.Id <> Id
End If
mDataObj.ColumnValue(InstrumentColumnInstrumentClassID) = Value.Id

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get InstrumentClass() As InstrumentClass
Const ProcName As String = "InstrumentClass"
On Error GoTo Err

Set InstrumentClass = mDB.InstrumentClassFactory.LoadByID(mDataObj.GetColumnValue(InstrumentColumnInstrumentClassID, 0))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get InstrumentClassId() As Long
Const ProcName As String = "InstrumentClassId"
On Error GoTo Err

InstrumentClassId = mDataObj.GetColumnValue(InstrumentColumnInstrumentClassID, 0)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let InstrumentClassName( _
                ByVal Value As String)
Const ProcName As String = "InstrumentClassName"
On Error GoTo Err

InstrumentClass = mDB.InstrumentClassFactory.LoadByName(Value)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get InstrumentClassName() As String
Const ProcName As String = "InstrumentClassName"
On Error GoTo Err

If Not mDataObj.IsNew Then
    InstrumentClassName = mDataObj.GetColumnValue(InstrumentColumnInstrumentClassName)
Else
    If CLng(mDataObj.GetColumnValue(InstrumentColumnInstrumentClassID, 0)) <> 0 Then
        InstrumentClassName = mDB.InstrumentClassFactory.LoadByID( _
                                    mDataObj.GetColumnValue(InstrumentColumnInstrumentClassID)).Name
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LocalSymbols() As InstrumentLocalSymbols
'If mLocalSymbols Is Nothing Then
'    Set mLocalSymbols = New InstrumentLocalSymbols
'    mLocalSymbols.ConnectionString = mDB.ConnectionString
'    mLocalSymbols.Load Id
'End If
'Set LocalSymbols = mLocalSymbols
End Property

'Public Property Get month() As String
'month = format(ExpiryDate, "yyyymm")
'End Property

Public Property Get Name() As String
Const ProcName As String = "Name"
On Error GoTo Err

Name = GetName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Name(ByVal Value As String)
' overrides generic property
Const ProcName As String = "Name"
On Error GoTo Err

Value = Trim$(Value)
If Value = GetName Then Exit Property
If Not mDataObj.RuleBroken(getBusRuleNameValidId, _
                            Len(Value) = 0 Or _
                            Len(Value) > mDataObj.ColumnMaxLength(GenericColumnName)) Then
    If InstrumentClassName = "" Then
        mDataObj.ColumnValue(GenericColumnName) = Left$(Value, mDataObj.ColumnMaxLength(GenericColumnName))
    ElseIf Not mDataObj.RuleBroken(getBusRuleNameValidId, _
                                    Not mDB.InstrumentFactory.LoadByName(ExchangeName & "/" & InstrumentClassName & "/" & Value) Is Nothing) Then
        mDataObj.ColumnValue(GenericColumnName) = Left$(Value, mDataObj.ColumnMaxLength(GenericColumnName))
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Notes() As String
Const ProcName As String = "Notes"
On Error GoTo Err

Notes = mDataObj.GetColumnValue(InstrumentColumnNotes, "")

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Notes(ByVal val As String)
Const ProcName As String = "Notes"
On Error GoTo Err

mDataObj.ColumnValue(InstrumentColumnNotes) = val

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OptionRight() As OptionRights
Const ProcName As String = "OptionRight"
On Error GoTo Err

OptionRight = OptionRightFromString(mDataObj.GetColumnValue(InstrumentColumnOptionRight, ""))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let OptionRight(ByVal Value As OptionRights)
Const ProcName As String = "OptionRight"
On Error GoTo Err

validateOptionRight Value
mDataObj.ColumnValue(InstrumentColumnOptionRight) = Left$(OptionRightToString(Value), mMaxOptRightLength)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SecType() As SecurityTypes
Const ProcName As String = "SecType"
On Error GoTo Err

SecType = gCategoryToSecType(mDataObj.GetColumnValue(InstrumentColumnInstrumentCategoryId))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SessionEndTime() As Date
Const ProcName As String = "SessionEndTime"
On Error GoTo Err

If Not mDataObj.IsNew Then
    SessionEndTime = mDataObj.GetColumnValue(InstrumentColumnSessionEndTime, 0)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SessionStartTime() As Date
Const ProcName As String = "SessionStartTime"
On Error GoTo Err

If Not mDataObj.IsNew Then
    SessionStartTime = mDataObj.GetColumnValue(InstrumentColumnSessionStartTime)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ShortName() As String
Const ProcName As String = "ShortName"
On Error GoTo Err

ShortName = mDataObj.GetColumnValue(InstrumentColumnShortName, "")

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ShortName(ByVal Value As String)
Const ProcName As String = "ShortName"
On Error GoTo Err

Value = Trim$(Value)
mDataObj.RuleBroken BusRuleInstrumentShortNameValid, _
    Len(Value) = 0 Or Len(Value) > mMaxShortNameLength
mDataObj.ColumnValue(InstrumentColumnShortName) = Left$(Value, mMaxShortNameLength)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StrikePrice() As Double
Const ProcName As String = "StrikePrice"
On Error GoTo Err

StrikePrice = mDataObj.GetColumnValue(InstrumentColumnStrikePrice, 0#)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let StrikePrice(ByVal Value As Double)
Const ProcName As String = "StrikePrice"
On Error GoTo Err

validateStrikePrice Value
mDataObj.ColumnValue(InstrumentColumnStrikePrice) = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let StrikePriceString(ByVal Value As String)
Const ProcName As String = "StrikePriceString"
On Error GoTo Err

If IsNumeric(Value) Then
    StrikePrice = CDbl(Value)
Else
    mDataObj.RuleBroken BusRuleInstrumentStrikePriceValid, Not IsNumeric(Value)
    StrikePrice = 999.999
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Symbol(ByVal Value As String)
Const ProcName As String = "Symbol"
On Error GoTo Err

Value = Trim$(Value)
mDataObj.RuleBroken BusRuleInstrumentSymbolValid, _
                    Len(Value) = 0 Or Len(Value) > mMaxSymbolLength
mDataObj.ColumnValue(InstrumentColumnSymbol) = Left$(Value, mMaxSymbolLength)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Symbol() As String
Const ProcName As String = "Symbol"
On Error GoTo Err

Symbol = mDataObj.GetColumnValue(InstrumentColumnSymbol, "")

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TickSize( _
                ByVal Value As Double)
Const ProcName As String = "TickSize"
On Error GoTo Err

mDataObj.RuleBroken BusRuleInstrumentTickSizeValid, _
                    Value < 0
If Value = 0 Then
    mDataObj.ColumnValue(InstrumentColumnTickSize) = Null
Else
    mDataObj.ColumnValue(InstrumentColumnTickSize) = Value
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TickSize() As Double
Const ProcName As String = "TickSize"
On Error GoTo Err

If mDataObj.IsNew Then
    TickSize = mDataObj.GetColumnValue(InstrumentColumnTickSize, 0)
Else
    TickSize = mDataObj.GetColumnValue(InstrumentColumnTickSizeE)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TickSizeInheritedFromClass() As Boolean
Const ProcName As String = "TickSizeInheritedFromClass"
On Error GoTo Err

If mDataObj.GetColumnValue(InstrumentColumnTickSize, 0) = 0 Then
    TickSizeInheritedFromClass = True
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TickSizeString( _
                ByVal Value As String)
Const ProcName As String = "TickSizeString"
On Error GoTo Err

If Value = "" Then
    TickSize = 0#
ElseIf Not mDataObj.RuleBroken(BusRuleInstrumentTickSizeValid, Not IsNumeric(Value)) Then
    TickSize = 999.999
Else
    TickSize = CDbl(Value)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TickValue( _
                ByVal Value As Double)
Const ProcName As String = "TickValue"
On Error GoTo Err

mDataObj.RuleBroken BusRuleInstrumentTickValueValid, Value < 0
If Value = 0 Then
    mDataObj.ColumnValue(InstrumentColumnTickValue) = Null
Else
    mDataObj.ColumnValue(InstrumentColumnTickValue) = Value
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TickValue() As Double
Const ProcName As String = "TickValue"
On Error GoTo Err

If mDataObj.IsNew Then
    TickValue = mDataObj.GetColumnValue(InstrumentColumnTickValue, 0)
Else
    TickValue = mDataObj.GetColumnValue(InstrumentColumnTickValueE)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TickValueInheritedFromClass() As Boolean
Const ProcName As String = "TickValueInheritedFromClass"
On Error GoTo Err

If mDataObj.GetColumnValue(InstrumentColumnTickValue, 0) = 0 Then
    TickValueInheritedFromClass = True
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TickValueString( _
                ByVal Value As String)
Const ProcName As String = "TickValueString"
On Error GoTo Err

If Value = "" Then
    TickValue = 0#
ElseIf Not mDataObj.RuleBroken(BusRuleInstrumentTickValueValid, Not IsNumeric(Value)) Then
    TickValue = 999.999
Else
    TickValue = CDbl(Value)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TimeZoneName() As String
Const ProcName As String = "TimeZoneName"
On Error GoTo Err

If mDataObj.IsNew Then
    TimeZoneName = ""
Else
    TimeZoneName = mDataObj.GetColumnValue(InstrumentColumnTimeZoneName)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TradingClass(ByVal Value As String)
Const ProcName As String = "TradingClass"
On Error GoTo Err

Value = Trim$(Value)
mDataObj.RuleBroken BusRuleInstrumentTradingClassValid, _
                    Len(Value) = 0 Or Len(Value) > mMaxTradingClassLength
mDataObj.ColumnValue(InstrumentColumnTradingClass) = Left$(Value, mMaxTradingClassLength)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TradingClass() As String
Const ProcName As String = "TradingClass"
On Error GoTo Err

TradingClass = mDataObj.GetColumnValue(InstrumentColumnTradingClass, "")

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateTickStream( _
                ByVal FromDate As Date, _
                ByVal ToDate As Date) As TickStream
Const ProcName As String = "CreateTickStream"
On Error GoTo Err

Set CreateTickStream = New TickStream
CreateTickStream.Initialise mDB, Id, FromDate, ToDate

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTickWriter( _
                ByVal useAsynchronousWrites As Boolean) As TickWriter
Const ProcName As String = "CreateTickWriter"
On Error GoTo Err

Set CreateTickWriter = New TickWriter
CreateTickWriter.Initialise mDB, Me, useAsynchronousWrites

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchBars( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pMaxNumberOfBars As Long, _
                Optional ByVal pFromDate As Date, _
                Optional ByVal pToDate As Date = MaxDateValue, _
                Optional ByVal pCustomSessionStartTime As Date, _
                Optional ByVal pCustomSessionEndTime As Date, _
                Optional ByVal pIncludeBarsOutsideSession As Boolean, _
                Optional ByVal pBarType As BarTypes = BarTypeTrade, _
                Optional ByVal pListener As IBarFetchListener, _
                Optional ByVal pDontReturnBars As Boolean) As Bars
Const ProcName As String = "FetchBars"
On Error GoTo Err

Select Case pBarTimePeriod.Units
Case TimePeriodSecond, _
    TimePeriodMinute, _
    TimePeriodHour, _
    TimePeriodDay, _
    TimePeriodWeek, _
    TimePeriodMonth, _
    TimePeriodYear, _
    TimePeriodVolume, _
    TimePeriodTickMovement, _
    TimePeriodTickVolume
Case Else
    AssertArgument False, "Bar time unit not supported"
End Select

Dim Specifier As BarDataSpecifier
Set Specifier = CreateBarDataSpecifier( _
                                pBarTimePeriod, _
                                pFromDate, _
                                pToDate, _
                                pMaxNumberOfBars, _
                                pBarType, _
                                False, _
                                pIncludeBarsOutsideSession, _
                                False, _
                                pCustomSessionStartTime, _
                                pCustomSessionEndTime)

Dim bf As BarFetcher
Set bf = New BarFetcher
bf.Initialise mDB, Me

Set FetchBars = bf.FetchBars(Specifier, pListener, pDontReturnBars)
Set bf = Nothing

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchBarsAsync( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pMaxNumberOfBars As Long, _
                Optional ByVal pFromDate As Date, _
                Optional ByVal pToDate As Date = MaxDateValue, _
                Optional ByVal pCustomSessionStartTime As Date, _
                Optional ByVal pCustomSessionEndTime As Date, _
                Optional ByVal pIncludeBarsOutsideSession As Boolean, _
                Optional ByVal pBarType As BarTypes = BarTypeTrade, _
                Optional ByVal pListener As IBarFetchListener, _
                Optional ByVal pDontReturnBars As Boolean, _
                Optional ByVal pCookie As Variant, _
                Optional ByVal pTaskName As String) As IFuture
Const ProcName As String = "FetchbarsAsync"
On Error GoTo Err

Select Case pBarTimePeriod.Units
Case TimePeriodSecond, _
    TimePeriodMinute, _
    TimePeriodHour, _
    TimePeriodDay, _
    TimePeriodWeek, _
    TimePeriodMonth, _
    TimePeriodYear, _
    TimePeriodVolume, _
    TimePeriodTickMovement, _
    TimePeriodTickVolume
Case Else
    AssertArgument False, "Bar time unit not supported"
End Select

Dim Specifier As BarDataSpecifier
Set Specifier = CreateBarDataSpecifier( _
                                pBarTimePeriod, _
                                pFromDate, _
                                pToDate, _
                                pMaxNumberOfBars, _
                                pBarType, _
                                False, _
                                pIncludeBarsOutsideSession, _
                                False, _
                                pCustomSessionStartTime, _
                                pCustomSessionEndTime)

Dim bf As New AsyncBarFetcher
bf.Initialise mDB, Me

bf.FetchBarsAsync Specifier, pListener, pDontReturnBars, pCookie, pTaskName

Set FetchBarsAsync = bf.BarsFuture

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                Optional ByVal rs As ADODB.Recordset)
Const ProcName As String = "Initialise"
On Error GoTo Err

GenericInitialise db, rs

mMaxNameLength = mDataObj.ColumnMaxLength(InstrumentColumnName)
mMaxShortNameLength = mDataObj.ColumnMaxLength(InstrumentColumnShortName)
mMaxSymbolLength = mDataObj.ColumnMaxLength(InstrumentColumnSymbol)
mMaxOptRightLength = mDataObj.ColumnMaxLength(InstrumentColumnOptionRight)
mMaxCurrencyCodeLength = mDataObj.ColumnMaxLength(InstrumentColumnCurrency)
mMaxTradingClassLength = mDataObj.ColumnMaxLength(InstrumentColumnTradingClass)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub WriteBar( _
                ByVal timestamp As Date, _
                ByVal BarType As Long, _
                ByVal barLength As Long, _
                ByVal openPrice As Double, _
                ByVal highPrice As Double, _
                ByVal lowPrice As Double, _
                ByVal closePrice As Double, _
                ByVal volume As Long, _
                ByVal tickVolume As Long, _
                ByVal openInterest As Long)
Const ProcName As String = "WriteBar"
On Error GoTo Err

Dim lBarWriter As New BarWriter
lBarWriter.Initialise mDB, Id, ShortName

lBarWriter.WriteBar timestamp, _
                    BarType, _
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function WriteBarAsync( _
                ByVal timestamp As Date, _
                ByVal BarType As Long, _
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
Const ProcName As String = "WriteBarAsync"
On Error GoTo Err

Dim lBarWriter As New BarWriter
lBarWriter.Initialise mDB, Id, ShortName

Set WriteBarAsync = lBarWriter.WriteBarAsync(timestamp, _
                                            BarType, _
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

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getBusRuleNameValidId() As Long
Const ProcName As String = "getBusRuleNameValidId"
getBusRuleNameValidId = BusRuleInstrumentNameValid
End Function

Private Function getFactory() As DataObjectFactory
Const ProcName As String = "getFactory"
On Error GoTo Err

Set getFactory = mDB.InstrumentFactory

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setInitialErrors()
Const ProcName As String = "setInitialErrors"
' set errors for fields that are invalid when a new record is created
' Note that this is done automatically by the AbstractDataObject for
' the Name field
mDataObj.RuleBroken BusRuleInstrumentShortNameValid, True
mDataObj.RuleBroken BusRuleInstrumentSymbolValid, True
End Sub

Private Sub validateOptionRight( _
                ByVal Value As OptionRights)
Const ProcName As String = "validateOptionRight"
On Error GoTo Err

Select Case Value
Case OptNone
Case OptCall
Case OptPut
Case Else
    mDataObj.RuleBroken BusRuleInstrumentOptionRightvalid, True
    Exit Sub
End Select

mDataObj.RuleBroken BusRuleInstrumentOptionRightvalid, _
                    ((SecType = SecTypeOption Or _
                        SecType = SecTypeFuturesOption) And _
                    Value = OptNone)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub validateStrikePrice( _
                ByVal Value As Double)
Const ProcName As String = "validateStrikePrice"
On Error GoTo Err

mDataObj.RuleBroken BusRuleInstrumentStrikePriceValid, _
                    (SecType = SecTypeOption Or _
                        SecType = SecTypeFuturesOption) And _
                    Value <= 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
Const ProcName As String = "mDataObj_Clean"
On Error GoTo Err

RaiseEvent Clean

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mDataObj_Deleted()
Const ProcName As String = "mDataObj_Deleted"
On Error GoTo Err

RaiseEvent Deleted

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
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

Public Sub SetName(ByVal Value As String)
Const ProcName As String = "SetName"
On Error GoTo Err

mDataObj.Name = Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub









