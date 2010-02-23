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
                ByVal value As String)
Const ProcName As String = "CurrencyCode"
Dim failpoint As String
On Error GoTo Err

value = Trim$(value)
mDataObj.RuleBroken BusRuleInstrumentCurrencyCodeValid, _
    Len(value) > mMaxCurrencyCodeLength
If Len(value) = 0 Then
    mDataObj.ColumnValue(InstrumentColumnCurrency) = Null
Else
    mDataObj.RuleBroken BusRuleInstrumentCurrencyCodeValid, _
        Not IsValidCurrencyCode(value)
    mDataObj.ColumnValue(InstrumentColumnCurrency) = Left$(value, mMaxCurrencyCodeLength)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get CurrencyCode() As String
Const ProcName As String = "CurrencyCode"
Dim failpoint As String
On Error GoTo Err

If mDataObj.IsNew Then
    CurrencyCode = mDataObj.GetColumnValue(InstrumentColumnCurrency, "")
Else
    CurrencyCode = mDataObj.GetColumnValue(InstrumentColumnCurrencyE, "")
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get CurrencyCodeInheritedFromClass() As Boolean
Const ProcName As String = "CurrencyCodeInheritedFromClass"
Dim failpoint As String
On Error GoTo Err

CurrencyCodeInheritedFromClass = (mDataObj.GetColumnValue(InstrumentColumnCurrency, "") = "")

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DaysBeforeExpiryToSwitch() As Long
Const ProcName As String = "DaysBeforeExpiryToSwitch"
Dim failpoint As String
On Error GoTo Err

DaysBeforeExpiryToSwitch = mDataObj.GetColumnValue(InstrumentColumnSwitchDay, 0)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ExchangeName() As String
Const ProcName As String = "ExchangeName"
Dim failpoint As String
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ExpiryDate() As Date
Const ProcName As String = "ExpiryDate"
Dim failpoint As String
On Error GoTo Err

ExpiryDate = mDataObj.GetColumnValue(InstrumentColumnExpiry, 0#)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ExpiryDate(ByVal value As Date)
Const ProcName As String = "ExpiryDate"
Dim failpoint As String
On Error GoTo Err

mDataObj.ColumnValue(InstrumentColumnExpiry) = Int(value)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ExpiryDateString(ByVal value As String)
Const ProcName As String = "ExpiryDateString"
Dim failpoint As String
On Error GoTo Err

If IsDate(value) Then
    ExpiryDate = CDate(value)
Else
    mDataObj.RuleBroken BusinessRuleIds.BusRuleInstrumentExpiryDateValid, Not IsDate(value)
    ExpiryDate = CDate(0)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get HasBarData() As Boolean
Const ProcName As String = "HasBarData"
Dim failpoint As String
On Error GoTo Err

If mDataObj.IsNew Then
    HasBarData = False
Else
    HasBarData = CBool(mDataObj.GetColumnValue(InstrumentColumnHasBarData, 0))
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get HasTickData() As Boolean
Const ProcName As String = "HasTickData"
Dim failpoint As String
On Error GoTo Err

If mDataObj.IsNew Then
    HasTickData = False
Else
    HasTickData = CBool(mDataObj.GetColumnValue(InstrumentColumnHasTickData, 0))
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let InstrumentClass( _
                ByVal value As InstrumentClass)
Dim instrmnt As instrument
Const ProcName As String = "InstrumentClass"
Dim failpoint As String
On Error GoTo Err

If mDataObj.RuleBroken(BusRuleInstrumentInstrumentClassValid, value Is Nothing) Then Exit Property

If Name <> "" Then
    Set instrmnt = mDB.InstrumentFactory.LoadByName("""" & value.ExchangeName & """/""" & _
                                                    value.Name & """/""" & _
                                                    Name & """")
    If Not instrmnt Is Nothing Then mDataObj.RuleBroken getBusRuleNameValidId, instrmnt.Id <> Id
End If
mDataObj.ColumnValue(InstrumentColumnInstrumentClassID) = value.Id

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get InstrumentClass() As InstrumentClass
Const ProcName As String = "InstrumentClass"
Dim failpoint As String
On Error GoTo Err

Set InstrumentClass = mDB.InstrumentClassFactory.LoadByID(mDataObj.GetColumnValue(InstrumentColumnInstrumentClassID, 0))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get InstrumentClassId() As Long
Const ProcName As String = "InstrumentClassId"
Dim failpoint As String
On Error GoTo Err

InstrumentClassId = mDataObj.GetColumnValue(InstrumentColumnInstrumentClassID, 0)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let InstrumentClassName( _
                ByVal value As String)
Const ProcName As String = "InstrumentClassName"
Dim failpoint As String
On Error GoTo Err

InstrumentClass = mDB.InstrumentClassFactory.LoadByName(value)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get InstrumentClassName() As String
Const ProcName As String = "InstrumentClassName"
Dim failpoint As String
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

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
Dim failpoint As String
On Error GoTo Err

Name = GetName

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Name(ByVal value As String)
' overrides generic property
Const ProcName As String = "Name"
Dim failpoint As String
On Error GoTo Err

value = Trim$(value)
If value = GetName Then Exit Property
If Not mDataObj.RuleBroken(getBusRuleNameValidId, _
                            Len(value) = 0 Or _
                            Len(value) > mDataObj.ColumnMaxLength(GenericColumnName)) Then
    If InstrumentClassName = "" Then
        mDataObj.ColumnValue(GenericColumnName) = Left$(value, mDataObj.ColumnMaxLength(GenericColumnName))
    ElseIf Not mDataObj.RuleBroken(getBusRuleNameValidId, _
                                    Not mDB.InstrumentFactory.LoadByName(ExchangeName & "/" & InstrumentClassName & "/" & value) Is Nothing) Then
        mDataObj.ColumnValue(GenericColumnName) = Left$(value, mDataObj.ColumnMaxLength(GenericColumnName))
    End If
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Property

Public Property Get Notes() As String
Const ProcName As String = "Notes"
Dim failpoint As String
On Error GoTo Err

Notes = mDataObj.GetColumnValue(InstrumentColumnNotes, "")

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Notes(ByVal val As String)
Const ProcName As String = "Notes"
Dim failpoint As String
On Error GoTo Err

mDataObj.ColumnValue(InstrumentColumnNotes) = val

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get OptionRight() As OptionRights
Const ProcName As String = "OptionRight"
Dim failpoint As String
On Error GoTo Err

OptionRight = OptionRightFromString(mDataObj.GetColumnValue(InstrumentColumnOptionRight, ""))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let OptionRight(ByVal value As OptionRights)
Const ProcName As String = "OptionRight"
Dim failpoint As String
On Error GoTo Err

validateOptionRight value
mDataObj.ColumnValue(InstrumentColumnOptionRight) = Left$(OptionRightToString(value), mMaxOptRightLength)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get SecType() As SecurityTypes
Const ProcName As String = "SecType"
Dim failpoint As String
On Error GoTo Err

SecType = gCategoryToSecType(mDataObj.GetColumnValue(InstrumentColumnInstrumentCategoryId))

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get SessionEndTime() As Date
Const ProcName As String = "SessionEndTime"
Dim failpoint As String
On Error GoTo Err

If Not mDataObj.IsNew Then
    SessionEndTime = mDataObj.GetColumnValue(InstrumentColumnSessionEndTime, 0)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get SessionStartTime() As Date
Const ProcName As String = "SessionStartTime"
Dim failpoint As String
On Error GoTo Err

If Not mDataObj.IsNew Then
    SessionStartTime = mDataObj.GetColumnValue(InstrumentColumnSessionStartTime)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ShortName() As String
Const ProcName As String = "ShortName"
Dim failpoint As String
On Error GoTo Err

ShortName = mDataObj.GetColumnValue(InstrumentColumnShortName, "")

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ShortName(ByVal value As String)
Const ProcName As String = "ShortName"
Dim failpoint As String
On Error GoTo Err

value = Trim$(value)
mDataObj.RuleBroken BusRuleInstrumentShortNameValid, _
    Len(value) = 0 Or Len(value) > mMaxShortNameLength
mDataObj.ColumnValue(InstrumentColumnShortName) = Left$(value, mMaxShortNameLength)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get StrikePrice() As Double
Const ProcName As String = "StrikePrice"
Dim failpoint As String
On Error GoTo Err

StrikePrice = mDataObj.GetColumnValue(InstrumentColumnStrikePrice, 0#)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let StrikePrice(ByVal value As Double)
Const ProcName As String = "StrikePrice"
Dim failpoint As String
On Error GoTo Err

validateStrikePrice value
mDataObj.ColumnValue(InstrumentColumnStrikePrice) = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let StrikePriceString(ByVal value As String)
Const ProcName As String = "StrikePriceString"
Dim failpoint As String
On Error GoTo Err

If IsNumeric(value) Then
    StrikePrice = CDbl(value)
Else
    mDataObj.RuleBroken BusRuleInstrumentStrikePriceValid, Not IsNumeric(value)
    StrikePrice = 999.999
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let Symbol(ByVal value As String)
Const ProcName As String = "Symbol"
Dim failpoint As String
On Error GoTo Err

value = Trim$(value)
mDataObj.RuleBroken BusRuleInstrumentSymbolValid, _
                    Len(value) = 0 Or Len(value) > mMaxSymbolLength
mDataObj.ColumnValue(InstrumentColumnSymbol) = Left$(value, mMaxSymbolLength)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Symbol() As String
Const ProcName As String = "Symbol"
Dim failpoint As String
On Error GoTo Err

Symbol = mDataObj.GetColumnValue(InstrumentColumnSymbol, "")

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let TickSize( _
                ByVal value As Double)
Const ProcName As String = "TickSize"
Dim failpoint As String
On Error GoTo Err

mDataObj.RuleBroken BusRuleInstrumentTickSizeValid, _
                    value < 0
If value = 0 Then
    mDataObj.ColumnValue(InstrumentColumnTickSize) = Null
Else
    mDataObj.ColumnValue(InstrumentColumnTickSize) = value
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TickSize() As Double
Const ProcName As String = "TickSize"
Dim failpoint As String
On Error GoTo Err

If mDataObj.IsNew Then
    TickSize = mDataObj.GetColumnValue(InstrumentColumnTickSize, 0)
Else
    TickSize = mDataObj.GetColumnValue(InstrumentColumnTickSizeE)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TickSizeInheritedFromClass() As Boolean
Const ProcName As String = "TickSizeInheritedFromClass"
Dim failpoint As String
On Error GoTo Err

If mDataObj.GetColumnValue(InstrumentColumnTickSize, 0) = 0 Then
    TickSizeInheritedFromClass = True
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let TickSizeString( _
                ByVal value As String)
Const ProcName As String = "TickSizeString"
Dim failpoint As String
On Error GoTo Err

If value = "" Then
    TickSize = 0#
ElseIf Not mDataObj.RuleBroken(BusRuleInstrumentTickSizeValid, Not IsNumeric(value)) Then
    TickSize = 999.999
Else
    TickSize = CDbl(value)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let TickValue( _
                ByVal value As Double)
Const ProcName As String = "TickValue"
Dim failpoint As String
On Error GoTo Err

mDataObj.RuleBroken BusRuleInstrumentTickValueValid, value < 0
If value = 0 Then
    mDataObj.ColumnValue(InstrumentColumnTickValue) = Null
Else
    mDataObj.ColumnValue(InstrumentColumnTickValue) = value
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TickValue() As Double
Const ProcName As String = "TickValue"
Dim failpoint As String
On Error GoTo Err

If mDataObj.IsNew Then
    TickValue = mDataObj.GetColumnValue(InstrumentColumnTickValue, 0)
Else
    TickValue = mDataObj.GetColumnValue(InstrumentColumnTickValueE)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TickValueInheritedFromClass() As Boolean
Const ProcName As String = "TickValueInheritedFromClass"
Dim failpoint As String
On Error GoTo Err

If mDataObj.GetColumnValue(InstrumentColumnTickValue, 0) = 0 Then
    TickValueInheritedFromClass = True
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let TickValueString( _
                ByVal value As String)
Const ProcName As String = "TickValueString"
Dim failpoint As String
On Error GoTo Err

If value = "" Then
    TickValue = 0#
ElseIf Not mDataObj.RuleBroken(BusRuleInstrumentTickValueValid, Not IsNumeric(value)) Then
    TickValue = 999.999
Else
    TickValue = CDbl(value)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TimeZoneName() As String
Const ProcName As String = "TimeZoneName"
Dim failpoint As String
On Error GoTo Err

If mDataObj.IsNew Then
    TimeZoneName = ""
Else
    TimeZoneName = mDataObj.GetColumnValue(InstrumentColumnTimeZoneName)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateTickStream( _
                ByVal fromDate As Date, _
                ByVal toDate As Date) As TickStream
Const ProcName As String = "CreateTickStream"
On Error GoTo Err

Set CreateTickStream = New TickStream
CreateTickStream.Initialise mDB, Id, fromDate, toDate

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function CreateTickWriter( _
                ByVal useAsynchronousWrites As Boolean) As TickWriter
Const ProcName As String = "CreateTickWriter"

On Error GoTo Err

Set CreateTickWriter = New TickWriter
CreateTickWriter.Initialise mDB, Me, useAsynchronousWrites

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function FetchBars( _
                ByVal barTimePeriod As TimePeriod, _
                ByVal maxNumberOfBars As Long, _
                Optional ByVal fromDate As Date, _
                Optional ByVal toDate As Date = MaxDateValue, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal barType As BarTypes = BarTypeTrade) As Bars
Dim bf As BarFetcher
Dim specifier As BarDataSpecifier

Const ProcName As String = "FetchBars"
On Error GoTo Err

Select Case barTimePeriod.Units
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
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Bar time unit not supported"
End Select

Set specifier = New BarDataSpecifier
specifier.barTimePeriod = barTimePeriod
specifier.barType = barType
specifier.fromDate = fromDate
specifier.includeBarsOutsideSession = includeBarsOutsideSession
specifier.maxNumberOfBars = maxNumberOfBars
specifier.toDate = toDate
specifier.SetCustomSessionTimes customSessionStartTime, customSessionEndTime

Set bf = New BarFetcher
bf.Initialise mDB, Me

Set FetchBars = bf.FetchBars(specifier)
Set bf = Nothing

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Public Function FetchbarsAsync( _
                ByVal barTimePeriod As TimePeriod, _
                ByVal maxNumberOfBars As Long, _
                Optional ByVal fromDate As Date, _
                Optional ByVal toDate As Date = MaxDateValue, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal barType As BarTypes = BarTypeTrade, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "FetchbarsAsync"
Dim bf As BarFetcher
Dim specifier As BarDataSpecifier


On Error GoTo Err

Select Case barTimePeriod.Units
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
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Bar time unit not supported"
End Select

Set specifier = New BarDataSpecifier
specifier.barTimePeriod = barTimePeriod
specifier.barType = barType
specifier.fromDate = fromDate
specifier.includeBarsOutsideSession = includeBarsOutsideSession
specifier.maxNumberOfBars = maxNumberOfBars
specifier.toDate = toDate
specifier.SetCustomSessionTimes customSessionStartTime, customSessionEndTime

Set bf = New BarFetcher
bf.Initialise mDB, Me

Set FetchbarsAsync = bf.FetchbarsAsync(specifier, cookie, taskName)
Set bf = Nothing

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "WriteBar"
Dim lBarWriter As New BarWriter


On Error GoTo Err

lBarWriter.Initialise mDB, Id, ShortName

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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "WriteBarAsync"
Dim lBarWriter As New BarWriter


On Error GoTo Err

lBarWriter.Initialise mDB, Id, ShortName

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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
                ByVal value As OptionRights)
Const ProcName As String = "validateOptionRight"



On Error GoTo Err

Select Case value
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
                    value = OptNone)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub validateStrikePrice( _
                ByVal value As Double)
Const ProcName As String = "validateStrikePrice"


On Error GoTo Err

mDataObj.RuleBroken BusRuleInstrumentStrikePriceValid, _
                    (SecType = SecTypeOption Or _
                        SecType = SecTypeFuturesOption) And _
                    value <= 0

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub BusinessDataObject_CancelEdit()
Const ProcName As String = "BusinessDataObject_CancelEdit"
On Error GoTo Err

CancelEdit

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub BusinessDataObject_Delete()
Const ProcName As String = "BusinessDataObject_Delete"
On Error GoTo Err

Delete

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get BusinessDataObject_ErrorList() As ErrorList
Const ProcName As String = "BusinessDataObject_ErrorList"
On Error GoTo Err

Set BusinessDataObject_ErrorList = mDataObj.ErrorList

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BusinessDataObject_EventSource() As BusinessObjEventSource
Const ProcName As String = "BusinessDataObject_EventSource"
On Error GoTo Err

Set BusinessDataObject_EventSource = mDataObj.eventSource

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BusinessDataObject_ID() As Long
Const ProcName As String = "BusinessDataObject_ID"
On Error GoTo Err

BusinessDataObject_ID = Id

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BusinessDataObject_IsDeleted() As Boolean
Const ProcName As String = "BusinessDataObject_IsDeleted"
On Error GoTo Err

BusinessDataObject_IsDeleted = IsDeleted

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BusinessDataObject_IsDirty() As Boolean
Const ProcName As String = "BusinessDataObject_IsDirty"
On Error GoTo Err

BusinessDataObject_IsDirty = IsDirty

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BusinessDataObject_IsNew() As Boolean
Const ProcName As String = "BusinessDataObject_IsNew"
On Error GoTo Err

BusinessDataObject_IsNew = IsNew

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BusinessDataObject_IsValid() As Boolean
Const ProcName As String = "BusinessDataObject_IsValid"
On Error GoTo Err

BusinessDataObject_IsValid = IsValid

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Let BusinessDataObject_Name(ByVal RHS As String)
Const ProcName As String = "BusinessDataObject_Name"
On Error GoTo Err

Name = RHS

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get BusinessDataObject_Name() As String
Const ProcName As String = "BusinessDataObject_Name"
On Error GoTo Err

BusinessDataObject_Name = Name

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub BusinessDataObject_Refresh()
Const ProcName As String = "BusinessDataObject_Refresh"
On Error GoTo Err

Refresh

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mDataObj_Deleted()
Const ProcName As String = "mDataObj_Deleted"
On Error GoTo Err

RaiseEvent Deleted

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Id() As Long
Const ProcName As String = "Id"
On Error GoTo Err

Id = mDataObj.Id

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IsValid() As Boolean
Const ProcName As String = "IsValid"
On Error GoTo Err

IsValid = mDataObj.IsValid

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IsDeleted() As Boolean
Const ProcName As String = "IsDeleted"
On Error GoTo Err

IsDeleted = mDataObj.IsDeleted

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IsDirty() As Boolean
Const ProcName As String = "IsDirty"
On Error GoTo Err

IsDirty = mDataObj.IsDirty

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IsNew() As Boolean
Const ProcName As String = "IsNew"
On Error GoTo Err

IsNew = mDataObj.IsNew

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub CancelEdit()
Const ProcName As String = "CancelEdit"


On Error GoTo Err

mDataObj.CancelEdit
If mDataObj.IsNew Then setInitialErrors

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub Delete()
Const ProcName As String = "Delete"


On Error GoTo Err

mDataObj.Delete

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function GetName() As String
Const ProcName As String = "GetName"


On Error GoTo Err

GetName = mDataObj.Name

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub Refresh()
Const ProcName As String = "Refresh"


On Error GoTo Err

mDataObj.Refresh

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub SetName(ByVal value As String)
Const ProcName As String = "SetName"


On Error GoTo Err

mDataObj.Name = value

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub









