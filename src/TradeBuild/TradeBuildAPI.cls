VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradeBuildAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

'Implements TradeBuildSP.IStreamingDataConsumer

'@================================================================================
' Events
'@================================================================================

Event Error(ByRef ev As ErrorEvent)
Event Notification(ByRef ev As NotificationEvent)

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mServiceProviderError As ServiceProviderError

Private mServiceProviders As ServiceProviders

Private mListeners As InfoListeners

Private mWorkSpaces As workspaces
Private mDefaultWorkSpace As WorkSpace
Private mTickers As Tickers
Private mExecutions As executions

Private mExchangeCodes() As String
Private mMaxExchangeCodesIndex As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

Set mServiceProviders = New ServiceProviders
mServiceProviders.tradeBuildAPI = Me

Set mListeners = New InfoListeners

Set mWorkSpaces = New workspaces
mWorkSpaces.tradeBuildAPI = Me

Set mDefaultWorkSpace = mWorkSpaces.add("$default")
Set mTickers = defaultWorkSpace.Tickers
Set mExecutions = defaultWorkSpace.executions

ReDim mExchangeCodes(100) As String
mMaxExchangeCodesIndex = -1

addExchangeCode "ACE"
addExchangeCode "AEB"
addExchangeCode "AMEX"
addExchangeCode "ARCA"

addExchangeCode "BELFOX"
addExchangeCode "BOX"
addExchangeCode "BRUT"
addExchangeCode "BTRADE"
addExchangeCode "BVME"

addExchangeCode "CAES"
addExchangeCode "CBOE"
addExchangeCode "CDE"
addExchangeCode "CFE"

addExchangeCode "DTB"

addExchangeCode "EBS"
addExchangeCode "ECBOT"
addExchangeCode "EUREX"
addExchangeCode "EUREXUS"

addExchangeCode "FTA"
addExchangeCode "FWB"

addExchangeCode "GLOBEX"

addExchangeCode "HKFE"

addExchangeCode "IBIS"
addExchangeCode "IDEAL"
addExchangeCode "IDEALPRO"
addExchangeCode "IDEM"
addExchangeCode "INET"
addExchangeCode "INSTINET"
addExchangeCode "ISE"
addExchangeCode "ISLAND"

addExchangeCode "LIFFE"
addExchangeCode "LIFFE_NF"
addExchangeCode "LSE"

addExchangeCode "MATIF"
addExchangeCode "MEFF"
addExchangeCode "MEFFRV"
addExchangeCode "MONEP"
addExchangeCode "MXT"

addExchangeCode "NASDAQ"
addExchangeCode "NQLX"
addExchangeCode "NYMEX"
addExchangeCode "NYSE"

addExchangeCode "OMS"
addExchangeCode "ONE"
addExchangeCode "OSE.JPN"

addExchangeCode "PHLX"
addExchangeCode "PINK"
addExchangeCode "PSE"

addExchangeCode "RDBK"

addExchangeCode "SBF"
addExchangeCode "SFB"
addExchangeCode "SGX"
addExchangeCode "SMART"
addExchangeCode "SNFE"
addExchangeCode "SOFFEX"
addExchangeCode "SWB"
addExchangeCode "SWX"

addExchangeCode "TSE"
addExchangeCode "TSE.JPN"

addExchangeCode "VENTURE"
addExchangeCode "VIRTX"
addExchangeCode "VWAP"

ReDim Preserve mExchangeCodes(mMaxExchangeCodesIndex) As String
End Sub

Private Sub Class_Terminate()
Debug.Print "TradeBuildAPI terminated"
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get defaultWorkSpace() As WorkSpace
Set defaultWorkSpace = mDefaultWorkSpace
End Property

Public Property Get executions() As executions
Set executions = mExecutions
End Property

Public Property Get ServiceProviders() As ServiceProviders
Set ServiceProviders = mServiceProviders
End Property

Public Property Get SupportedHistoricalDataPeriods() As TimePeriod()
SupportedHistoricalDataPeriods = mServiceProviders.SupportedHistoricalDataPeriods
End Property

Public Property Get SupportedInputTickfileFormats() As TradeBuild25.TickfileFormatSpecifier()
SupportedInputTickfileFormats = mServiceProviders.SupportedInputTickfileFormats
End Property

Public Property Get SupportedOutputTickfileFormats() As TradeBuild25.TickfileFormatSpecifier()
SupportedOutputTickfileFormats = mServiceProviders.SupportedOutputTickfileFormats
End Property

Public Property Get Tickers() As Tickers
Set Tickers = mTickers
End Property

Public Property Get workspaces() As workspaces
Set workspaces = mWorkSpaces
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function addInfoListener(ByVal newListener As InfoListener, _
                            Optional ByRef valueType As Long = TradeBuildListenValueTypes.VTAll) As String
addInfoListener = mListeners.add(newListener, valueType)
End Function

Public Function GetExchangeCodes() As String()
GetExchangeCodes = mExchangeCodes
End Function

Public Function getServiceProviderError() As ServiceProviderError
Set getServiceProviderError = mServiceProviderError
Set mServiceProviderError = Nothing
End Function

Public Function IsSupportedHistoricalDataPeriod( _
                ByVal periodLength As Long, _
                ByVal periodUnits As TimePeriodUnits)
Dim period As TimePeriod
period.length = periodLength
period.units = periodUnits
IsSupportedHistoricalDataPeriod = mServiceProviders.IsSupportedHistoricalDataPeriod(period)
End Function

Public Function IsValidExchangeCode(ByVal code As String) As Boolean
Dim bottom As Long
Dim top As Long
Dim middle As Long

code = UCase$(code)
bottom = 0
top = mMaxExchangeCodesIndex
middle = Fix((bottom + top) / 2)

Do
    If code < mExchangeCodes(middle) Then
        top = middle
    ElseIf code > mExchangeCodes(middle) Then
        bottom = middle
    Else
        IsValidExchangeCode = True
        Exit Function
    End If
    middle = Fix((bottom + top) / 2)
Loop Until bottom = middle

If code = mExchangeCodes(middle) Then IsValidExchangeCode = True
End Function

'Public Function NewBars( _
'                ByVal barLength As Long, _
'                ByVal barType As Long) As bars
'Set NewBars = createbars(
'NewBars.tradeBuildAPI = Me
'NewBars.barLength = barLength
'NewBars.barType = barType
'End Function

Public Function NewContracts( _
                ByVal contractSpec As contractSpecifier) As Contracts
Set NewContracts = New Contracts
NewContracts.ServiceProviders = mServiceProviders
NewContracts.contractSpecifier = contractSpec
End Function

Public Function newContractSpecifier( _
                Optional ByVal localSymbol As String, _
                Optional ByVal symbol As String, _
                Optional ByVal exchange As String, _
                Optional ByVal sectype As SecurityTypes = SecTypeNone, _
                Optional ByVal currencyCode As String, _
                Optional ByVal expiry As String, _
                Optional ByVal strike As Double, _
                Optional ByVal right As OptionRights = OptNone) As TradeBuild25.contractSpecifier

Dim source As String: source = "TradeBuild25.TradeBuildAPI::newContractSpecifier"

If localSymbol = "" And symbol = "" Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                source, _
                "Symbol must be supplied if localSymbol is not supplied"
End If

If exchange <> "" And _
    Not IsValidExchangeCode(exchange) _
Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                source, _
                "'" & exchange & "' is not a valid exchange code"
End If

Select Case sectype
Case 0  ' ie not supplied
Case SecTypeStock
Case SecTypeFuture
    If expiry <> "" Then
' need to write a function to validate the expiry format
'        err.Raise ErrorCodes.ErrIllegalArgumentException, _
'                    source, _
'                    "expiry must be supplied"
    End If
Case SecTypeOption, SecTypeFuturesOption
    If expiry <> "" Then
' need to write a function to validate the expiry format
'        err.Raise ErrorCodes.ErrIllegalArgumentException, _
'                    source, _
'                    "expiry must be supplied"
    End If
    If strike < 0 Then
        err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    source, _
                    "strike must be > 0"
    End If
    Select Case right
    Case OptCall
    Case OptPut
    Case OptNone
    Case Else
        err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    source, _
                    "'" & right & "' is not a valid option right"
    End Select
Case SecTypeCash
Case SecTypeBag
Case SecTypeIndex
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                source, _
                "'" & sectype & "' is not a valid secType"
End Select

Set newContractSpecifier = New TradeBuild25.contractSpecifier
With newContractSpecifier
    .localSymbol = localSymbol
    .symbol = symbol
    .exchange = exchange
    .sectype = sectype
    .currencyCode = currencyCode
    .expiry = expiry
    .strike = strike
    .right = right
End With
End Function

Friend Sub notifyFatalError( _
                ByVal errorCode As Long, _
                ByVal errorMessage As String, _
                ByVal source As String)
Dim ev As ErrorEvent

On Error Resume Next

' give the application a chance to handle this

Set ev.source = Me
ev.errorCode = errorCode
ev.errorMsg = errorMessage & "( occurred in " & source & ")"
RaiseEvent Error(ev)

' kill all the service providers if the app hasn't already done so
mServiceProviders.RemoveAll
End Sub

Friend Sub notifyListeners( _
                            ByVal value As Variant, _
                            ByVal valueType As Long, _
                            ByVal source As Object)
mListeners.notify value, valueType, source
End Sub
                            
Public Sub RemoveListener(ByVal key As String)
mListeners.remove key
End Sub

Friend Sub SetServiceProviderError( _
                ByVal errorCode As Long, _
                ByVal errorMessage As String, _
                ByVal serviceProviderName As String)
Dim ev As NotificationEvent

Set mServiceProviderError = New ServiceProviderError
mServiceProviderError.errorCode = errorCode
mServiceProviderError.message = errorMessage
mServiceProviderError.serviceProviderName = serviceProviderName
mServiceProviderError.timestamp = GetTimestamp

Set ev.source = Me
ev.eventCode = ApiNotifyCodes.ApiNotifyServiceProviderError
ev.eventMessage = "A Service Provider error has occurred"
RaiseEvent Notification(ev)
                    
End Sub

'Public Function SetStudyDefaultParameters( _
'                ByVal studyName As String, _
'                ByVal serviceProviderName As String, _
'                ByVal params As IParameters) As Boolean
'SetStudyDefaultParameters = mServiceProviders.SetStudyDefaultParameters( _
'                studyName, _
'                serviceProviderName, _
'                params)
'End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addExchangeCode(ByVal code As String)
mMaxExchangeCodesIndex = mMaxExchangeCodesIndex + 1
If mMaxExchangeCodesIndex > UBound(mExchangeCodes) Then
    ReDim Preserve mExchangeCodes(UBound(mExchangeCodes) + 10) As String
End If
mExchangeCodes(mMaxExchangeCodesIndex) = UCase$(code)
End Sub

