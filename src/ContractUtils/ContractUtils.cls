VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractUtils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
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

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Public Enum ContractSortKeyIds
    ContractSortKeyNone
    ContractSortKeyLocalSymbol
    ContractSortKeySymbol
    ContractSortKeySecType
    ContractSortKeyExchange
    ContractSortKeyExpiry
    ContractSortKeyCurrency
    ContractSortKeyRight
    ContractSortKeyStrike
End Enum

Public Enum ContractStoreCapabilities
    'ContractStoreCanStore = 1
    ContractStoreCanQuery = 2
End Enum

Public Enum OptionRights
    OptNone = 0
    OptCall
    OptPut
End Enum
    
Public Enum SecurityTypes
    SecTypeNone
    SecTypeStock
    SecTypeFuture
    SecTypeOption
    SecTypeFuturesOption
    SecTypeCash
    SecTypeCombo
    SecTypeIndex
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type CurrencyDescriptor
    code                As String
    Description         As String
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "ContractUtils"

'@================================================================================
' Member variables
'@================================================================================

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

Public Property Let ExactSixtyFourthIndicators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gExactSixtyFourthIndicators = ar
End Property
                
Public Property Get ExactSixtyFourthIndicators() As String()
ExactSixtyFourthIndicators = gExactSixtyFourthIndicators
End Property
                
Public Property Let ExactThirtySecondIndicators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gExactThirtySecondIndicators = ar
End Property
                
Public Property Get ExactThirtySecondIndicators() As String()
ExactThirtySecondIndicators = gExactThirtySecondIndicators
End Property
                
Public Property Let HalfSixtyFourthIndicators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gHalfSixtyFourthIndicators = ar
End Property
                
Public Property Get HalfSixtyFourthIndicators() As String()
HalfSixtyFourthIndicators = gHalfSixtyFourthIndicators
End Property
                
Public Property Let HalfThirtySecondIndicators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gHalfThirtySecondIndicators = ar
End Property
                
Public Property Get HalfThirtySecondIndicators() As String()
HalfThirtySecondIndicators = gHalfThirtySecondIndicators
End Property
                
Public Property Let QuarterThirtySecondIndicators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gQuarterThirtySecondIndicators = ar
End Property
                
Public Property Get QuarterThirtySecondIndicators() As String()
QuarterThirtySecondIndicators = gQuarterThirtySecondIndicators
End Property
                
Public Property Let SixtyFourthsAndFractionsSeparators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gSixtyFourthsAndFractionsSeparators = ar
End Property
                
Public Property Get SixtyFourthsAndFractionsSeparators() As String()
SixtyFourthsAndFractionsSeparators = gSixtyFourthsAndFractionsSeparators
End Property
                
Public Property Let SixtyFourthsAndFractionsTerminators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gSixtyFourthsAndFractionsTerminators = ar
End Property
                
Public Property Get SixtyFourthsAndFractionsTerminators() As String()
SixtyFourthsAndFractionsTerminators = gSixtyFourthsAndFractionsTerminators
End Property
                
Public Property Let SixtyFourthsSeparators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gSixtyFourthsSeparators = ar
End Property
                
Public Property Get SixtyFourthsSeparators() As String()
SixtyFourthsSeparators = gSixtyFourthsSeparators
End Property
                
Public Property Let SixtyFourthsTerminators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gSixtyFourthsTerminators = ar
End Property
                
Public Property Get SixtyFourthsTerminators() As String()
SixtyFourthsTerminators = gSixtyFourthsTerminators
End Property
                
Public Property Let ThirtySecondsAndFractionsSeparators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gThirtySecondsAndFractionsSeparators = ar
End Property
                
Public Property Get ThirtySecondsAndFractionsSeparators() As String()
ThirtySecondsAndFractionsSeparators = gThirtySecondsAndFractionsSeparators
End Property
                
Public Property Let ThirtySecondsAndFractionsTerminators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gThirtySecondsAndFractionsTerminators = ar
End Property
                
Public Property Get ThirtySecondsAndFractionsTerminators() As String()
ThirtySecondsAndFractionsTerminators = gThirtySecondsAndFractionsTerminators
End Property
                
Public Property Let ThirtySecondsSeparators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gThirtySecondsSeparators = ar
End Property
                
Public Property Get ThirtySecondsSeparators() As String()
ThirtySecondsSeparators = gThirtySecondsSeparators
End Property
                
Public Property Let ThirtySecondsTerminators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gThirtySecondsTerminators = ar
End Property
                
Public Property Get ThirtySecondsTerminators() As String()
ThirtySecondsTerminators = gThirtySecondsTerminators
End Property
                
Public Property Let ThreeQuarterThirtySecondIndicators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gThreeQuarterThirtySecondIndicators = ar
End Property
                
Public Property Get ThreeQuarterThirtySecondIndicators() As String()
ThreeQuarterThirtySecondIndicators = gThreeQuarterThirtySecondIndicators
End Property
                
'@================================================================================
' Methods
'@================================================================================

Public Function ContractSpecsCompare( _
                ByVal pContractSpec1 As IContractSpecifier, _
                ByVal pContractSpec2 As IContractSpecifier, _
                ByRef pSortkeys() As ContractSortKeyIds) As Long
Const ProcName As String = "ContractSpecsCompare"
On Error GoTo Err

ContractSpecsCompare = gContractSpecsCompare(pContractSpec1, pContractSpec2, pSortkeys)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ContractSpecsCompatible( _
                ByVal pContractSpec1 As IContractSpecifier, _
                ByVal pContractSpec2 As IContractSpecifier) As Boolean
Const ProcName As String = "ContractSpecsCompatible"
On Error GoTo Err

ContractSpecsCompatible = gContractSpecsCompatible(pContractSpec1, pContractSpec2)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ContractSpecsEqual( _
                ByVal pContractSpec1 As IContract, _
                ByVal pContractSpec2 As IContract) As Boolean
Const ProcName As String = "ContractSpecsEqual"
On Error GoTo Err

ContractSpecsEqual = gContractSpecsEqual(pContractSpec1, pContractSpec2)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ContractToString(ByVal pContract As IContract) As String
Const ProcName As String = "ContractToString"
On Error GoTo Err

ContractToString = gContractToString(pContract)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ContractToXML(ByVal pContract As IContract) As String
Const ProcName As String = "ContractToXML"
On Error GoTo Err

ContractToXML = gContractToXML(pContract)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateClockFuture( _
                ByVal pContractFuture As IFuture, _
                Optional ByVal pIsSimulated As Boolean, _
                Optional ByVal pClockRate As Single = 1!) As IFuture
Const ProcName As String = "CreateClockFuture"
On Error GoTo Err

Dim lClockFutureBuilder As New ClockFutureBuilder
lClockFutureBuilder.Initialise pContractFuture, pIsSimulated, pClockRate
Set CreateClockFuture = lClockFutureBuilder.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


Public Function CreateComboLeg( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pIsBuyLeg As Boolean, _
                ByVal pRatio As Long) As ComboLeg
Set CreateComboLeg = New ComboLeg
CreateComboLeg.Initialise pContractSpec, pIsBuyLeg, pRatio
End Function


Public Function CreateContractBuilder( _
                ByVal Specifier As IContractSpecifier) As ContractBuilder
Const ProcName As String = "CreateContractBuilder"
On Error GoTo Err

Set CreateContractBuilder = New ContractBuilder
CreateContractBuilder.Initialise Specifier

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateContractBuilderFromContract( _
                ByVal pContract As IContract) As ContractBuilder
Const ProcName As String = "CreateContractBuilderFromContract"
On Error GoTo Err

Set CreateContractBuilderFromContract = New ContractBuilder
CreateContractBuilderFromContract.Initialise pContract.Specifier
CreateContractBuilderFromContract.BuildFrom pContract

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateContractSpecifier( _
                Optional ByVal LocalSymbol As String, _
                Optional ByVal Symbol As String, _
                Optional ByVal Exchange As String, _
                Optional ByVal SecType As SecurityTypes = SecTypeNone, _
                Optional ByVal CurrencyCode As String, _
                Optional ByVal Expiry As String, _
                Optional ByVal Strike As Double, _
                Optional ByVal Right As OptionRights = OptNone) As IContractSpecifier
Const ProcName As String = "CreateContractSpecifier"
On Error GoTo Err

Set CreateContractSpecifier = gCreateContractSpecifier(LocalSymbol, Symbol, Exchange, SecType, CurrencyCode, Expiry, Strike, Right)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateContractFromXML( _
                ByVal xmlString As String) As IContract
Const ProcName As String = "CreateContractFromXML"
On Error GoTo Err

Dim lContract As New Contract
lContract.FromXML xmlString
Set CreateContractFromXML = lContract

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateSessionBuilderFutureFromContractFuture(ByVal pContractFuture As IFuture, ByVal pUseExchangeTimeZone As Boolean) As IFuture
Const ProcName As String = "CreateSessionBuilderFutureFromContractFuture"
On Error GoTo Err

Dim lBuilder As New SessionBuilderFutBldr
lBuilder.Initialise pContractFuture, pUseExchangeTimeZone
Set CreateSessionBuilderFutureFromContractFuture = lBuilder.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetContractSpecKey(ByVal pContractSpec As IContractSpecifier) As String
Const ProcName As String = "GetContractSpecKey"
On Error GoTo Err

GetContractSpecKey = gGetContractSpecKey(pContractSpec)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetCurrencyDescriptor( _
                ByVal CurrencyCode As String) As CurrencyDescriptor
Const ProcName As String = "GetCurrencyDescriptor"
On Error GoTo Err

GetCurrencyDescriptor = gGetCurrencyDescriptor(CurrencyCode)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetCurrencyDescriptors() As CurrencyDescriptor()
Const ProcName As String = "GetCurrencyDescriptors"
On Error GoTo Err

GetCurrencyDescriptors = gGetCurrencyDescriptors

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetExchangeCodes() As String()
Const ProcName As String = "GetExchangeCodes"
On Error GoTo Err

GetExchangeCodes = gGetExchangeCodes

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchContract( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pPriority As TaskPriorities = TaskPriorities.PriorityNormal, _
                Optional ByVal pTaskName As String, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchContract"
On Error GoTo Err

AssertArgument Not pPrimaryContractStore Is Nothing, "pPrimaryContractStore cannot be Nothing"

If IsEmpty(pCookie) Then pCookie = GenerateGUIDString

Dim t As New ContractFetchTask
t.Initialise pContractSpec, pPrimaryContractStore, pSecondaryContractStore, pCookie, Nothing, True
StartTask t, pPriority, pTaskName

Set FetchContract = t.ContractFuture

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchContracts( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pListener As IContractFetchListener, _
                Optional ByVal pPriority As TaskPriorities = TaskPriorities.PriorityNormal, _
                Optional ByVal pTaskName As String, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchContracts"
On Error GoTo Err

AssertArgument Not pPrimaryContractStore Is Nothing, "pPrimaryContractStore cannot be Nothing"

If IsEmpty(pCookie) Then pCookie = GenerateGUIDString

Dim t As New ContractFetchTask
t.Initialise pContractSpec, pPrimaryContractStore, pSecondaryContractStore, pCookie, pListener, False
StartTask t, pPriority, pTaskName

Set FetchContracts = t.ContractsFuture

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPrice( _
                ByVal pPrice As Double, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double) As String
Const ProcName As String = "FormatPrice"
On Error GoTo Err

' see http://www.cmegroup.com/trading/interest-rates/files/TreasuryFuturesPriceRoundingConventions_Mar_24_Final.pdf
' for details of price presentation, especially sections (2) and (7)

FormatPrice = gFormatPrice(pPrice, pSecType, pTickSize)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAs32nds( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs32nds"
On Error GoTo Err

FormatPriceAs32nds = gFormatPriceAs32nds(pPrice)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAs32ndsAndFractions( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs32ndsAndFractions"
On Error GoTo Err

FormatPriceAs32ndsAndFractions = gFormatPriceAs32ndsAndFractions(pPrice)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAs64ths( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs64ths"
On Error GoTo Err

FormatPriceAs64ths = gFormatPriceAs64ths(pPrice)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAs64thsAndFractions( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs64thsAndFractions"
On Error GoTo Err

FormatPriceAs64thsAndFractions = gFormatPriceAs64thsAndFractions(pPrice)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAsDecimals( _
                ByVal pPrice As Double, _
                ByVal pTickSize As Double) As String
Const ProcName As String = "FormatPriceAsDecimals"

On Error GoTo Err

FormatPriceAsDecimals = gFormatPriceAsDecimals(pPrice, pTickSize)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsContractExpired(ByVal pContract As IContract) As Boolean
Const ProcName As String = "IsContractExpired"
On Error GoTo Err

IsContractExpired = gIsContractExpired(pContract)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidCurrencyCode(ByVal CurrencyCode As String) As Boolean
Const ProcName As String = "IsValidCurrencyCode"
On Error GoTo Err

IsValidCurrencyCode = gIsValidCurrencyCode(CurrencyCode)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidExchangeCode(ByVal ExchangeCode As String) As Boolean
Const ProcName As String = "IsValidExchangeCode"
On Error GoTo Err

IsValidExchangeCode = gIsValidExchangeCode(ExchangeCode)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidExpiry( _
                ByVal Value As String) As Boolean
Const ProcName As String = "IsValidExpiry"
On Error GoTo Err

IsValidExpiry = gIsValidExpiry(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidPrice( _
                ByVal pPrice As Double, _
                ByVal pPrevValidPrice As Double, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double) As Boolean
Const ProcName As String = "IsValidPrice"
On Error GoTo Err

IsValidPrice = gIsValidPrice(pPrice, pPrevValidPrice, pSecType, pTickSize)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidSecType( _
                ByVal Value As Long) As Boolean
Const ProcName As String = "IsValidSecType"
On Error GoTo Err

IsValidSecType = gIsValidSecType(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadContractFromConfig(ByVal pConfig As ConfigurationSection) As IContract
Const ProcName As String = "LoadContractFromConfig"
On Error GoTo Err

Set LoadContractFromConfig = gLoadContractFromConfig(pConfig)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadContractSpecFromConfig(ByVal pConfig As ConfigurationSection) As IContractSpecifier
Const ProcName As String = "LoadContractSpecFromConfig"
On Error GoTo Err

Set LoadContractSpecFromConfig = gLoadContractSpecFromConfig(pConfig)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function OptionRightFromString( _
                ByVal Value As String) As OptionRights
Const ProcName As String = "OptionRightFromString"
On Error GoTo Err

OptionRightFromString = gOptionRightFromString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function OptionRightToString( _
                ByVal Value As OptionRights) As String
Const ProcName As String = "OptionRightToString"
On Error GoTo Err

OptionRightToString = gOptionRightToString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePrice( _
                ByVal pPriceString As String, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePrice"
On Error GoTo Err

ParsePrice = gParsePrice(pPriceString, pSecType, pTickSize, pPrice)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAs32nds( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs32nds"
On Error GoTo Err

ParsePriceAs32nds = gParsePriceAs32nds(pPriceString, pPrice)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAs32ndsAndFractions( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs32ndsAndFractions"
On Error GoTo Err

ParsePriceAs32ndsAndFractions = gParsePriceAs32ndsAndFractions(pPriceString, pPrice)
    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAs64ths( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs64ths"
On Error GoTo Err

ParsePriceAs64ths = gParsePriceAs64ths(pPriceString, pPrice)
    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAs64thsAndFractions( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs64thsAndFractions"
On Error GoTo Err

ParsePriceAs64thsAndFractions = gParsePriceAs64thsAndFractions(pPriceString, pPrice)
    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAsDecimals( _
                ByVal pPriceString As String, _
                ByVal pTickSize As Double, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAsDecimals"
On Error GoTo Err

ParsePriceAsDecimals = gParsePriceAsDecimals(pPriceString, pTickSize, pPrice)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub SaveContractSpecToConfig(ByVal pContractSpec As IContractSpecifier, ByVal pConfig As ConfigurationSection)
Const ProcName As String = "SaveContractSpecToConfig"
On Error GoTo Err

gSaveContractSpecToConfig pContractSpec, pConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SaveContractToConfig(ByVal pContract As IContract, ByVal pConfig As ConfigurationSection)
Const ProcName As String = "SaveContractToConfig"
On Error GoTo Err

gSaveContractToConfig pContract, pConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function SecTypeFromString( _
                ByVal Value As String) As SecurityTypes
Const ProcName As String = "SecTypeFromString"
On Error GoTo Err

SecTypeFromString = gSecTypeFromString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SecTypeToShortString( _
                ByVal Value As SecurityTypes) As String
Const ProcName As String = "SecTypeToShortString"
On Error GoTo Err

SecTypeToShortString = gSecTypeToShortString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SecTypeToString( _
                ByVal Value As SecurityTypes) As String
Const ProcName As String = "SecTypeToString"
On Error GoTo Err

SecTypeToString = gSecTypeToString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


