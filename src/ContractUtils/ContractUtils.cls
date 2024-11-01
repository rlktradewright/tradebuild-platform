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
    ContractSortKeyMultiplier
    ContractSortKeyCurrency
    ContractSortKeyRight
    ContractSortKeyStrike
End Enum

Public Enum ContractStoreCapabilities
    'ContractStoreCanStore = 1
    ContractStoreCanQuery = 2
    ContractStoreOptionExpiries = 4
    ContractStoreOptionStrikes = 8
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
    SecTypeWarrant
    SecTypeCFD
    SecTypeCrypto
End Enum

'@================================================================================
' Types
'@================================================================================

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

Public Property Get MaxContractExpiryOffset() As Long
MaxContractExpiryOffset = GContractUtils.MaxContractExpiryOffset
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
                ByRef pSortkeys() As ContractSortKeyIds, _
                Optional ByVal pDescending As Boolean) As Long
Const ProcName As String = "ContractSpecsCompare"
On Error GoTo Err

ContractSpecsCompare = GContractUtils.ContractSpecsCompare( _
                                pContractSpec1, _
                                pContractSpec2, _
                                pSortkeys, _
                                Not pDescending)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ContractSpecsCompatible( _
                ByVal pContractSpec1 As IContractSpecifier, _
                ByVal pContractSpec2 As IContractSpecifier) As Boolean
Const ProcName As String = "ContractSpecsCompatible"
On Error GoTo Err

ContractSpecsCompatible = GContractUtils.ContractSpecsCompatible(pContractSpec1, pContractSpec2)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ContractSpecsEqual( _
                ByVal pContractSpec1 As IContractSpecifier, _
                ByVal pContractSpec2 As IContractSpecifier) As Boolean
Const ProcName As String = "ContractSpecsEqual"
On Error GoTo Err

ContractSpecsEqual = GContractUtils.ContractSpecsEqual(pContractSpec1, pContractSpec2)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ContractToString(ByVal pContract As IContract) As String
Const ProcName As String = "ContractToString"
On Error GoTo Err

ContractToString = GContractUtils.ContractToString(pContract)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ContractToXML(ByVal pContract As IContract) As String
Const ProcName As String = "ContractToXML"
On Error GoTo Err

ContractToXML = GContractUtils.ContractToXML(pContract)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateClockFuture( _
                ByVal pContractFuture As IFuture, _
                Optional ByVal pIsSimulated As Boolean, _
                Optional ByVal pClockRate As Single = 1!) As IFuture
Const ProcName As String = "CreateClockFuture"
On Error GoTo Err

Set CreateClockFuture = GContractUtils.CreateClockFuture(pContractFuture, pIsSimulated, pClockRate)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function


Public Function CreateComboLeg( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pIsBuyLeg As Boolean, _
                ByVal pRatio As Long) As ComboLeg
Set CreateComboLeg = GContractUtils.CreateComboLeg(pContractSpec, pIsBuyLeg, pRatio)
End Function


Public Function CreateContractBuilder( _
                ByVal Specifier As IContractSpecifier) As ContractBuilder
Const ProcName As String = "CreateContractBuilder"
On Error GoTo Err

Set CreateContractBuilder = GContractUtils.CreateContractBuilder(Specifier)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateContractBuilderFromContract( _
                ByVal pContract As IContract) As ContractBuilder
Const ProcName As String = "CreateContractBuilderFromContract"
On Error GoTo Err

Set CreateContractBuilderFromContract = GContractUtils.CreateContractBuilderFromContract(pContract)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateContractSpecifier( _
                Optional ByVal LocalSymbol As String, _
                Optional ByVal Symbol As String, _
                Optional ByVal TradingClass As String, _
                Optional ByVal Exchange As String, _
                Optional ByVal SecType As SecurityTypes = SecTypeNone, _
                Optional ByVal CurrencyCode As String, _
                Optional ByVal Expiry As String, _
                Optional ByVal Multiplier As Double = 0#, _
                Optional ByVal Strike As Double, _
                Optional ByVal Right As OptionRights = OptNone) As IContractSpecifier
Const ProcName As String = "CreateContractSpecifier"
On Error GoTo Err

Set CreateContractSpecifier = GContractUtils.CreateContractSpecifier( _
                                    LocalSymbol, _
                                    Symbol, _
                                    TradingClass, _
                                    Exchange, _
                                    SecType, _
                                    CurrencyCode, _
                                    Expiry, _
                                    Multiplier, _
                                    Strike, _
                                    Right)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateContractSpecifierFromString( _
                ByVal SpecString As String) As IContractSpecifier
Const ProcName As String = "CreateContractSpecifierFromString"
On Error GoTo Err

Set CreateContractSpecifierFromString = GContractUtils.CreateContractSpecifierFromString(SpecString)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateContractFromXML( _
                ByVal xmlString As String) As IContract
Const ProcName As String = "CreateContractFromXML"
On Error GoTo Err

Set CreateContractFromXML = GContractUtils.CreateContractFromXML(xmlString)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateSessionBuilderFutureFromContractFuture( _
                ByVal pContractFuture As IFuture, _
                ByVal pUseExchangeTimeZone As Boolean, _
                ByVal pUseFullSession As Boolean) As IFuture
Const ProcName As String = "CreateSessionBuilderFutureFromContractFuture"
On Error GoTo Err

Set CreateSessionBuilderFutureFromContractFuture = GContractUtils.CreateSessionBuilderFutureFromContractFuture( _
                                                        pContractFuture, _
                                                        pUseExchangeTimeZone, _
                                                        pUseFullSession)
                                                        

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetContractSpecKey(ByVal pContractSpec As IContractSpecifier) As String
Const ProcName As String = "GetContractSpecKey"
On Error GoTo Err

GetContractSpecKey = GContractUtils.GetContractSpecKey(pContractSpec)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetExchangeCodes() As String()
Const ProcName As String = "GetExchangeCodes"
On Error GoTo Err

GetExchangeCodes = GContractUtils.GetExchangeCodes

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
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

Set FetchContract = GContractUtils.FetchContracts(pContractSpec, _
                                True, _
                                pPrimaryContractStore, _
                                pSecondaryContractStore, _
                                Nothing, _
                                pPriority, _
                                pTaskName, _
                                pCookie)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
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

Set FetchContracts = GContractUtils.FetchContracts(pContractSpec, _
                                False, _
                                pPrimaryContractStore, _
                                pSecondaryContractStore, _
                                pListener, _
                                pPriority, _
                                pTaskName, _
                                pCookie)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchContractsSorted( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByRef pSortkeys() As ContractSortKeyIds, _
                Optional ByVal pSortDescending As Boolean = False, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pPriority As TaskPriorities = TaskPriorities.PriorityNormal, _
                Optional ByVal pTaskName As String, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchContractsSorted"
On Error GoTo Err

Set FetchContractsSorted = GContractUtils.FetchContractsSorted( _
                                pContractSpec, _
                                pPrimaryContractStore, _
                                pSortkeys, _
                                pSortDescending, _
                                pSecondaryContractStore, _
                                pPriority, _
                                pTaskName, _
                                pCookie)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchOptionExpiries( _
                ByVal pUnderlyingContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pStrike As Double = 0#, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pPriority As TaskPriorities = TaskPriorities.PriorityNormal, _
                Optional ByVal pTaskName As String, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchOptionExpiries"
On Error GoTo Err

Set FetchOptionExpiries = GContractUtils.FetchOptionExpiries( _
                                    pUnderlyingContractSpecifier, _
                                    pExchange, _
                                    pPrimaryContractStore, _
                                    pStrike, _
                                    pSecondaryContractStore, _
                                    pPriority, _
                                    pTaskName, _
                                    pCookie)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchOptionStrikes( _
                ByVal pUnderlyingContractSpecifier, _
                ByVal pExchange As String, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pExpiry As String, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pPriority As TaskPriorities = TaskPriorities.PriorityNormal, _
                Optional ByVal pTaskName As String, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchOptionStrikes"
On Error GoTo Err

Set FetchOptionStrikes = GContractUtils.FetchOptionStrikes( _
                                    pUnderlyingContractSpecifier, _
                                    pExchange, _
                                    pPrimaryContractStore, _
                                    pExpiry, _
                                    pSecondaryContractStore, _
                                    pPriority, _
                                    pTaskName, _
                                    pCookie)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
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
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAs32nds( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs32nds"
On Error GoTo Err

FormatPriceAs32nds = gFormatPriceAs32nds(pPrice)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAs32ndsAndFractions( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs32ndsAndFractions"
On Error GoTo Err

FormatPriceAs32ndsAndFractions = gFormatPriceAs32ndsAndFractions(pPrice)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAs64ths( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs64ths"
On Error GoTo Err

FormatPriceAs64ths = gFormatPriceAs64ths(pPrice)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAs64thsAndFractions( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs64thsAndFractions"
On Error GoTo Err

FormatPriceAs64thsAndFractions = gFormatPriceAs64thsAndFractions(pPrice)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAsDecimals( _
                ByVal pPrice As Double, _
                ByVal pTickSize As Double) As String
Const ProcName As String = "FormatPriceAsDecimals"

On Error GoTo Err

FormatPriceAsDecimals = gFormatPriceAsDecimals(pPrice, pTickSize)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsContractExpired( _
                ByVal pContract As IContract, _
                Optional ByVal pClock As Clock) As Boolean
Const ProcName As String = "IsContractExpired"
On Error GoTo Err

IsContractExpired = GContractUtils.IsContractExpired(pContract, pClock)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsContractSpecOffsetExpiry( _
                ByVal pContractSpec As IContractSpecifier, _
                Optional ByRef pErrorMessage) As Boolean
IsContractSpecOffsetExpiry = GContractUtils.IsContractSpecOffsetExpiry( _
                                    pContractSpec, _
                                    pErrorMessage)
End Function

Public Function IsValidExchangeCode(ByVal ExchangeCode As String) As Boolean
Const ProcName As String = "IsValidExchangeCode"
On Error GoTo Err

IsValidExchangeCode = GContractUtils.IsValidExchangeCode(ExchangeCode)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidExpiry( _
                ByVal Value As String, _
                Optional ByRef pErrorMessage As String) As Boolean
Const ProcName As String = "IsValidExpiry"
On Error GoTo Err

IsValidExpiry = GContractUtils.IsValidExpiry(Value, pErrorMessage)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidPrice( _
                ByVal pPrice As Double, _
                ByVal pPrevValidPrice As Double, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double) As Boolean
Const ProcName As String = "IsValidPrice"
On Error GoTo Err

IsValidPrice = GContractUtils.IsValidPrice(pPrice, pPrevValidPrice, pSecType, pTickSize)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidSecType( _
                ByVal Value As Long) As Boolean
Const ProcName As String = "IsValidSecType"
On Error GoTo Err

IsValidSecType = GContractUtils.IsValidSecType(Value)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadContractFromConfig(ByVal pConfig As ConfigurationSection) As IContract
Const ProcName As String = "LoadContractFromConfig"
On Error GoTo Err

Set LoadContractFromConfig = GContractUtils.LoadContractFromConfig(pConfig)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadContractSpecFromConfig(ByVal pConfig As ConfigurationSection) As IContractSpecifier
Const ProcName As String = "LoadContractSpecFromConfig"
On Error GoTo Err

Set LoadContractSpecFromConfig = GContractUtils.LoadContractSpecFromConfig(pConfig)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OptionRightFromString( _
                ByVal Value As String) As OptionRights
Const ProcName As String = "OptionRightFromString"
On Error GoTo Err

OptionRightFromString = GContractUtils.OptionRightFromString(Value)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OptionRightToString( _
                ByVal Value As OptionRights) As String
Const ProcName As String = "OptionRightToString"
On Error GoTo Err

OptionRightToString = GContractUtils.OptionRightToString(Value)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub ParseOffsetExpiry( _
                ByVal Value As String, _
                ByRef pExpiryOffset As Long, _
                ByRef pDaysBeforeExpiryToSwitch As Long)
Const ProcName As String = "ParseOffsetExpiry"
On Error GoTo Err

GContractUtils.ParseOffsetExpiry Value, _
                            pExpiryOffset, _
                            pDaysBeforeExpiryToSwitch

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

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
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAs32nds( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs32nds"
On Error GoTo Err

ParsePriceAs32nds = gParsePriceAs32nds(pPriceString, pPrice)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAs32ndsAndFractions( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs32ndsAndFractions"
On Error GoTo Err

ParsePriceAs32ndsAndFractions = gParsePriceAs32ndsAndFractions(pPriceString, pPrice)
    
Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAs64ths( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs64ths"
On Error GoTo Err

ParsePriceAs64ths = gParsePriceAs64ths(pPriceString, pPrice)
    
Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAs64thsAndFractions( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs64thsAndFractions"
On Error GoTo Err

ParsePriceAs64thsAndFractions = gParsePriceAs64thsAndFractions(pPriceString, pPrice)
    
Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
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
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub SaveContractSpecToConfig(ByVal pContractSpec As IContractSpecifier, ByVal pConfig As ConfigurationSection)
Const ProcName As String = "SaveContractSpecToConfig"
On Error GoTo Err

GContractUtils.SaveContractSpecToConfig pContractSpec, pConfig

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SaveContractToConfig(ByVal pContract As IContract, ByVal pConfig As ConfigurationSection)
Const ProcName As String = "SaveContractToConfig"
On Error GoTo Err

GContractUtils.SaveContractToConfig pContract, pConfig

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function SecTypeFromString( _
                ByVal Value As String) As SecurityTypes
Const ProcName As String = "SecTypeFromString"
On Error GoTo Err

SecTypeFromString = GContractUtils.SecTypeFromString(Value)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function SecTypeToShortString( _
                ByVal Value As SecurityTypes) As String
Const ProcName As String = "SecTypeToShortString"
On Error GoTo Err

SecTypeToShortString = GContractUtils.SecTypeToShortString(Value)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function SecTypeToString( _
                ByVal Value As SecurityTypes) As String
Const ProcName As String = "SecTypeToString"
On Error GoTo Err

SecTypeToString = GContractUtils.SecTypeToString(Value)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function TryCreateContractSpecifierFromString( _
                ByVal SpecString As String, _
                ByRef ContractSpec As IContractSpecifier, _
                Optional ByRef ErrorMessage As String) As Boolean
Const ProcName As String = "CreateContractSpecifierFromString"
On Error GoTo Err


Set ContractSpec = GContractUtils.CreateContractSpecifierFromString(SpecString)

Exit Function

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    ErrorMessage = Err.Description
    TryCreateContractSpecifierFromString = False
Else
    GContracts.HandleUnexpectedError ProcName, ModuleName
End If
End Function

Public Function TryParseOffsetExpiry( _
                ByVal Value As String, _
                ByRef pExpiryOffset As Long, _
                ByRef pDaysBeforeExpiryToSwitch As Long, _
                Optional ByRef pErrorMessage As String) As Boolean
Const ProcName As String = "TryParseOffsetExpiry"
On Error GoTo Err

TryParseOffsetExpiry = GContractUtils.TryParseOffsetExpiry(Value, _
                                    pExpiryOffset, _
                                    pDaysBeforeExpiryToSwitch, _
                                    pErrorMessage)
                                    
Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


