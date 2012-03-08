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
    ContractStoreCanStore = 1
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

'@================================================================================
' Methods
'@================================================================================

Public Function ContractsCompare( _
                ByVal pContract1 As IContract, _
                ByVal pContract2 As IContract, _
                ByRef pSortKeys() As ContractSortKeyIds) As Long
Const ProcName As String = "ContractsCompare"
On Error GoTo Err

ContractsCompare = gContractsCompare(pContract1, pContract2, pSortKeys)

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

Public Function CreateContractsBuilder( _
                ByVal Specifier As IContractSpecifier) As ContractsBuilder
Const ProcName As String = "CreateContractsBuilder"

On Error GoTo Err

Set CreateContractsBuilder = New ContractsBuilder
CreateContractsBuilder.Initialise Specifier

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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


