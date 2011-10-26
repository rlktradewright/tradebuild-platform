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

Public Function CreateComboLeg( _
                ByVal pContractSpec As ContractSpecifier, _
                ByVal pIsBuyLeg As Boolean, _
                ByVal pRatio As Long) As ComboLeg
Set CreateComboLeg = New ComboLeg
CreateComboLeg.Initialise pContractSpec, pIsBuyLeg, pRatio
End Function


Public Function CreateContractBuilder( _
                ByVal Specifier As ContractSpecifier) As ContractBuilder
Const ProcName As String = "CreateContractBuilder"
Dim failpoint As String
On Error GoTo Err

Set CreateContractBuilder = New ContractBuilder
CreateContractBuilder.Initialise Specifier

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function CreateContractBuilderFromContract( _
                ByVal pContract As Contract) As ContractBuilder
Const ProcName As String = "CreateContractBuilderFromContract"
Dim failpoint As String
On Error GoTo Err

Set CreateContractBuilderFromContract = New ContractBuilder
CreateContractBuilderFromContract.Initialise pContract.Specifier
CreateContractBuilderFromContract.BuildFrom pContract

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function CreateContractsBuilder( _
                ByVal Specifier As ContractSpecifier) As ContractsBuilder
Const ProcName As String = "CreateContractsBuilder"
Dim failpoint As String
On Error GoTo Err

Set CreateContractsBuilder = New ContractsBuilder
CreateContractsBuilder.Initialise Specifier

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function CreateContractSpecifier( _
                Optional ByVal LocalSymbol As String, _
                Optional ByVal Symbol As String, _
                Optional ByVal Exchange As String, _
                Optional ByVal SecType As SecurityTypes = SecTypeNone, _
                Optional ByVal CurrencyCode As String, _
                Optional ByVal Expiry As String, _
                Optional ByVal Strike As Double, _
                Optional ByVal Right As OptionRights = OptNone) As ContractSpecifier
Const ProcName As String = "CreateContractSpecifier"
Dim failpoint As String
On Error GoTo Err

Set CreateContractSpecifier = gCreateContractSpecifier(LocalSymbol, Symbol, Exchange, SecType, CurrencyCode, Expiry, Strike, Right)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function CreateContractFromXML( _
                ByVal xmlString As String) As Contract
Const ProcName As String = "CreateContractFromXML"
Dim failpoint As String
On Error GoTo Err

Set CreateContractFromXML = New Contract
CreateContractFromXML.FromXML xmlString

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function GetCurrencyDescriptor( _
                ByVal CurrencyCode As String) As CurrencyDescriptor
Const ProcName As String = "GetCurrencyDescriptor"
Dim failpoint As String
On Error GoTo Err

GetCurrencyDescriptor = gGetCurrencyDescriptor(CurrencyCode)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function GetCurrencyDescriptors() As CurrencyDescriptor()
Const ProcName As String = "GetCurrencyDescriptors"
Dim failpoint As String
On Error GoTo Err

GetCurrencyDescriptors = gGetCurrencyDescriptors

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function GetExchangeCodes() As String()
Const ProcName As String = "GetExchangeCodes"
Dim failpoint As String
On Error GoTo Err

GetExchangeCodes = gGetExchangeCodes

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function IsValidCurrencyCode(ByVal CurrencyCode As String) As Boolean
Const ProcName As String = "IsValidCurrencyCode"
Dim failpoint As String
On Error GoTo Err

IsValidCurrencyCode = gIsValidCurrencyCode(CurrencyCode)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function IsValidExchangeCode(ByVal ExchangeCode As String) As Boolean
Const ProcName As String = "IsValidExchangeCode"
Dim failpoint As String
On Error GoTo Err

IsValidExchangeCode = gIsValidExchangeCode(ExchangeCode)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function IsValidExpiry( _
                ByVal value As String) As Boolean
Const ProcName As String = "IsValidExpiry"
Dim failpoint As String
On Error GoTo Err

IsValidExpiry = gIsValidExpiry(value)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
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
                ByVal value As Long) As Boolean
Const ProcName As String = "IsValidSecType"
Dim failpoint As String
On Error GoTo Err

IsValidSecType = gIsValidSecType(value)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function OptionRightFromString( _
                ByVal value As String) As OptionRights
Const ProcName As String = "OptionRightFromString"
Dim failpoint As String
On Error GoTo Err

OptionRightFromString = gOptionRightFromString(value)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function OptionRightToString( _
                ByVal value As OptionRights) As String
Const ProcName As String = "OptionRightToString"
Dim failpoint As String
On Error GoTo Err

OptionRightToString = gOptionRightToString(value)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function SecTypeFromString( _
                ByVal value As String) As SecurityTypes
Const ProcName As String = "SecTypeFromString"
Dim failpoint As String
On Error GoTo Err

SecTypeFromString = gSecTypeFromString(value)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function SecTypeToShortString( _
                ByVal value As SecurityTypes) As String
Const ProcName As String = "SecTypeToShortString"
Dim failpoint As String
On Error GoTo Err

SecTypeToShortString = gSecTypeToShortString(value)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function SecTypeToString( _
                ByVal value As SecurityTypes) As String
Const ProcName As String = "SecTypeToString"
Dim failpoint As String
On Error GoTo Err

SecTypeToString = gSecTypeToString(value)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

'@================================================================================
' Helper Functions
'@================================================================================


