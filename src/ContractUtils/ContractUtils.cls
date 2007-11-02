VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalClass"
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

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "ContractUtils26"
Private Const ModuleName                    As String = "GlobalClass"

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

Public Function CreateContractBuilder( _
                ByVal specifier As ContractSpecifier) As ContractBuilder
Set CreateContractBuilder = New ContractBuilder
CreateContractBuilder.initialise specifier
End Function

Public Function CreateContractBuilderFromContract( _
                ByVal pContract As Contract) As ContractBuilder
Set CreateContractBuilderFromContract = New ContractBuilder
CreateContractBuilderFromContract.initialise pContract.specifier
CreateContractBuilderFromContract.buildFrom pContract
End Function

Public Function CreateContractsBuilder( _
                ByVal specifier As ContractSpecifier) As ContractsBuilder
Set CreateContractsBuilder = New ContractsBuilder
CreateContractsBuilder.initialise specifier
End Function

Public Function CreateContractSpecifier( _
                Optional ByVal localSymbol As String, _
                Optional ByVal symbol As String, _
                Optional ByVal exchange As String, _
                Optional ByVal sectype As SecurityTypes = SecTypeNone, _
                Optional ByVal currencyCode As String, _
                Optional ByVal expiry As String, _
                Optional ByVal strike As Double, _
                Optional ByVal right As OptionRights = OptNone) As ContractSpecifier

If localSymbol = "" And symbol = "" Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "CreateContractSpecifier", _
            "Symbol must be supplied if localSymbol is not supplied"
End If

If exchange <> "" And _
    Not IsValidExchangeCode(exchange) _
Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "CreateContractSpecifier", _
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
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & "CreateContractSpecifier", _
                "strike must be > 0"
    End If
    Select Case right
    Case OptCall
    Case OptPut
    Case OptNone
    Case Else
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & "CreateContractSpecifier", _
                "'" & right & "' is not a valid option right"
    End Select
Case SecTypeCash
Case SecTypeCombo
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "CreateContractSpecifier", _
            "Sectype 'combo' is not permissible"
Case SecTypeIndex
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "CreateContractSpecifier", _
            "'" & sectype & "' is not a valid secType"
End Select

Set CreateContractSpecifier = New ContractSpecifier
With CreateContractSpecifier
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

Public Function CreateContractFromXML( _
                ByVal xmlString As String) As Contract
Set CreateContractFromXML = New Contract
CreateContractFromXML.FromXML xmlString
End Function

Public Function GetExchangeCodes() As String()
GetExchangeCodes = gGetExchangeCodes
End Function

Public Function IsValidExchangeCode(ByVal code As String) As Boolean
IsValidExchangeCode = gIsValidExchangeCode(code)
End Function

Public Function OptionRightFromString( _
                ByVal value As String) As OptionRights
OptionRightFromString = gOptionRightFromString(value)
End Function

Public Function OptionRightToString( _
                ByVal value As OptionRights) As String
OptionRightToString = gOptionRightToString(value)
End Function

Public Function SecTypeFromString( _
                ByVal value As String) As SecurityTypes
SecTypeFromString = gSecTypeFromString(value)
End Function

Public Function SecTypeToShortString( _
                ByVal value As SecurityTypes) As String
SecTypeToShortString = gSecTypeToShortString(value)
End Function

Public Function SecTypeToString( _
                ByVal value As SecurityTypes) As String
SecTypeToString = gSecTypeToString(value)
End Function

'@================================================================================
' Helper Functions
'@================================================================================


