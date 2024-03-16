VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractsBuilder"
Attribute VB_GlobalNameSpace = False
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

Implements IContractsBuilder

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "ContractsBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mContracts                          As Contracts

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mContracts = New Contracts
End Sub

'@================================================================================
' IContractsBuilder Interface Members
'@================================================================================

Private Sub IContractsBuilder_Add(ByVal pContract As IContract)
Const ProcName As String = "IContractsBuilder_Add"
On Error GoTo Err

Add pContract

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IContractsBuilder_Contracts() As IContracts
Set IContractsBuilder_Contracts = Contracts
End Property

Private Property Let IContractsBuilder_ContractSpecifier(ByVal Value As IContractSpecifier)
ContractSpecifier = Value
End Property

Private Property Let IContractsBuilder_SortAscending(ByVal Value As Boolean)
SortAscending = Value
End Property

Private Property Get IContractsBuilder_SortAscending() As Boolean
IContractsBuilder_SortAscending = SortAscending
End Property

Private Property Let IContractsBuilder_SortKeys(ByRef Value() As ContractSortKeyIds)
Dim ar() As ContractSortKeyIds
ar = Value

SortKeys = ar
End Property

Private Property Get IContractsBuilder_SortKeys() As ContractSortKeyIds()
IContractsBuilder_SortKeys = SortKeys
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let ContractSpecifier(ByVal Value As IContractSpecifier)
mContracts.ContractSpecifier = Value
End Property

Public Property Get Contracts() As Contracts
Set Contracts = mContracts
End Property

Public Property Let SortAscending(ByVal Value As Boolean)
mContracts.SortAscending = Value
End Property

Public Property Get SortAscending() As Boolean
SortAscending = mContracts.SortAscending
End Property

Public Property Let SortKeys(ByRef Value() As ContractSortKeyIds)
Const ProcName As String = "SortKeys"
On Error GoTo Err

Dim ar() As ContractSortKeyIds
ar = Value

mContracts.SortKeys = ar

Exit Property

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SortKeys() As ContractSortKeyIds()
SortKeys = mContracts.SortKeys
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Add(ByVal pContract As IContract)
Const ProcName As String = "Add"
On Error GoTo Err

mContracts.Add pContract

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


