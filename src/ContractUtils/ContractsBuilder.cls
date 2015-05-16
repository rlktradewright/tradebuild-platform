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

Private mContracts                          As New Contracts

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IContractsBuilder Interface Members
'@================================================================================

Private Sub IContractsBuilder_Add(ByVal pContract As IContract)
Const ProcName As String = "IContractsBuilder_Add"
On Error GoTo Err

'AssertArgument gContractSpecsCompatible(mContracts.ContractSpecifier, pContract.Specifier), "Contract not compatible with contract spec"
Add pContract

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IContractsBuilder_Contracts() As IContracts
Set IContractsBuilder_Contracts = Contracts
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Contracts() As Contracts
Set Contracts = mContracts
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Friend Sub Initialise( _
'                ByVal ContractSpec As IContractSpecifier)
'Const ProcName As String = "Initialise"
'On Error GoTo Err
'
'Set mContracts = New Contracts
'mContracts.Initialise ContractSpec
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

'@================================================================================
' Helper Functions
'@================================================================================


