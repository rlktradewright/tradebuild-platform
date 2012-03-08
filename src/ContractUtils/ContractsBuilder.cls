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
Attribute VB_Creatable = False
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

'@================================================================================
' XXXX Interface Members
'@================================================================================

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

Public Sub AddContract(ByVal pContract As IContract)
Const ProcName As String = "AddContract"

On Error GoTo Err

mContracts.Add pContract

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Public Sub ContractSpecifierInvalid(ByVal reason As String)
'Const ProcName As String = "ContractSpecifierInvalid"
'
'On Error GoTo Err
'
'mContracts.ContractSpecifierInvalid reason
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

Friend Sub Initialise( _
                ByVal ContractSpec As IContractSpecifier)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mContracts = New Contracts
mContracts.Initialise ContractSpec

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Public Sub LoadComplete()
'Const ProcName As String = "LoadComplete"
'
'On Error GoTo Err
'
'mContracts.LoadComplete
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

'Public Sub Notify( _
'                ByRef ev As NotificationEventData)
'Const ProcName As String = "Notify"
'
'On Error GoTo Err
'
'mContracts.Notify ev
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

Public Sub Remove(ByVal index As Variant)
Const ProcName As String = "Remove"

On Error GoTo Err

mContracts.Remove index

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Public Sub UnableToLoad( _
'                ByVal reason As String, _
'                ByVal serviceProviderName As String)
'Const ProcName As String = "UnableToLoad"
'
'On Error GoTo Err
'
'mContracts.ServiceProviderCantHandle reason, serviceProviderName
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

'@================================================================================
' Helper Functions
'@================================================================================


