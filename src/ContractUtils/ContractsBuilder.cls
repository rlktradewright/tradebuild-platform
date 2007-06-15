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

Private Const ProjectName                   As String = "ContractUtils26"
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

Public Sub AddContract(ByVal pContract As Contract)
mContracts.Add pContract
End Sub

Public Sub ContractSpecifierInvalid(ByVal reason As String)
mContracts.ContractSpecifierInvalid reason
End Sub

Friend Sub initialise( _
                ByVal contractSpec As ContractSpecifier)
Set mContracts = New Contracts
mContracts.initialise contractSpec
End Sub

Public Sub LoadComplete()
mContracts.LoadComplete
End Sub

Public Sub notify( _
                ByRef ev As NotificationEvent)
mContracts.notify ev
End Sub

Public Sub remove(ByVal index As Variant)
mContracts.remove index
End Sub

Public Sub unableToLoad( _
                ByVal reason As String, _
                ByVal serviceProviderName As String)
mContracts.serviceProviderCantHandle reason, serviceProviderName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


