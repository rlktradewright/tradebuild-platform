VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IContractsBuilder"
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

Private Const ModuleName                    As String = "IContractsBuilder"

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

Public Property Get Contracts() As IContracts
End Property

Public Property Let ContractSpecifier(ByVal Value As IContractSpecifier)
End Property

Public Property Let SortAscending(ByVal Value As Boolean)
End Property

Public Property Get SortAscending() As Boolean
End Property

Public Property Let SortKeys( _
                ByRef Value() As ContractSortKeyIds)
End Property
                
Public Property Get SortKeys() As ContractSortKeyIds()
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Add(ByVal pContract As IContract)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




