VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IContracts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
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

Private Const ModuleName                            As String = "IContracts"

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

Public Property Get Count() As Long
End Property

Public Property Get Enumerator() As Enumerator
End Property

Public Property Let SortKeys( _
                ByRef Value() As ContractSortKeyIds)
End Property
                
Public Property Get SortKeys() As ContractSortKeyIds()
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddCollectionChangeListener( _
                ByVal pListener As ICollectionChangeListener)
End Sub

Public Function Contains(ByVal pKey As IContractSpecifier) As Boolean
End Function

Public Function ItemAtIndex(ByVal pIndex As Long) As IContract
End Function

Public Function Item(ByVal pKey As IContractSpecifier) As IContract
Attribute Item.VB_UserMemId = 0
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
End Function

Public Sub RemoveCollectionChangeListener( _
                ByVal pListener As ICollectionChangeListener)
End Sub

Public Function ToString() As String
End Function

'@================================================================================
' Helper Functions
'@================================================================================




