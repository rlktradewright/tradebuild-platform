VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "ContractMap"

'@================================================================================
' Member variables
'@================================================================================

Private mIndexes                                    As EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mIndexes = New EnumerableCollection
End Sub

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

Friend Sub Add( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pIndex As Long)
Const ProcName As String = "Add"
On Error GoTo Err

mIndexes.Add pIndex, pContractSpec.Key

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function Contains( _
                ByVal pContractSpec As IContractSpecifier) As Boolean
Const ProcName As String = "Contains"
On Error GoTo Err

Contains = mIndexes.Contains(pContractSpec.Key)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Item( _
                ByVal pContractSpec As IContractSpecifier) As Long
Const ProcName As String = "Item"
On Error GoTo Err

Item = mIndexes.Item(pContractSpec.Key)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mIndexes.[NewEnum]
End Function


Friend Sub Remove( _
                ByVal pContractSpec As IContractSpecifier)
Const ProcName As String = "Remove"
On Error GoTo Err

mIndexes.Remove pContractSpec.Key

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


