VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractTickerMap"
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

Private Const ModuleName                            As String = "ContractTickerMap"

'@================================================================================
' Member variables
'@================================================================================

Private mTickerIndexes                              As EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mTickerIndexes = New EnumerableCollection
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
                ByVal pContract As TwsContract, _
                ByVal pTickerId As Long)
Const ProcName As String = "Add"
On Error GoTo Err

mTickerIndexes.Add pTickerId, pContract.ToString

Exit Sub

Err:
If Err.number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Function Contains( _
                ByVal pContract As TwsContract) As Boolean
Const ProcName As String = "Remove"
On Error GoTo Err

Contains = mTickerIndexes.Contains(pContract.ToString)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function Item( _
                ByVal pContract As TwsContract) As Long
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
On Error GoTo Err

Item = mTickerIndexes.Item(pContract.ToString)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mTickerIndexes.[NewEnum]
End Function


Friend Sub Remove( _
                ByVal pContract As TwsContract)
Const ProcName As String = "Remove"
On Error GoTo Err

mTickerIndexes.Remove pContract.ToString

Exit Sub

Err:
If Err.number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




