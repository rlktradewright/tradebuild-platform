VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractComparable"
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

Implements IComparable

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

Private Const ModuleName                            As String = "ContractComparable"

'@================================================================================
' Member variables
'@================================================================================

Private mSortKeys()                                 As ContractSortKeyIds
Private mContractSpec                               As IContractSpecifier
Private mAscending                                  As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IComparable Interface Members
'@================================================================================

Private Function IComparable_CompareTo(ByVal pObj As IComparable) As Long
Const ProcName As String = "Comparable_CompareTo"
On Error GoTo Err

Assert TypeOf pObj Is ContractComparable, "Type of pobj is not ContractComparable"

Dim lComparable As ContractComparable
Set lComparable = pObj
IComparable_CompareTo = gContractSpecsCompare( _
                                    mContractSpec, _
                                    lComparable.ContractSpec, _
                                    mSortKeys, _
                                    mAscending)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ContractSpec() As IContractSpecifier
Set ContractSpec = mContractSpec
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractSpec As IContractSpecifier, _
                ByRef pSortkeys() As ContractSortKeyIds, _
                ByVal pAscending As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractSpec = pContractSpec
mSortKeys = pSortkeys
mAscending = pAscending

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
'@================================================================================
' Helper Functions
'@================================================================================




