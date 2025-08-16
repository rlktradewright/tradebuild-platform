VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsContracts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "TwsContracts"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpecifier              As TwsContractSpecifier

Private mContracts                      As SortedDictionary

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mContracts = CreateSortedDictionary(KeyTypeString)
End Sub

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Count() As Long
Const ProcName As String = "Count"
On Error GoTo Err

Count = mContracts.Count

Exit Property

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Property

Public Property Get Enumerator() As Enumerator
Set Enumerator = mContracts.Enumerator
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add( _
                ByVal pContract As TwsContract) As Boolean
Const ProcName As String = "Add"
On Error GoTo Err

Dim lKey As String: lKey = (pContract.key)
If mContracts.Contains(lKey) Then Exit Function

mContracts.Add pContract, lKey
Add = True

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function Contains(ByVal pKey As TwsContractSpecifier) As Boolean
Const ProcName As String = "Contains"
On Error GoTo Err

Contains = mContracts.Contains(pKey.key)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function Item(ByVal pKey As TwsContractSpecifier) As TwsContract
Const ProcName As String = "Item"
On Error GoTo Err

Set Item = mContracts.Item(pKey.key)

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function ItemAtIndex(ByVal pIndex As Long) As TwsContract
Const ProcName As String = "ItemAtIndex"
On Error GoTo Err

AssertArgument pIndex > 0 And pIndex <= mContracts.Count, "Invalid index"

Dim en As Enumerator
Set en = Enumerator

Dim i As Long
Do While en.MoveNext
    i = i + 1
    If i = pIndex Then
        Set ItemAtIndex = en.Current
        Exit Function
    End If
Loop

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mContracts.Enumerator
End Function

Public Function ToString() As String
Const ProcName As String = "ToString"
On Error GoTo Err


Dim en As Enumerator
Set en = Enumerator
Do While en.MoveNext
    Dim lContract As TwsContract
    Set lContract = en.Current
    ToString = ToString & lContract.Specifier.ToString & vbCrLf
Loop


Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

