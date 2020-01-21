VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Groups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Private Const ModuleName                            As String = "Groups"

'@================================================================================
' Member variables
'@================================================================================

Private mDict                                       As SortedDictionary

Private mContractStore                              As IContractStore

Private mMarketDataManager                          As IMarketDataManager
Private mOrderManager                               As OrderManager
Private mScopeName                                  As String
Private mOrderSubmitterFactory                      As IOrderSubmitterFactory

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mDict = CreateSortedDictionary(KeyTypeString)
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

Friend Property Get Enumerator() As Enumerator
Set Enumerator = mDict.Enumerator
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Add(ByVal pGroupName As String) As GroupResources
Const ProcName As String = "Add"
On Error GoTo Err

Dim lResources As GroupResources

If mDict.TryItem(UCase$(pGroupName), lResources) Then
    Set Add = lResources
    Exit Function
End If

Set lResources = New GroupResources
lResources.Initialise pGroupName, _
                        mContractStore, _
                        mMarketDataManager, _
                        mOrderManager, _
                        mScopeName, _
                        mOrderSubmitterFactory

mDict.Add lResources, UCase$(pGroupName)

Set Add = lResources

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Clear()
mDict.Clear
End Sub

Friend Function Contains(ByVal pGroupName As String) As Boolean
Contains = mDict.Contains(UCase$(pGroupName))
End Function

Friend Function Count() As Long
Count = mDict.Count
End Function

Friend Sub Initialise( _
                ByVal pContractStore As IContractStore, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pOrderManager As OrderManager, _
                ByVal pScopeName As String, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory)
Set mContractStore = pContractStore
Set mMarketDataManager = pMarketDataManager
Set mOrderManager = pOrderManager
mScopeName = pScopeName
Set mOrderSubmitterFactory = pOrderSubmitterFactory
End Sub

Friend Function Item(ByVal pGroupName As String) As GroupResources
Const ProcName As String = "Item"
On Error GoTo Err

Set Item = mDict.Item(UCase$(pGroupName))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
On Error GoTo Err

Set NewEnum = mDict.Enumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Remove(ByVal pGroupName As String)
Const ProcName As String = "Remove"
On Error GoTo Err

mDict.Remove UCase$(pGroupName)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function TryItem( _
                ByVal pGroupName As String, _
                ByRef pItem As GroupResources) As Boolean
Const ProcName As String = "TryItem"
On Error GoTo Err

TryItem = mDict.TryItem(UCase$(pGroupName), pItem)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

