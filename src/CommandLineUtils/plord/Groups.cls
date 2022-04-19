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

Private mGroupsResources                                       As SortedDictionary

Private mContractStore                              As IContractStore

Private mMarketDataManager                          As IMarketDataManager
Private mOrderManager                               As OrderManager
Private mScopeName                                  As String
Private mOrderSubmitterFactory                      As IOrderSubmitterFactory
Private mMoneyManager                               As MoneyManager

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mGroupsResources = CreateSortedDictionary(KeyTypeString)
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
Set Enumerator = mGroupsResources.Enumerator
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Add(ByVal pGroupName As String) As GroupResources
Const ProcName As String = "Add"
On Error GoTo Err

Dim lResources As GroupResources

If mGroupsResources.TryItem(UCase$(pGroupName), lResources) Then
    Set Add = lResources
    Exit Function
End If

Set lResources = New GroupResources
lResources.Initialise pGroupName, _
                        mContractStore, _
                        mMarketDataManager, _
                        mOrderManager, _
                        mScopeName, _
                        mOrderSubmitterFactory, _
                        mMoneyManager

mGroupsResources.Add lResources, UCase$(pGroupName)

Set Add = lResources

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Clear()
mGroupsResources.Clear
End Sub

Friend Function Contains(ByVal pGroupName As String) As Boolean
Contains = mGroupsResources.Contains(UCase$(pGroupName))
End Function

Friend Function Count() As Long
Count = mGroupsResources.Count
End Function

Friend Sub DiscardPendingOrdersForAllGroups()
Const ProcName As String = "DiscardPendingOrdersForAllGroups"
On Error GoTo Err

Dim lResources As GroupResources
For Each lResources In mGroupsResources
    doDiscardPendingOrdersForGroup lResources
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub DiscardPendingOrdersForGroup( _
                ByVal pGroupName As String)
Const ProcName As String = "DiscardPendingOrdersForGroup"
On Error GoTo Err

Dim lResources As GroupResources
Set lResources = mGroupsResources.Item(pGroupName)

doDiscardPendingOrdersForGroup lResources

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractStore As IContractStore, _
                ByVal pMarketDataManager As RealTimeDataManager, _
                ByVal pOrderManager As OrderManager, _
                ByVal pScopeName As String, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory, _
                ByVal pMoneyManager As MoneyManager)
Set mContractStore = pContractStore
Set mMarketDataManager = pMarketDataManager
Set mOrderManager = pOrderManager
mScopeName = pScopeName
Set mOrderSubmitterFactory = pOrderSubmitterFactory
Set mMoneyManager = pMoneyManager
End Sub

Friend Function Item(ByVal pGroupName As String) As GroupResources
Const ProcName As String = "Item"
On Error GoTo Err

Set Item = mGroupsResources.Item(UCase$(pGroupName))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
On Error GoTo Err

Set NewEnum = mGroupsResources.Enumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Remove(ByVal pGroupName As String)
Const ProcName As String = "Remove"
On Error GoTo Err

mGroupsResources.Remove UCase$(pGroupName)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function TryItem( _
                ByVal pGroupName As String, _
                ByRef pItem As GroupResources) As Boolean
Const ProcName As String = "TryItem"
On Error GoTo Err

TryItem = mGroupsResources.TryItem(UCase$(pGroupName), pItem)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub doDiscardPendingOrdersForGroup( _
                ByVal pResources As GroupResources)
Const ProcName As String = "doDiscardPendingOrdersForGroup"
On Error GoTo Err

Dim lCP As ContractProcessor
For Each lCP In pResources.ContractProcessors
    Dim lEnum As Enumerator: Set lEnum = lCP.PendingCommandsEnumerator
    If lEnum.MoveNext Then
        gWriteLineToConsole "Discarding commands for " & lCP.ContractName
        Do
            Dim lEntry As DictionaryEntry: Set lEntry = lEnum.Current
            Dim p As PendingCommand: Set p = lEntry.Data
            gWriteLineToConsole "    " & p.Command.name & " " & p.Params
            lEnum.Remove
        Loop While lEnum.MoveNext
    End If
Next

Dim lOP As OrderPlacer
For Each lOP In pResources.OrderPlacers
    lOP.Cancel "Closeout"
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



