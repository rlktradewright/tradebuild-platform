VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPersistenceDataGroup"
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

Implements Enumerable

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

Private Const ModuleName                            As String = "OrderPersistenceDataGroup"

'@================================================================================
' Member variables
'@================================================================================

Private mOrderPersistenceItems                      As EnumerableCollection
Private mOrderPersistenceItemsEn                    As Enumerator

Private mPositionManagerItem                        As OrderPersistenceItem

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mOrderPersistenceItems = New EnumerableCollection

Exit Sub

Err:
UnhandledErrorHandler.notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get Enumerable_Enumerator() As TWUtilities30.Enumerator
Const ProcName As String = "Enumerable_Enumerator"
On Error GoTo Err

Set Enumerable_Enumerator = mOrderPersistenceItems.Enumerator

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Function Enumerable_NextItem(Item As Variant, cookie As Variant) As Boolean
Const ProcName As String = "Enumerable_NextItem"
On Error GoTo Err

If cookie = -1 Then Exit Function

Enumerable_NextItem = mOrderPersistenceItemsEn.MoveNext
If Enumerable_NextItem Then
    Set Item = mOrderPersistenceItemsEn.current
Else
    cookie = -1
    If Not mPositionManagerItem Is Nothing Then
        Set Item = mPositionManagerItem
        Enumerable_NextItem = True
    End If
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub Enumerable_Remove(cookie As Variant)
Dim lEnumable As Enumerable
Const ProcName As String = "Enumerable_Remove"
On Error GoTo Err

Set lEnumable = mOrderPersistenceItems
lEnumable.Remove cookie

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function Enumerable_Skip(ByVal numToSkip As Long, cookie As Variant) As Boolean
Dim lEnumable As Enumerable
Const ProcName As String = "Enumerable_Skip"
On Error GoTo Err

Set lEnumable = mOrderPersistenceItems
Enumerable_Skip = lEnumable.Skip(numToSkip, cookie)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Get Enumerable_UpdateSequence() As Long
Dim lEnumable As Enumerable
Const ProcName As String = "Enumerable_UpdateSequence"
On Error GoTo Err

Set lEnumable = mOrderPersistenceItems
Enumerable_UpdateSequence = lEnumable.UpdateSequence

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Enumerator() As TWUtilities30.Enumerator
Const ProcName As String = "Enumerator"

On Error GoTo Err

Set Enumerator = CreateEnumerator(Me)
Set mOrderPersistenceItemsEn = mOrderPersistenceItems.Enumerator

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add(ByRef pItem As OrderPersistenceItem)
Const ProcName As String = "Add"
On Error GoTo Err
If TypeOf pItem.Data Is PositionManagerSummary Then
    'we only want to keep the latest of these
    Set mPositionManagerItem = pItem
Else
    mOrderPersistenceItems.Add pItem
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


