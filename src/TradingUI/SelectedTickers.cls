VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectedTickers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

Event SelectionChanged()

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "SelectedTickers"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTickers                        As New EnumerableCollection

Private mChanging                       As Boolean
Private mChanged                        As Boolean

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

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add(ByVal pTicker As IMarketDataSource)
Const ProcName As String = "Add"
On Error GoTo Err

mTickers.Add pTicker
If Not mChanging Then
    RaiseEvent SelectionChanged
Else
    mChanged = True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub BeginChange()
If mChanging Then Exit Sub
mChanging = True
mChanged = False
End Sub

Friend Sub EndChange()
mChanging = False
If mChanged Then
    mChanged = False
    RaiseEvent SelectionChanged
End If
End Sub

Friend Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

If mTickers.Count = 0 Then Exit Sub
mTickers.Clear
If Not mChanging Then
    RaiseEvent SelectionChanged
Else
    mChanged = True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Contains( _
                ByVal pTicker As IMarketDataSource) As Boolean
Const ProcName As String = "Contains"
On Error GoTo Err

Contains = mTickers.Contains(pTicker)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Count() As Long
Const ProcName As String = "Count"
On Error GoTo Err

Count = mTickers.Count

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Item(ByVal pIndex As Variant) As IMarketDataSource
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "item"
On Error GoTo Err

Set Item = mTickers.Item(pIndex)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mTickers.Enumerator
End Function

Friend Sub Remove(ByVal pTicker As IMarketDataSource)
Const ProcName As String = "Remove"
On Error GoTo Err

If mTickers.Contains(pTicker) Then mTickers.Remove pTicker
If Not mChanging Then
    RaiseEvent SelectionChanged
Else
    mChanged = True
End If

Exit Sub

Err:
End Sub

'@================================================================================
' Helper Functions
'@================================================================================





