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

Private mTickers As Collection

Private mChanging As Boolean
Private mChanged As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mTickers = New Collection

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
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

Friend Sub Add(ByVal pTicker As Ticker)
On Error GoTo Err
mTickers.Add pTicker, CStr(ObjPtr(pTicker))
If Not mChanging Then
    RaiseEvent SelectionChanged
Else
    mChanged = True
End If
Exit Sub

Err:
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
Set mTickers = New Collection
If Not mChanging Then
    RaiseEvent SelectionChanged
Else
    mChanged = True
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function Contains( _
                ByVal pTicker As Ticker) As Boolean
On Error GoTo Err
mTickers.item CStr(ObjPtr(pTicker))
Contains = True
Exit Function

Err:
End Function

Public Function Count() As Long
Const ProcName As String = "Count"
Dim failpoint As String
On Error GoTo Err

Count = mTickers.Count

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Public Function item(ByVal index As Variant) As Ticker
Attribute item.VB_UserMemId = 0
Const ProcName As String = "item"
Dim failpoint As String
On Error GoTo Err

Set item = mTickers.item(index)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mTickers.[_NewEnum]
End Function

Friend Sub Remove(ByVal pTicker As Ticker)
Const ProcName As String = "Remove"
On Error GoTo Err

mTickers.Remove CStr(ObjPtr(pTicker))
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





