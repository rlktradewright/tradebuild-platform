VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketOrderListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements IBracketOrderErrorListener
Implements IBracketOrderMsgListener
Implements IChangeListener

'@================================================================================
' Events
'@================================================================================

Event BracketOrderAdded(ev As StateChangeEventData)
Event NotifyBracketOrderError(ev As BracketOrderErrorEventData)
Event NotifyBracketOrderMessage(ev As BracketOrderMessageEventData)
Event Change(ev As ChangeEventData)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "BracketOrderListener"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IBracketOrderErrorListener Interface Members
'@================================================================================

Private Sub IBracketOrderErrorListener_NotifyBracketOrderError(ev As BracketOrderErrorEventData)
Const ProcName As String = "IBracketOrderErrorListener_NotifyBracketOrderError"
On Error GoTo Err

RaiseEvent NotifyBracketOrderError(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IBracketOrderMsgListener Interface Members
'@================================================================================

Private Sub IBracketOrderMsgListener_NotifyBracketOrderMessage(ev As BracketOrderMessageEventData)
Const ProcName As String = "IBracketOrderMsgListener_NotifyBracketOrderMessage"
On Error GoTo Err

RaiseEvent NotifyBracketOrderMessage(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
On Error GoTo Err

RaiseEvent Change(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub Add(ByVal pBO As IBracketOrder)
pBO.AddBracketOrderErrorListener Me
pBO.AddBracketOrderMessageListener Me
pBO.AddChangeListener Me

Dim ev As StateChangeEventData
Set ev.Source = pBO
ev.State = pBO.State
RaiseEvent BracketOrderAdded(ev)
End Sub

Public Sub Remove(ByVal pBO As IBracketOrder)
pBO.RemoveBracketOrderErrorListener Me
pBO.RemoveBracketOrderMessageListener Me
pBO.RemoveChangeListener Me
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




