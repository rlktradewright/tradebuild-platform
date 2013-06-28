VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickStream"
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

Implements ITickStream

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

Private Const ModuleName                            As String = "TickStream"

'@================================================================================
' Member variables
'@================================================================================

Private mListeners                                  As New EnumerableCollection

Private mStreamId                                   As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITickStream Interface Members
'@================================================================================

Private Sub ITickStream_AddGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "ITickStream_AddGenericTickListener"
On Error GoTo Err

AddGenericTickListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickStream_Finish()
Const ProcName As String = "ITickStream_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickStream_RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "ITickStream_RemoveGenericTickListener"
On Error GoTo Err

RemoveGenericTickListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get ITickStream_StreamId() As Long
ITickStream_StreamId = StreamId
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get StreamId() As Long
StreamId = mStreamId
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "AddGenericTickListener"
On Error GoTo Err

mListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()

End Sub

Friend Sub Initialise( _
                ByVal pStreamId As Long, _
                ByVal pListener As IGenericTickListener)
Const ProcName As String = "Initialise"
On Error GoTo Err

mStreamId = pStreamId
mListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTick(ByRef pTick As GenericTick)
Const ProcName As String = "NotifyTick"
On Error GoTo Err

Dim ev As GenericTickEventData

ev.StreamId = StreamId
ev.Tick = pTick

Set ev.Source = Me

Dim lListener As IGenericTickListener
For Each lListener In mListeners
    lListener.NotifyTick ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "RemoveGenericTickListener"
On Error GoTo Err

mListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================




