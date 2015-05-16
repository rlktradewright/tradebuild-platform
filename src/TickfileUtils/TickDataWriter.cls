VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickDataWriter"
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

Implements IGenericTickListener

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

Private Const ModuleName                            As String = "TickDataWriter"

Private mTickWriter                                 As ITickfileWriter

Private mSelfRef                                    As Object

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)
Set mSelfRef = Nothing
End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

mTickWriter.WriteTick ev.Tick

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

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

mTickWriter.Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTickSource As IGenericTickSource, _
                ByVal pContractFuture As IFuture, _
                ByVal pOutputMonitor As ITickfileOutputMonitor, _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pFormatIdentifier As String, _
                ByVal pLocation As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument Not pTickSource Is Nothing, "pTickSource is Nothing"
AssertArgument Not pTickfileStore Is Nothing, "pTickfileStore is Nothing"

Set mTickWriter = gCreateBufferedTickfileWriter(pTickfileStore, pOutputMonitor, pContractFuture, pFormatIdentifier, pLocation)

pTickSource.AddGenericTickListener Me

Set mSelfRef = Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




