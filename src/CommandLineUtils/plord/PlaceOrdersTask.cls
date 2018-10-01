VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlaceOrdersTask"
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

Implements ITask

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

Private Const ModuleName                            As String = "PlaceOrdersTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private mProcessorsEn                               As Enumerator

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

Static sProcessor As Processor
Static sBracketOrdersEn As Enumerator


If sProcessor Is Nothing Then
    If Not mProcessorsEn.MoveNext Then
        gInputPaused = False
        mTaskContext.Finish Empty, False
        gInputPaused = False
        Exit Sub
    Else
        Set sProcessor = mProcessorsEn.Current
        Set sBracketOrdersEn = sProcessor.BracketOrders.Enumerator
    End If
End If

If sBracketOrdersEn.MoveNext Then
    sProcessor.PlaceOrder (sBracketOrdersEn.Current)
    mTaskContext.Suspend 25
Else
    Set sProcessor = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ByVal pProcessors As EnumerableCollection)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mProcessorsEn = pProcessors.Enumerator

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




