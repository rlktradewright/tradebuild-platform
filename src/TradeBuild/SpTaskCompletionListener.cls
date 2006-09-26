VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpTaskCompletionListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TaskCompletionListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mSpTaskCompletionListener As TradeBuildSP.ITaskCompletionListener

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' TaskCompletionListener Interface Members
'================================================================================

Private Sub TaskCompletionListener_taskCompleted( _
                ev As TaskCompletionEvent)
Dim spEv As TradeBuildSP.TaskCompletionEvent
Dim spRunn As SpTask

If TypeOf ev.source Is SpTask Then
    Set spRunn = ev.source
    Set spEv.source = spRunn.SpTaskObject
Else
    Set spEv.source = ev.source
End If
spEv.message = ev.message
spEv.resultCode = ev.resultCode
mSpTaskCompletionListener.taskCompleted spEv
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let SpTaskCompletionListener( _
                ByVal value As TradeBuildSP.ITaskCompletionListener)
Set mSpTaskCompletionListener = value
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

