VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RealtimeBarRequester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "RealtimeBarRequester"

'@================================================================================
' Member variables
'@================================================================================

Private mRequesterReference                         As WeakReference

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

Friend Sub Initialise(ByVal pRequester As IMarketDataConsumer)
Set mRequesterReference = CreateWeakReference(pRequester)
End Sub

Friend Sub NotifyError(ByVal pRequesterId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "NotifyError"
On Error GoTo Err

getRequester.NotifyError pRequesterId, pErrorCode, pErrorMsg

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyRealtimeBar(ByVal pRequesterId As Long, pBar As TwsRealtimeBar)
Const ProcName As String = "NotifyRealtimeBar"
On Error GoTo Err

With pBar
    notifyPrice pRequesterId, .OpenValue
    If .LowValue < .CloseValue Then
        notifyPrice pRequesterId, .LowValue
        notifyPrice pRequesterId, .HighValue
    Else
        notifyPrice pRequesterId, .HighValue
        notifyPrice pRequesterId, .LowValue
    End If
    notifyPrice pRequesterId, .CloseValue
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getRequester() As IMarketDataConsumer
Set getRequester = mRequesterReference.Target
End Function

Private Sub notifyPrice(ByVal pRequesterId As Long, ByVal pPrice As Double)
Const ProcName As String = "notifyPrice"
On Error GoTo Err

Static sPrevNotifiedPrice As Double

If pPrice = sPrevNotifiedPrice Then Exit Sub

sPrevNotifiedPrice = pPrice
getRequester.NotifyTickPrice pRequesterId, TwsTickLast, pPrice, 0, False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

