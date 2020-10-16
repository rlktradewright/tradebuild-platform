VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsConnectionListener"
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

Implements ITwsConnectionStateListener

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

Private Const ModuleName                            As String = "TwsConnectionListener"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITwsConnectionStateListener Interface Members
'@================================================================================

Private Sub ITwsConnectionStateListener_NotifyAPIConnectionStateChange( _
                ByVal pSource As Object, _
                ByVal pState As ApiConnectionStates, _
                ByVal pMessage As String)
Dim lTimestamp As String
lTimestamp = FormatTimestamp(GetTimestamp, TimestampDateAndTimeISO8601) & "  "

Select Case pState
Case ApiConnNotConnected
    gWriteLineToConsole lTimestamp & "Disconnected from TWS: " & pMessage, True
Case ApiConnConnecting
    gWriteLineToConsole lTimestamp & "Connecting to TWS: " & pMessage, True
Case ApiConnConnected
    gWriteLineToConsole lTimestamp & "Connected to TWS: " & pMessage, True
Case ApiConnFailed
    gWriteLineToConsole lTimestamp & "Failed connecting to TWS: " & pMessage, True
End Select
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionClosed( _
                ByVal pSource As Object)

End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionRecovered( _
                ByVal pSource As Object, _
                ByVal pDataLost As Boolean)

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

'@================================================================================
' Helper Functions
'@================================================================================




