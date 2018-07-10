VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarFetchListener"
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

Implements IBarFetchListener

'@================================================================================
' Events
'@================================================================================

Event FetchCancelled(ByVal pCookie As Variant)

Event FetchCompleted(ByVal pCookie As Variant)

Event FetchDataRetrieved(ByVal pCookie As Variant)

Event FetchFailed(ByVal pCookie As Variant, ByVal pErrorCode As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)

Event FetchStarted(ByVal pCookie As Variant)

Event NotifyBar(ByVal pCookie As Variant, ByVal pBar As Bar)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "BarFetchListener"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IBarFetchListener Interface Members
'@================================================================================

Private Sub IBarFetchListener_FetchCancelled(ByVal pCookie As Variant)
RaiseEvent FetchCancelled(pCookie)
End Sub

Private Sub IBarFetchListener_FetchCompleted(ByVal pCookie As Variant)
RaiseEvent FetchCompleted(pCookie)
End Sub

Private Sub IBarFetchListener_FetchDataRetrieved(ByVal pCookie As Variant)
RaiseEvent FetchDataRetrieved(pCookie)
End Sub

Private Sub IBarFetchListener_FetchFailed(ByVal pCookie As Variant, ByVal pErrorCode As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
RaiseEvent FetchFailed(pCookie, pErrorCode, pErrorMessage, pErrorSource)
End Sub

Private Sub IBarFetchListener_FetchStarted(ByVal pCookie As Variant)
RaiseEvent FetchStarted(pCookie)
End Sub

Private Sub IBarFetchListener_NotifyBar(ByVal pCookie As Variant, ByVal pBar As Bar)
RaiseEvent NotifyBar(pCookie, pBar)
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




