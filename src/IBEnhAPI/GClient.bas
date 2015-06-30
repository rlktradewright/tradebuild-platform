Attribute VB_Name = "GClient"
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

Private Const ModuleName                            As String = "GClient"

'@================================================================================
' Member variables
'@================================================================================

Private mClientCollection                           As New EnumerableCollection

''@================================================================================
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

Public Function gGetClient( _
                ByVal pSessionID As String, _
                ByVal pServer As String, _
                ByVal pPort As Long, _
                ByVal pClientId As Long, _
                ByVal pConnectionRetryIntervalSecs As Long, _
                ByVal pDeferConnection As Boolean, _
                ByVal pConnectionStateListener As ITwsConnectionStateListener, _
                ByVal pProgramErrorHandler As IProgramErrorListener, _
                ByVal pApiErrorListener As IErrorListener, _
                ByVal pApiNotificationListener As INotificationListener) As Client
Const ProcName As String = "gGetClient"
On Error GoTo Err

Dim lKey As String

If pServer = "" Then pServer = "127.0.0.1"

lKey = generateTwsKey(pServer, pPort, pClientId)

If Not mClientCollection.Contains(lKey) Then
    Set gGetClient = New Client
    mClientCollection.Add gGetClient, lKey
    
    gGetClient.Initialise pSessionID, pServer, pPort, pClientId, pConnectionRetryIntervalSecs, pDeferConnection, pConnectionStateListener, pProgramErrorHandler, pApiErrorListener, pApiNotificationListener
Else
    Set gGetClient = mClientCollection(lKey)
    Assert gGetClient.SessionID = pSessionID, "Client already started in another session"
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gReleaseClient(ByVal pClient As Client)
Const ProcName As String = "gReleaseClient"
On Error GoTo Err

mClientCollection.Remove generateTwsKey(pClient.Server, pClient.Port, pClient.ClientId)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function generateTwsKey( _
                ByVal pServer As String, _
                ByVal pPort As Long, _
                ByVal pClientId As Long) As String
generateTwsKey = pServer & vbNullChar & pPort & vbNullChar & pClientId
End Function




