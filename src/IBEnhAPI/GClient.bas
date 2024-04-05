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

Private mContractCache                              As New ContractCache

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

Public Property Get gContractCache() As ContractCache
Set gContractCache = mContractCache
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function GetClient( _
                ByVal pSessionID As String, _
                ByVal pServer As String, _
                ByVal pPort As Long, _
                ByVal pClientId As Long, _
                ByVal pConnectionRetryIntervalSecs As Long, _
                ByVal pLogApiMessages As ApiMessageLoggingOptions, _
                ByVal pLogRawApiMessages As ApiMessageLoggingOptions, _
                ByVal pLogApiMessageStats As Boolean, _
                ByVal pDeferConnection As Boolean, _
                ByVal pConnectionStateListener As ITwsConnectionStateListener, _
                ByVal pProgramErrorHandler As IProgramErrorListener, _
                ByVal pApiErrorListener As IErrorListener, _
                ByVal pApiNotificationListener As INotificationListener) As Client
Const ProcName As String = "GetClient"
On Error GoTo Err

Dim lKey As String

If pServer = "" Then pServer = "127.0.0.1"

lKey = generateTwsKey(pServer, pPort, pClientId)

If Not mClientCollection.Contains(lKey) Then
    Set GetClient = New Client
    mClientCollection.Add GetClient, lKey
    
    GetClient.Initialise pSessionID, _
                        pServer, _
                        pPort, _
                        pClientId, _
                        pConnectionRetryIntervalSecs, _
                        pLogApiMessages, _
                        pLogRawApiMessages, _
                        pLogApiMessageStats, _
                        pDeferConnection, _
                        pConnectionStateListener, _
                        pProgramErrorHandler, _
                        pApiErrorListener, _
                        pApiNotificationListener
Else
    Set GetClient = mClientCollection(lKey)
    Assert GetClient.SessionID = pSessionID, "Client already started in another session"
End If

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gReleaseClient(ByVal pClient As Client)
Const ProcName As String = "gReleaseClient"
On Error GoTo Err

mClientCollection.Remove generateTwsKey(pClient.Server, pClient.Port, pClient.ClientId)

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
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




