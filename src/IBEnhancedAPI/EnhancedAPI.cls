VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnhancedIBAPI"
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

Private Const ModuleName                            As String = "EnhancedIBAPI"

'@================================================================================
' Member variables
'@================================================================================

Private mClientCollection                           As New EnumerableCollection

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

Public Function GetClient( _
                ByVal pServer As String, _
                ByVal pPort As Long, _
                ByVal pClientId As Long, _
                Optional ByVal pConnectionRetryIntervalSecs As Long = 10, _
                Optional ByVal pDeferConnection As Boolean = False, _
                Optional ByVal pConnectionStateListener As ITwsConnectionStateListener, _
                Optional ByVal pProgramErrorHandler As ProgramErrorListener, _
                Optional ByVal pApiErrorListener As ErrorListener, _
                Optional ByVal pApiNotificationListener As NotificationListener) As Client
Const ProcName As String = "GetClient"
On Error GoTo Err

If pServer = "" Then pServer = "127.0.0.1"

Dim lKey As String
lKey = generateTwsKey(pServer, pPort, pClientId)

If Not mClientCollection.Contains(lKey) Then
    Set GetClient = New Client
    mClientCollection.Add GetClient, lKey
    
    GetClient.Initialise Me, pServer, pPort, pClientId, pConnectionRetryIntervalSecs, pDeferConnection, pConnectionStateListener, pProgramErrorHandler, pApiErrorListener, pApiNotificationListener
Else
    Set GetClient = mClientCollection(lKey)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ReleaseClient(ByVal pClient As Client)
Const ProcName As String = "ReleaseClient"
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







