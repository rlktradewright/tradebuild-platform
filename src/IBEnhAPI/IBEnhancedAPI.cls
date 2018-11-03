VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBEnhancedAPI"
Attribute VB_GlobalNameSpace = True
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
' Constants
'@================================================================================

Private Const ModuleName                            As String = "IBEnhancedAPI"

'@================================================================================
' Enums
'@================================================================================

Public Enum ApiConnectionStates
    ApiConnNotConnected
    ApiConnConnecting
    ApiConnConnected
    ApiConnFailed
End Enum

' These are defined indentically to those in IBAPI. This is so that users of this
' library do not also need a reference to the IBAPI project.
Public Enum TwsLogLevels
    TwsLogLevelSystem = 1
    TwsLogLevelError
    TwsLogLevelWarning
    TwsLogLevelInformation
    TwsLogLevelDetail
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type ContractCacheEntry
    ContractDetails         As TwsContractDetails
    CacheDate               As Date
End Type

Public Type OrderIdMapEntry
    Order                   As IOrder
    OrderSubmitter          As OrderSubmitter
    Cancelled               As Boolean
    ParentId                As Long
End Type

'@================================================================================
' Member variables
'@================================================================================

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
                Optional ByVal pSessionID As String, _
                Optional ByVal pConnectionRetryIntervalSecs As Long = 60, _
                Optional ByVal pLogTwsMessages As Boolean, _
                Optional ByVal pDeferConnection As Boolean, _
                Optional ByVal pConnectionStateListener As ITwsConnectionStateListener, _
                Optional ByVal pProgramErrorHandler As IProgramErrorListener, _
                Optional ByVal pApiErrorListener As IErrorListener, _
                Optional ByVal pApiNotificationListener As INotificationListener) As Client
Const ProcName As String = "GetClient"
On Error GoTo Err

If pSessionID = "" Then pSessionID = GenerateGUIDString

Set GetClient = gGetClient(pSessionID, pServer, pPort, pClientId, pConnectionRetryIntervalSecs, pLogTwsMessages, pDeferConnection, pConnectionStateListener, pProgramErrorHandler, pApiErrorListener, pApiNotificationListener)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================





