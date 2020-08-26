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

Public Enum ApiMessageLoggingOptions
    ApiMessageLoggingOptionNone = IBAPIV100.ApiMessageLoggingOptionNone
    ApiMessageLoggingOptionDefault = IBAPIV100.ApiMessageLoggingOptionDefault
    ApiMessageLoggingOptionAlways = IBAPIV100.ApiMessageLoggingOptionAlways
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
    ContractsFutureBuilder          As FutureBuilder
    Contracts                       As IContracts
    TwsContracts                    As EnumerableCollection
    CacheDate                       As Date
End Type

Public Type OrderIdMapEntry
    Order                   As Iorder
    OrderSubmitter          As OrderSubmitter
    Cancelled               As Boolean
    ParentId                As Long
End Type

Public Type PriceIncrementRulesetEntry
    RulesetFutureBuilder    As FutureBuilder
    Ruleset                 As PriceIncrementRuleset
    CacheDate               As Date
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
                Optional ByVal pLogApiMessages As ApiMessageLoggingOptions = ApiMessageLoggingOptionDefault, _
                Optional ByVal pLogRawApiMessages As ApiMessageLoggingOptions = ApiMessageLoggingOptionDefault, _
                Optional ByVal pLogApiMessageStats As Boolean = False, _
                Optional ByVal pDeferConnection As Boolean, _
                Optional ByVal pConnectionStateListener As ITwsConnectionStateListener, _
                Optional ByVal pProgramErrorHandler As IProgramErrorListener, _
                Optional ByVal pApiErrorListener As IErrorListener, _
                Optional ByVal pApiNotificationListener As INotificationListener) As Client
Const ProcName As String = "GetClient"
On Error GoTo Err

If pSessionID = "" Then pSessionID = GenerateGUIDString

Set GetClient = gGetClient(pSessionID, _
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
                            pApiNotificationListener)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================





